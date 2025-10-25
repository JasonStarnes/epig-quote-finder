
import streamlit as st
import pandas as pd
import math
import os
from typing import Dict, Any, Tuple, List, Set

st.set_page_config(page_title="EPIG Quote Finder", page_icon="🚚", layout="wide")

# ---- Global CSS: fix Reset button at bottom of sidebar & style it red
st.markdown(
    """
    <style>
    /* Make sidebar a flex column so we can push the reset zone to the bottom */
    div[data-testid="stSidebar"] > div {
        display: flex;
        flex-direction: column;
        height: 100%;
    }
    /* The content container grows; the reset container sits at the bottom */
    .sidebar-flex-grow {
        flex: 1 1 auto;
    }
    .sidebar-reset-zone {
        margin-top: auto;
        padding-top: 0.75rem;
        border-top: 1px solid rgba(0,0,0,0.2);
    }
    /* Force the Reset button to be red */
    .epig-reset-btn button {
        background-color: #ff4b4b !important;
        color: #ffffff !important;
        border: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("🚚 EPIG Quote Finder — v15 (Fixed Reset Button)")

@st.cache_data(show_spinner=False)
def load_rules():
    if not os.path.exists("carrier_rules_master.xlsx"):
        st.error("`carrier_rules_master.xlsx` not found. Place it next to this file and restart.")
        st.stop()
    xls = pd.ExcelFile("carrier_rules_master.xlsx", engine="openpyxl")
    return {
        "agency": pd.read_excel(xls, "Agency_Appetite"),
        "inputs": pd.read_excel(xls, "Agent_Input_Fields"),
        "rules": pd.read_excel(xls, "Carrier_Rules"),
        "fields": pd.read_excel(xls, "Field_Dictionary"),
    }

def norm(x):
    if x is None or (isinstance(x,float) and math.isnan(x)): return ""
    return str(x).strip().lower()

def to_list(val):
    if val is None or (isinstance(val, float) and math.isnan(val)): return []
    if isinstance(val, list): return [norm(x) for x in val]
    return [norm(p) for p in str(val).split(",") if norm(p)!=""]

def try_float(s):
    try:
        return float(str(s).replace(",","").replace("$","").strip())
    except Exception:
        return float("nan")

def between_val(value):
    parts = str(value).replace(" ", "").split("-")
    if len(parts)==2:
        return try_float(parts[0]), try_float(parts[1])
    return float("nan"), float("nan")

EXCLUSION_KEYS = {
    "vehicle_type","operation","cargo_type","trailer_type",
    "safer_cargo","safer_operation",
    "ineligible_operations","excluded_operations"
}

BASE_REQUIRED_KEYS = {
    "state_code","eld_installed","dashcam_eld_installed",
    "power_units","fleet_size","driver_experience_years",
    "years_in_business","license_required",
    "mvr_age_days","fmcsa_status","mcs90_required","vehicles_listed_vs_operated",
    "operation_type","operation_class","radius_min_for_new_federal_filing"
}

SEMI_ONLY = {"bulldog national rrg","southwind rrg","berkley small business"}
LENIENT_EXP = {"progressive","geico"}
ALL_VEHICLES = {"progressive","geico"}  # accept all vehicle types (subject to state limitations elsewhere)
AGE_DECLINE = {"cover whale","nirvana","bulldog national rrg","southwind rrg"}
NIRVANA_STATES = {"az","ga","il","in","ia","mi","mo","nv","nc","sc","oh","pa","tn","wi"}  # lowercase
GEICO_STATES = {
    "or","id","nv","az","wy","co","nm","nd","sd","ne","ks","ok","tx","ia","mo","ar","wi","il","tn","ms",
    "in","al","oh","wv","sc","fl","md","ga","va","pa","mn"
}

def eval_rule(op, rule_val, input_val):
    o = norm(op)
    if o in ["==","="]:   return norm(input_val) == norm(rule_val)
    if o == "!=":         return norm(input_val) != norm(rule_val)
    if o in [">",">=","<","<="]:
        try:
            a = float(str(input_val).replace(",","").replace("$","").strip())
            b = float(str(rule_val).replace(",","").replace("$","").strip())
        except Exception:
            return False
        return (a>b) if o==">" else (a>=b) if o==">=" else (a<b) if o=="<" else (a<=b) if o=="<=" else False
    if o == "in":
        vals = set(to_list(rule_val))
        if isinstance(input_val, list):
            return len(set([norm(x) for x in input_val]).intersection(vals))>0
        return norm(input_val) in vals
    if o == "not in":
        vals = set(to_list(rule_val))
        if isinstance(input_val, list):
            return len(set([norm(x) for x in input_val]).intersection(vals))==0
        return norm(input_val) not in vals
    if o == "between":
        lo,hi = between_val(rule_val)
        try:
            a = float(str(input_val).replace(",","").replace("$","").strip())
        except Exception:
            return False
        if math.isnan(lo) or math.isnan(hi): return False
        return lo <= a <= hi
    return True

def evaluate_carrier(carrier, rules_df, inputs):
    rows = rules_df[rules_df["Carrier"]==carrier].copy()
    carrier_norm = norm(carrier)

    portal = rows["Portal_URL"].dropna().astype(str).head(1).tolist()
    portal = portal[0] if portal else ""
    submission = rows["Submission_Type"].dropna().astype(str).head(1).tolist()
    submission = submission[0] if submission else ""
    req_fields = rows["Required_Fields"].dropna().astype(str).head(1).tolist()
    req_fields = req_fields[0] if req_fields else ""

    hard_fails, required_fails, soft_fails, matches = [], [], [], []
    score = 0.0

    # Semi-only carriers
    if carrier_norm in SEMI_ONLY and inputs.get("vehicle_type") != "semi":
        return {
            "carrier": carrier,
            "status": "Ineligible",
            "score": 0.0,
            "portal": portal,
            "submission": submission,
            "required_fields": req_fields,
            "matches": matches,
            "required_fails": required_fails,
            "soft_fails": soft_fails,
            "hard_fails": [f"Ineligible — {carrier} writes semi-trucks only (selected: {inputs.get('vehicle_type')})"],
        }

    # Progressive / GEICO lenient on exp/license
    required_keys = set(BASE_REQUIRED_KEYS)
    if carrier_norm in LENIENT_EXP:
        required_keys.discard("driver_experience_years")
        required_keys.discard("license_required")

    # Cover Whale: min 2 yrs experience
    if carrier_norm == "cover whale":
        try:
            exp_years = float(inputs.get("driver_experience_years", 0))
        except Exception:
            exp_years = 0.0
        if exp_years < 2:
            hard_fails.append("Ineligible — Cover Whale requires 2+ years of like-vehicle driving experience")

    # Age of equipment decline for specific carriers
    if norm(inputs.get("equipment_age_bucket","")) == "older_than_20" and carrier_norm in AGE_DECLINE:
        hard_fails.append("Ineligible — Vehicle/trailer older than 20 years (carrier requires ≤20 years old)")

    # Nirvana must be in approved state list
    if carrier_norm == "nirvana":
        state = norm(inputs.get("state_code",""))
        if state not in NIRVANA_STATES:
            hard_fails.append(f"Ineligible — Nirvana does not write in {state.upper()}")

    # Cover Whale auto-decline for Cargo Van
    if carrier_norm == "cover whale" and norm(inputs.get("vehicle_type","")) == "cargo_van":
        hard_fails.append("Ineligible — Cover Whale does not write Cargo Van policies")

    # Progressive & GEICO not in California
    if norm(inputs.get("state_code","")) == "ca" and carrier_norm in {"progressive","geico"}:
        hard_fails.append(f"Ineligible — {carrier} not available in California")

    # GEICO state whitelist
    if carrier_norm == "geico":
        st_code = norm(inputs.get("state_code",""))
        if st_code not in GEICO_STATES:
            hard_fails.append(f"Ineligible — GEICO does not write in {st_code.upper()}")

    for _,r in rows.iterrows():
        key = norm(r.get("Criteria_Key",""))
        op  = r.get("Operator","")
        val = r.get("Value","")
        wt  = 0 if pd.isna(r.get("Weight",0)) else float(r.get("Weight",0))
        excl_notes = r.get("Excluded_Notes","")
        is_exclusion = (key in EXCLUSION_KEYS) or (isinstance(excl_notes,str) and excl_notes.strip()!="")

        user_val = inputs.get(key, None)
        if key == "state_code":
            user_val = inputs.get("state_code")

        # Ignore vehicle_type checks for carriers that accept all vehicle types
        if key == "vehicle_type" and carrier_norm in ALL_VEHICLES:
            matches.append("vehicle type accepted (carrier writes all types)")
            score += wt
            continue

        # UIIA special handling
        if key == "uiia_required":
            if norm(user_val) == "true":
                hard_fails.append("Ineligible — UIIA agreement required (carrier excludes intermodal/UIIA risks)")
            else:
                matches.append("UIIA not required")
                score += wt
            continue

        if user_val is None and not is_exclusion:
            continue

        passed = eval_rule(op, val, user_val if user_val is not None else "")

        if is_exclusion:
            if passed:
                if key in {"operation","cargo_type","trailer_type","safer_cargo","safer_operation","ineligible_operations","excluded_operations"}:
                    hard_fails.append(f"Ineligible — excluded {key.replace('_',' ')}: {val}")
                else:
                    hard_fails.append(f"Ineligible — {key} conflicts with carrier rule: {val}")
            else:
                matches.append(f"avoided exclusion: {key}")
                score += wt
        elif key in required_keys:
            if passed:
                matches.append(f"{key} ok ({op} {val})")
                score += wt
            else:
                required_fails.append(f"{key} must satisfy ({op} {val})")
        else:
            if passed:
                matches.append(f"{key} ok ({op} {val})")
                score += wt
            else:
                soft_fails.append(f"{key} expected ({op} {val})")

    status = "Eligible"
    if hard_fails:
        status = "Ineligible"
    elif required_fails:
        status = "Conditional"

    return {
        "carrier": carrier,
        "status": status,
        "score": score,
        "portal": portal,
        "submission": submission,
        "required_fields": req_fields,
        "matches": matches,
        "required_fails": required_fails,
        "soft_fails": soft_fails,
        "hard_fails": hard_fails,
    }

# ---------------- UI ----------------
data = load_rules()
rules = data["rules"]

# We build the sidebar in two zones: growable content + fixed reset zone
with st.sidebar:
    grow = st.container()
    reset_zone = st.container()

with grow:
    # Add a marker class for flex grow via HTML (empty but useful for CSS)
    st.markdown('<div class="sidebar-flex-grow"></div>', unsafe_allow_html=True)

    st.header("Agent Intake")

    def extract_states_from_rules(rules_df) -> Set[str]:
        states: Set[str] = set()
        if "State_Code" in rules_df.columns:
            for v in rules_df["State_Code"].dropna().astype(str).tolist():
                for s in v.split(","):
                    s = s.strip().upper()
                    if len(s)==2 and s.isalpha():
                        states.add(s)
        if "Criteria_Key" in rules_df.columns and "Value" in rules_df.columns:
            subset = rules_df[rules_df["Criteria_Key"].astype(str).str.lower()=="state_code"]
            for v in subset["Value"].dropna().astype(str).tolist():
                for s in v.split(","):
                    s = s.strip().upper()
                    if len(s)==2 and s.isalpha():
                        states.add(s)
        return states

    state_values = sorted(list(extract_states_from_rules(rules)))
    US_STATES = ["AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA","ME",
                "MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA",
                "RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY","DC"]
    if not state_values:
        state_values = US_STATES

    state_code = st.selectbox("Risk State", options=state_values, index=state_values.index("TX") if "TX" in state_values else 0)
    radius = st.selectbox("Radius", ["local","regional","interstate","long_haul"])

    vehicle_display_to_key = {
        "Cargo Van": "cargo_van",
        "Hotshot": "hotshot_class2",
        "Box Truck": "box_truck",
        "Semi-truck": "semi",
    }
    vehicle_display = st.selectbox("Vehicle Type", list(vehicle_display_to_key.keys()))
    vehicle_type = vehicle_display_to_key[vehicle_display]

    # Age of Equipment
    age_label = st.selectbox("Age of Equipment", ["Newer than 20 years", "Older than 20 years"])
    equipment_age_bucket = "newer_than_20" if age_label.startswith("Newer") else "older_than_20"

    # Trailer choices (includes gooseneck)
    trailer_choices = ["dry_van", "flatbed", "reefer", "straight", "intermodal", "non_tanker", "heavy_haul", "gooseneck"]

    # Determine whether trailers apply
    trailers_allowed = vehicle_type in {"hotshot_class2", "semi"}

    if trailers_allowed:
        trailer_selected = st.multiselect("Trailer Type(s)", trailer_choices, default=["dry_van"])
        operation_class = (trailer_selected[0] if trailer_selected else "dry_van")
    else:
        trailer_selected = []
        # Default operation_class when no trailer: cargo van -> non_tanker; box truck -> straight
        operation_class = "non_tanker" if vehicle_type == "cargo_van" else ("straight" if vehicle_type == "box_truck" else "dry_van")
        st.caption("Trailer Type is only applicable for Hotshot or Semi-truck. Hidden for Cargo Van and Box Truck.")

    # Cargo commodities (long curated list)
    cargo_list = [
        "Air Conditioning Equipment","Air Freight (Mail, UPS, & FedEx)","Aircraft-- Engines","Aircraft-- Parts","Amazon Goods",
        "Appliances-- Major","Appliances-- Small","ATVs, Snowmobiles & Waverunners","Audio, Video & Studio Equipment",
        "Auto Parts & Accessories (not Tires)","Bakery Goods & Products","Batteries","Beverage-- Beer","Beverage-- Liquor",
        "Beverage-- Non-Alcoholic","Beverage-- Wine","Bicycles, Segways & Hoverboards","Books","Bottles-- Glass","Bottles-- Plastic",
        "Building Materials (non-flatbed)","Cable & Wire (not Copper/Fiber Optics)","Cable & Wire-- Fiber Optics","Camera Supplies & Accessories",
        "Cameras & Photographic Equipment","Candy & Confectionary","Canned Goods","Carpets & Rugs (other than Oriental Rugs)","Cell Phones",
        "Chemicals (not Red Label Placard)","Chemicals-- Red Label Placard","China, Ceramics, Glassware & Pottery",
        "Cigarettes, Cigars or Tobacco (mfgd.)","Clothing (other than Athletic, Designer Clothing)","Clothing-- Athletic",
        "Clothing-- Blue Jeans","Clothing-- Designer","Clothing-- T Shirts","Computer-- Desktop, Laptop & Network Systems",
        "Computer-- Games & Software","Computer-- Gaming Systems","Containerized Freight","Copper and Copper Products","Cosmetics (not Perfume)",
        "Cotton (baled cotton)","Dept Store Merchandise (Macy's, Nordstrom's)","Dept Store Merchandise (Wal-Mart, Target, etc.)",
        "DVD Players, MP3 Players, iPods & iPads","DVDs, Compact Discs & Tapes","Electrical Supplies (other than Copper)","Electronic-- Accessories",
        "Electronic-- Data Processing (not Computers)","Explosives","Fashion Accessories","Film","Fine Arts","Firearms and Ammunition",
        "Furniture (new from factory)","Furs","General Merchandise/Dry Freight","Glass","Groceries (other than produce)","Hardware",
        "Hazardous Materials (Placards)","Household Goods (Movers)","Jewelry & Watches","Lawn and Garden Equipment","Leather and Leather Goods",
        "Machinery (Light & Non-precision)","Medical Diagnostic Equipment","Medical Supplies (not Medicine or Drugs)","Memorabilia & Collectibles",
        "Metal Products (finished-- not precious)","Money & Securities","Motorcycles","Office Equipment","Office Supplies","Optical Goods",
        "Oriental Rugs","Paint","Paper, Paper Products & Printed Matter","Perfume","Pharmaceutical Products-- Over the Counter",
        "Pharmaceutical Products-- Prescription","Plants-- Flowers","Plants-- Trees & Shrubs (other than Temp Controlled)","Plastic Products",
        "Plumbing Supplies (other than Copper)","Precious Metals, Bullion & Alloys","Produce (other than refrigerated)","Purses, Handbags, Wallets & Belts",
        "Radioactive Material","Rubber Products (other than Tires)","Shoes (other than Designer & Athletic)","Shoes-- Designer & Athletic","Signs",
        "Silk & Silk Products","Soap Products, Household Cleaners","Sporting Goods & Equipment","Stereos & Radios","Televisions","Textiles (Cloth)",
        "Tires","Tobacco (unrefined)","Toiletries","Tools","Toys (Hobbies & Crafts)","Wood Products (other than Furniture)","Butter","Cheese",
        "Dairy Products","Eggs","Frozen Food (other than Meat & Seafood)","Ice Cream","Juice","Meat-- Boxed","Meat-- Hanging/Swinging","Milk",
        "Plants-- Cut Flowers","Plants-- Trees & Shrubs","Poultry","Produce","Reefer processed foods (other than frozen)",
        "Seafood-- Fresh (other than Canned Seafood)","Seafood-- Frozen (other than Canned Seafood)","Aggregate","Agricultural Products","Asphalt (Liquid)",
        "Bulk Hazardous Materials (Placard)","Cement","Coal","Copper and Copper Products (Bulk)","Cotton (other than Baled Cotton)","Feed (Animal/Livestock)",
        "Fertilizer (Nitrate)","Fertilizer (other than Nitrates)","Frac Sand","Fruits & Vegetables (fresh from field)","Garbage, Trash & Refuse",
        "Gas, Fuel & Petroleum Products","Grain","Gravel, Stone & Rock","Hay","Heating Oil","Liquid (Non-flammable)","Metals (raw, scrap, coils-- other than copper)",
        "Milk (Tanker)","Nuts - Peanuts, Almonds, Pecans, Pistachios, etc.","Oil (unrefined)","Salt","Sand","Tar","Top Soil, Dirt & Fill",
        "Wood Chips, Mulch & Garden Ties","Agricultural Equipment","Building Materials","Coiled Steel","Contractors Equipment",
        "Copper and Copper Products (Flatbed)","Flatbed Cable & Wire (not Copper/Fiber Optics)","Flatbed Cable & Wire - Fiber Optics","Landscaping Products",
        "Logs, Timber & Pulpwood (Logging)","Lumber, Pallets & Wood (processed)","Machinery-- Heavy > 10,000 lbs (flatbed)","Machinery-- Light < 10,000 lbs (flatbed)",
        "Metal (other than Copper)","Mobile Homes (no motor)","Oil Field - Pipe & Valves","Oil Field-- Heavy Equipment","Oversized/Overweight Loads",
        "Pipe (other than Copper)","Printing Presses","Rigging (any property requiring)","Solar Panels","Stone Slab/Products (Marble, etc.)",
        "Swimming Pools, Spas & Hot Tubs","Trailers (New for delivery)","Transformers","Turbines","Antique Automobiles/Vehicles","New Automobiles",
        "New Boats","Recreational Vehicles","Reposessed Vehicles","Towing","Used Automobiles","Used Boats","Wrecker Service","Cattle","Equine (other than Racing)",
        "Equine (racing)","Exotic Animals (Alpaca, Llama, etc.)","Goats","Poultry (Livestock)","Sheep","Swine",
    ]
    cargo_selected = st.multiselect("Cargo Commodities", cargo_list, default=[])

    power_units = st.number_input("Power Units (fleet size)", min_value=1, max_value=100, value=1, step=1)
    years_in_business = st.number_input("Years in Business", min_value=0, max_value=100, value=2, step=1)
    driver_experience_years = st.number_input("Min Driver Experience (years)", min_value=0, max_value=50, value=2, step=1)
    # MVR Age hidden; internal default set later
    license_required = st.selectbox("License Type", ["non_cdl","cdl"])
    eld_installed = st.selectbox("ELD / Dashcam Installed?", ["true","false"])
    operation_type = st.selectbox("Operation Type", ["for_hire","private"])
    uiia_required = st.selectbox("UIIA Agreement Required?", ["false","true"])

    # Optional Details
    with st.expander("Optional Details"):
        op_choices = ["dry_van","flatbed","reefer","straight","intermodal","non_tanker","heavy_haul","gooseneck"]
        operation_override = st.selectbox("Operation Class (override)", [""] + op_choices, index=0)
        if operation_override:
            operation_class = operation_override
        safer_cargo = st.text_input("SAFER Cargo Selection(s) (comma-separated)", "")
        safer_operation = st.text_input("SAFER Operation Selection(s) (comma-separated)", "")

    submitted = st.button("Find Carriers", use_container_width=True)

DEFAULT_MVR_AGE_DAYS = 30

# Collect inputs after sidebar creation
user_inputs = {
    "state_code": norm(state_code),
    "radius": norm(radius),
    "vehicle_type": norm(vehicle_type),
    "trailer_type": [norm(x) for x in (trailer_selected if trailers_allowed else [])],
    "cargo_type": [norm(x) for x in cargo_selected],
    "power_units": power_units,
    "years_in_business": years_in_business,
    "driver_experience_years": driver_experience_years,
    "mvr_age_days": DEFAULT_MVR_AGE_DAYS,
    "license_required": norm(license_required),
    "eld_installed": norm(eld_installed),
    "dashcam_eld_installed": norm(eld_installed),
    "operation_type": norm(operation_type),
    "uiia_required": norm(uiia_required),
    "operation_class": norm(operation_class),
    "operation": norm(operation_class),
    "equipment_age_bucket": norm(equipment_age_bucket),
    "safer_cargo": to_list(safer_cargo),
    "safer_operation": to_list(safer_operation),
}

# Bottom fixed reset zone
with reset_zone:
    st.markdown('<div class="sidebar-reset-zone"></div>', unsafe_allow_html=True)
    col = st.container()
    with col:
        st.markdown('<div class="epig-reset-btn">', unsafe_allow_html=True)
        if st.button("🔴 Reset Form", use_container_width=True):
            (_ := (getattr(st, 'rerun', None) or getattr(st, 'experimental_rerun', None)))()
        st.markdown('</div>', unsafe_allow_html=True)

st.write("---")
st.subheader("📋 Results")

def evaluate_all(rules_df, inputs):
    carriers = sorted(rules_df["Carrier"].dropna().unique().tolist())
    results = [evaluate_carrier(c, rules_df, inputs) for c in carriers]
    status_rank = {"Eligible":0,"Conditional":1,"Ineligible":2}
    results.sort(key=lambda r: (status_rank.get(r["status"],9), -r["score"], r["carrier"]))
    return results

data = load_rules()
rules = data["rules"]

if submitted:
    results = evaluate_all(rules, user_inputs)

    if all(r["status"]=="Ineligible" for r in results):
        st.info("If everything is **Ineligible**, double-check the Risk State, vehicle type, equipment age, UIIA, trailer & cargo selections, and excluded operations.")

    for res in results:
        with st.container(border=True):
            st.markdown(f"### {res['carrier']} — **{res['status']}**  |  Score: {int(res['score'])}")
            cols = st.columns(3)
            cols[0].markdown(f"**Portal:** {res['portal'] or '—'}")
            cols[1].markdown(f"**Submission:** {res['submission'] or '—'}")
            cols[2].markdown(f"**Required Fields:** {res['required_fields'] or '—'}")

            if res["hard_fails"]:
                st.error("Ineligible triggers:"); st.write("• " + "\n• ".join(res["hard_fails"][:12]))
            if res["required_fails"]:
                st.warning("Must-fix to qualify:"); st.write("• " + "\n• ".join(res["required_fails"][:12]))
            if res["soft_fails"]:
                st.info("Considerations:"); st.write("• " + "\n• ".join(res["soft_fails"][:12]))
            if res["matches"]:
                st.success("Matched appetite:"); st.write("• " + "\n• ".join(res["matches"][:12]))
else:
    st.caption("Fixed red Reset button added. All v14 rules remain (CW Cargo Van decline; GEICO whitelist & CA block; Nirvana states; 20yr equipment rule; CW 2+ yrs; trailer logic).")
