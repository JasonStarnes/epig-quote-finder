
import streamlit as st
import pandas as pd
import math
import os
from typing import Dict, Any, Tuple, List, Set

st.set_page_config(page_title="EPIG Quote Finder", page_icon="🚚", layout="wide")
st.title("🚚 EPIG Quote Finder — v3")

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

# Hard exclusions
EXCLUSION_KEYS = {
    "vehicle_type","operation","cargo_type","trailer_type",
    "safer_cargo","safer_operation",
    "ineligible_operations","excluded_operations"
}

# REQUIRED rules (must pass if present)
REQUIRED_KEYS = {
    "state_code","eld_installed","dashcam_eld_installed",
    "power_units","fleet_size","driver_experience_years",
    "years_in_business","license_required","gvw",
    "mvr_age_days","fmcsa_status","mcs90_required","vehicles_listed_vs_operated",
    "operation_type","operation_class","radius_min_for_new_federal_filing"
}

def eval_rule(op, rule_val, input_val):
    o = norm(op)
    if o in ["==","="]:   return norm(input_val) == norm(rule_val)
    if o == "!=":         return norm(input_val) != norm(rule_val)
    if o in [">",">=","<","<="]:
        a = try_float(input_val); b = try_float(rule_val)
        if math.isnan(a) or math.isnan(b): return False
        return (a>b) if o==">" else (a>=b) if o==">=" else (a<b) if o=="<" else (a<=b)
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
        lo,hi = between_val(rule_val); a = try_float(input_val)
        if math.isnan(a) or math.isnan(lo) or math.isnan(hi): return False
        return lo <= a <= hi
    return True

def evaluate_carrier(carrier, rules_df, inputs):
    rows = rules_df[rules_df["Carrier"]==carrier].copy()

    portal = rows["Portal_URL"].dropna().astype(str).head(1).tolist()
    portal = portal[0] if portal else ""
    submission = rows["Submission_Type"].dropna().astype(str).head(1).tolist()
    submission = submission[0] if submission else ""
    req_fields = rows["Required_Fields"].dropna().astype(str).head(1).tolist()
    req_fields = req_fields[0] if req_fields else ""

    hard_fails, required_fails, soft_fails, matches = [], [], [], []
    score = 0.0

    for _,r in rows.iterrows():
        key = norm(r.get("Criteria_Key",""))
        op  = r.get("Operator","")
        val = r.get("Value","")
        wt  = 0 if pd.isna(r.get("Weight",0)) else float(r.get("Weight",0))
        excl_notes = r.get("Excluded_Notes","")
        is_exclusion = (key in EXCLUSION_KEYS) or (isinstance(excl_notes,str) and excl_notes.strip()!="")

        # map inputs
        user_val = inputs.get(key, None)
        if key == "state_code":
            user_val = inputs.get("state_code")

        if user_val is None and not is_exclusion:
            continue

        passed = eval_rule(op, val, user_val if user_val is not None else "")

        if is_exclusion:
            if passed:
                hard_fails.append(f"{key} matches exclusion: {val} {('— '+str(excl_notes)) if isinstance(excl_notes,str) and excl_notes else ''}")
            else:
                matches.append(f"avoided exclusion: {key}")
                score += wt
        elif key in REQUIRED_KEYS:
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

st.sidebar.header("Agent Intake")

# Build comprehensive state list:
def extract_states_from_rules(rules_df) -> Set[str]:
    states: Set[str] = set()
    # 1) State_Code column
    if "State_Code" in rules_df.columns:
        for v in rules_df["State_Code"].dropna().astype(str).tolist():
            for s in v.split(","):
                s = s.strip().upper()
                if len(s)==2 and s.isalpha():
                    states.add(s)
    # 2) Any rows where Criteria_Key == 'state_code' and values in Value column are comma-separated
    if "Criteria_Key" in rules_df.columns and "Value" in rules_df.columns:
        subset = rules_df[rules_df["Criteria_Key"].astype(str).str.lower()=="state_code"]
        for v in subset["Value"].dropna().astype(str).tolist():
            for s in v.split(","):
                s = s.strip().upper()
                if len(s)==2 and s.isalpha():
                    states.add(s)
    return states

state_values = sorted(list(extract_states_from_rules(rules)))

# Fallback to all US states if none detected
US_STATES = ["AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA","ME",
             "MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA",
             "RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY","DC"]

if not state_values:
    state_values = US_STATES

state_code = st.sidebar.selectbox("Risk State", options=state_values, index=state_values.index("TX") if "TX" in state_values else 0)

radius = st.sidebar.selectbox("Radius", ["local","regional","interstate","long_haul"])
vehicle_type = st.sidebar.selectbox("Vehicle Type", ["cargo_van","box_truck","hotshot_class2","semi"])
gvw = st.sidebar.number_input("GVW (lbs)", min_value=0, max_value=200000, value=12000, step=500)
power_units = st.sidebar.number_input("Power Units (fleet size)", min_value=1, max_value=100, value=1, step=1)
years_in_business = st.sidebar.number_input("Years in Business", min_value=0, max_value=100, value=2, step=1)

driver_experience_years = st.sidebar.number_input("Min Driver Experience (years)", min_value=0, max_value=50, value=2, step=1)
mvr_age_days = st.sidebar.number_input("MVR Age (days)", min_value=0, max_value=365, value=30, step=1)
license_required = st.sidebar.selectbox("License Type", ["non_cdl","cdl"])

eld_installed = st.sidebar.selectbox("ELD / Dashcam Installed?", ["true","false"])
operation_type = st.sidebar.selectbox("Operation Type", ["for_hire","private"])

with st.sidebar.expander("Optional Details"):
    operation = st.selectbox("Operation Class", ["","brokerage","drive_away","livery","residential_delivery","final_delivery","towing","logging","hazmat","last_mile","dry_van","flatbed","reefer","straight","intermodal","non_tanker","heavy_haul"])
    cargo_type = st.text_input("Cargo Type(s) (comma-separated)", "")
    trailer_type = st.text_input("Trailer Type(s) (comma-separated)", "")
    safer_cargo = st.text_input("SAFER Cargo Selection(s) (comma-separated)", "")
    safer_operation = st.text_input("SAFER Operation Selection(s) (comma-separated)", "")

submitted = st.sidebar.button("Find Carriers", use_container_width=True)

user_inputs = {
    "state_code": norm(state_code),
    "radius": norm(radius),
    "vehicle_type": norm(vehicle_type),
    "gvw": gvw,
    "power_units": power_units,
    "years_in_business": years_in_business,
    "driver_experience_years": driver_experience_years,
    "mvr_age_days": mvr_age_days,
    "license_required": norm(license_required),
    "eld_installed": norm(eld_installed),
    "dashcam_eld_installed": norm(eld_installed),
    "operation_type": norm(operation_type),
    "operation_class": norm(operation) if operation else "",
    "operation": norm(operation) if operation else "",
    "cargo_type": to_list(cargo_type),
    "trailer_type": to_list(trailer_type),
    "safer_cargo": to_list(safer_cargo),
    "safer_operation": to_list(safer_operation),
}

st.write("---")
st.subheader("📋 Results")

def evaluate_all(rules_df, inputs):
    carriers = sorted(rules_df["Carrier"].dropna().unique().tolist())
    results = [evaluate_carrier(c, rules_df, inputs) for c in carriers]
    status_rank = {"Eligible":0,"Conditional":1,"Ineligible":2}
    results.sort(key=lambda r: (status_rank.get(r["status"],9), -r["score"], r["carrier"]))
    return results

if submitted:
    results = evaluate_all(rules, user_inputs)

    # If every carrier is ineligible because of state mismatch, hint to user
    if all(r["status"]=="Ineligible" for r in results):
        st.info("If everything is showing **Ineligible**, double-check the Risk State and excluded operations/cargo.")

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
    st.caption("Pick your Risk State (now includes all states from rules), add Years in Business (Berkley requires 3+), then click **Find Carriers**.")
