
import streamlit as st
import pandas as pd
import math
from typing import Dict, Any, Tuple, List

st.set_page_config(page_title="EPIG Quote Finder", page_icon="🚚", layout="wide")

st.title("🚚 EPIG Quote Finder")
st.caption("Uses your `carrier_rules_master.xlsx` to match carriers based on appetite & exclusions.")

@st.cache_data(show_spinner=False)
def load_rules() -> Dict[str, pd.DataFrame]:
    # The Excel must be placed in the same folder as this app.
    xls = pd.ExcelFile("carrier_rules_master.xlsx")
    data = {
        "agency": pd.read_excel(xls, "Agency_Appetite"),
        "inputs": pd.read_excel(xls, "Agent_Input_Fields"),
        "rules": pd.read_excel(xls, "Carrier_Rules"),
        "fields": pd.read_excel(xls, "Field_Dictionary"),
    }
    # Normalize some columns
    for col in ["Enabled","Priority","Weight"]:
        if col in data["rules"].columns:
            pass
    return data

def to_list(val: Any) -> List[str]:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return []
    if isinstance(val, list):
        return [str(x).strip() for x in val]
    s = str(val)
    # split on comma
    return [x.strip() for x in s.split(",") if str(x).strip() != ""]

def normalize(s: Any) -> str:
    if s is None:
        return ""
    return str(s).strip().lower()

def try_float(s: Any) -> float:
    try:
        return float(str(s).replace(",","").replace("$","").strip())
    except Exception:
        return float("nan")

def between_val(value: str) -> Tuple[float,float]:
    # supports "1-9" or "23-70"
    parts = str(value).replace(" ", "").split("-")
    if len(parts) == 2:
        return try_float(parts[0]), try_float(parts[1])
    return float("nan"), float("nan")

# Which criteria are strong "exclusions"
EXCLUSION_KEYS = set([
    "vehicle_type","operation","cargo_type","trailer_type",
    "safer_cargo","safer_operation",
    "ineligible_operations","excluded_operations"
])

def eval_rule(op: str, rule_val: str, input_val: Any) -> bool:
    op = normalize(op)
    if op in ["==","="]:
        return normalize(input_val) == normalize(rule_val)
    if op == "!=":
        return normalize(input_val) != normalize(rule_val)
    if op in [">",">=","<","<="]:
        a = try_float(input_val)
        b = try_float(rule_val)
        if math.isnan(a) or math.isnan(b):
            return False
        if op == ">": return a > b
        if op == ">=": return a >= b
        if op == "<": return a < b
        if op == "<=": return a <= b
    if op == "in":
        # input_val may be a single value or list; consider intersection
        vals = set([normalize(x) for x in to_list(rule_val)])
        if isinstance(input_val, list):
            inps = set([normalize(x) for x in input_val])
            return len(vals.intersection(inps)) > 0
        return normalize(input_val) in vals
    if op == "not in":
        vals = set([normalize(x) for x in to_list(rule_val)])
        if isinstance(input_val, list):
            inps = set([normalize(x) for x in input_val])
            return len(vals.intersection(inps)) == 0
        return normalize(input_val) not in vals
    if op == "between":
        lo, hi = between_val(rule_val)
        a = try_float(input_val)
        if math.isnan(a) or math.isnan(lo) or math.isnan(hi):
            return False
        return lo <= a <= hi
    # default: unknown operator -> can't evaluate -> treat as pass (non-blocking)
    return True

def evaluate_carrier(carrier: str, rules_df: pd.DataFrame, inputs: Dict[str, Any]) -> Dict[str, Any]:
    # Consider only this carrier's rules
    rows = rules_df[rules_df["Carrier"] == carrier].copy()

    score = 0.0
    hard_fails = []    # exclusions with notes
    soft_fails = []    # unmet preferences / eligibility we can evaluate
    matches = []       # satisfied rules we evaluated
    portal = ""
    submission = ""
    req_fields = ""
    notes = []

    for _, r in rows.iterrows():
        key   = normalize(r.get("Criteria_Key",""))
        op    = normalize(r.get("Operator",""))
        val   = r.get("Value","")
        wt    = r.get("Weight", 0) if not pd.isna(r.get("Weight", 0)) else 0
        note  = r.get("Notes","")
        excl  = r.get("Excluded_Notes","")
        state_override = normalize(r.get("State_Code",""))
        if not portal and isinstance(r.get("Portal_URL",None), str):
            portal = r.get("Portal_URL","")
        if not submission and isinstance(r.get("Submission_Type",None), str):
            submission = r.get("Submission_Type","")
        if not req_fields and isinstance(r.get("Required_Fields",None), str):
            req_fields = r.get("Required_Fields","")
        if isinstance(note,str) and note:
            notes.append(note)

        # State scope: if the rule is limited by State_Code and user's state doesn't match, skip the rule.
        if state_override and inputs.get("state_code"):
            if state_override != normalize(inputs["state_code"]):
                # some rows store many states in Value instead of State_Code; handled in operator
                # here, only skip if State_Code column is a specific single state and doesn't match
                pass

        # Pull user's value for this key
        user_val = inputs.get(key, inputs.get(key.replace("min_","").replace("max_",""), None))

        # Special handling for state_code rules: compare against user's state directly
        if key == "state_code":
            user_val = inputs.get("state_code")

        # If we don't have the user's value for this criterion and it's not clearly an exclusion list, skip
        if user_val is None and key not in EXCLUSION_KEYS and key != "state_code":
            continue

        # Evaluate the rule
        ok = eval_rule(op, val, user_val if user_val is not None else "")

        # Determine if this is an exclusion-type row
        is_exclusion = (normalize(excl) != "") or (key in EXCLUSION_KEYS and op in ["in","=="])

        if is_exclusion:
            if ok:  # input in an exclusion list
                hard_fails.append(f"{key} matches exclusion: {val} ({excl or 'Ineligible'})")
            else:
                # did not hit exclusion (good)
                score += float(wt or 0)
                matches.append(f"{key} avoided exclusion")
        else:
            # Non-exclusion appetite rule
            if ok:
                score += float(wt or 0)
                matches.append(f"{key} {op} {val}")
            else:
                soft_fails.append(f"{key} expected {op} {val}")

    status = "Eligible"
    if len(hard_fails) > 0:
        status = "Ineligible"
    elif len(soft_fails) > 0:
        status = "Conditional"

    return {
        "carrier": carrier,
        "status": status,
        "score": score,
        "portal": portal,
        "submission": submission,
        "required_fields": req_fields,
        "matches": matches,
        "soft_fails": soft_fails,
        "hard_fails": hard_fails,
        "notes": list(dict.fromkeys(notes))[:5]  # unique order, cap length
    }

# ---------------- Sidebar Form ----------------
data = load_rules()
rules = data["rules"]

st.sidebar.header("Agent Intake")

col1, col2 = st.sidebar.columns(2)
with col1:
    state_code = st.selectbox("Risk State", options=sorted({x for x in sum([to_list(v) for v in rules.get("State_Code","").fillna("").tolist()], []) if x}), index=0) if "State_Code" in rules.columns else st.text_input("Risk State (2-letter)", "TX")
with col2:
    radius = st.selectbox("Radius", ["local","regional","interstate","long_haul"])

vehicle_type = st.sidebar.selectbox("Vehicle Type", ["cargo_van","box_truck","hotshot_class2","semi"])
gvw = st.sidebar.number_input("GVW (lbs)", min_value=0, max_value=200000, value=12000, step=500)
power_units = st.sidebar.number_input("Power Units (fleet size)", min_value=1, max_value=100, value=1, step=1)

driver_experience_years = st.sidebar.number_input("Min Driver Experience (years)", min_value=0, max_value=50, value=2, step=1)
mvr_quality = st.sidebar.selectbox("MVR Quality", ["clean","mixed","poor","unknown"])
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

# Build input dict for evaluation
user_inputs = {
    "state_code": normalize(state_code),
    "radius": normalize(radius),
    "vehicle_type": normalize(vehicle_type),
    "gvw": gvw,
    "power_units": power_units,
    "driver_experience_years": driver_experience_years,
    "mvr_quality": normalize(mvr_quality),
    "license_required": normalize(license_required),
    "eld_installed": normalize(eld_installed),
    "dashcam_eld_installed": normalize(eld_installed),
    "operation_type": normalize(operation_type),
    "operation": normalize(operation) if operation else "",
    "cargo_type": [normalize(x) for x in to_list(cargo_type)],
    "trailer_type": [normalize(x) for x in to_list(trailer_type)],
    "safer_cargo": [normalize(x) for x in to_list(safer_cargo)],
    "safer_operation": [normalize(x) for x in to_list(safer_operation)],
}

# ---------------- Results ----------------
st.write("---")
st.subheader("📋 Results")

if submitted:
    carriers = sorted(rules["Carrier"].dropna().unique().tolist())
    results = []
    for c in carriers:
        results.append(evaluate_carrier(c, rules, user_inputs))

    # Sort: Ineligible last, by status then score desc
    status_rank = {"Eligible": 0, "Conditional": 1, "Ineligible": 2}
    results.sort(key=lambda r: (status_rank.get(r["status"], 9), -r["score"], r["carrier"]))

    # Display all matches
    for res in results:
        with st.container(border=True):
            header = f"**{res['carrier']}** — {res['status']}  |  Score: {int(res['score'])}"
            st.markdown(header)
            cols = st.columns(3)
            cols[0].markdown(f"**Portal:** {res['portal'] or '—'}")
            cols[1].markdown(f"**Submission:** {res['submission'] or '—'}")
            cols[2].markdown(f"**Required Fields:** {res['required_fields'] or '—'}")

            if res["matches"]:
                st.markdown("**Why it matched:**")
                st.write("• " + "\n• ".join(res["matches"][:8]))

            if res["soft_fails"]:
                st.warning("Needs review:")
                st.write("• " + "\n• ".join(res["soft_fails"][:8]))

            if res["hard_fails"]:
                st.error("Ineligible triggers:")
                st.write("• " + "\n• ".join(res["hard_fails"][:8]))

            if res["notes"]:
                st.caption("Notes: " + " | ".join(res["notes"]))

    st.info("Tip: If a carrier shows as **Ineligible**, check exclusions & SAFER selections.")

else:
    st.write("Fill out the form on the left and click **Find Carriers** to see matches.")
    st.caption("Make sure `carrier_rules_master.xlsx` is in the same folder as this app.")
