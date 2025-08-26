
import streamlit as st, yaml
st.title("01c — Defense Questionnaire (Targeted)")
st.caption("Short, high-impact prompts with conditional follow-ups for unmanned systems.")

bm = yaml.safe_load(open("templates.yaml","r",encoding="utf-8")).get("profiles",{}).get(st.session_state.get("profile_key"),{}).get("benchmarks",{})
def q_block(title):
    with st.expander(title):
        fpy = st.slider("Electronics FPY (%)", 50.0, 100.0, 96.0, 0.5, key=title+"_fpy")
        if fpy < bm.get("fpy_best",99.0):
            st.number_input("Top 3 SMT defects per 1k boards", value=120, step=10, key=title+"_smtdef")
        eco = st.number_input("Avg ECO cycle time (days)", value=12.0, step=1.0, key=title+"_eco")
        if eco > bm.get("eco_cycle_days_target",5):
            st.selectbox("Main ECO cause", ["Late requirement change","Supplier issue","Design escape","Process capability"], key=title+"_ecocause")
        fai = st.number_input("FAI First-Pass (%)", value=92.0, step=0.5, key=title+"_fai")
        if fai < bm.get("fai_pass_target_pct",95):
            st.checkbox("Missing ballooned drawings?", key=title+"_balloon")
            st.checkbox("Incomplete process routings?", key=title+"_routings")
        mtbf = st.number_input("MTBF (flight hrs)", value=300.0, step=10.0, key=title+"_mtbf")
        if mtbf < bm.get("mtbf_hours_target",500):
            st.selectbox("Dominant failure mode", ["Power","Comms link","Airframe","Payload"], key=title+"_failmode")
        st.slider("Traceability retrieval time (min)", 1, 180, 30, key=title+"_trace")

for t in ["Avionics/Embedded","Airframe & Structures","GCS & Comms","Payloads"]:
    q_block(t)

st.success("Inputs feed Observations, Defense KPIs and PACE.")
