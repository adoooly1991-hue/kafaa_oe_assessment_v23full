
import streamlit as st, yaml
st.title("01b — Guided Questionnaire")
bm = yaml.safe_load(open("templates.yaml","r",encoding="utf-8")).get("profiles",{}).get(st.session_state.get("profile_key"),{}).get("benchmarks",{})
best_fpy = bm.get("fpy_best", 98.0); best_chg = bm.get("chg_min", 10.0); best_dio = bm.get("inventory_days", 45)
stages = st.session_state.get("stages") or ["Ordering","Inbound Logistics","Raw Storage","Issue to Prod","Forming","Assembly","Welding","Coating","FG Storage","Outbound","Shipping"]
def ask(name):
    with st.expander(f"🔎 {name}"):
        st.caption(f"Best practice: FPY ≥ {best_fpy}%, Changeover ≤ {best_chg} min, DIO ≤ {best_dio} days.")
        fpy = st.slider("First Pass Yield (%)", 50.0, 100.0, 95.0, 0.5, key=f"fpy_{name}")
        scrap = st.slider("Scrap/Rework (% of output)", 0.0, 30.0, 3.0, 0.5, key=f"scrap_{name}")
        changeover = st.number_input("Average changeover (min)", value=20.0, step=5.0, key=f"chg_{name}")
        wip = st.number_input("WIP before this step (units)", value=0.0, step=10.0, key=f"wip_{name}")
        wait = st.selectbox("Material waits > 30 min?", ["No","Sometimes","Often"], key=f"wait_{name}")
        sev = max(0, (best_fpy - fpy))*0.5 + scrap*0.8 + (changeover/10.0) + (5.0 if wait=='Sometimes' else (10.0 if wait=='Often' else 0)) + (wip/50.0)
        wastes = {}
        if fpy < best_fpy or scrap>0: wastes["defects"]= (100-fpy) + scrap
        if changeover>best_chg: wastes["waiting"]= changeover/2
        if wip>0: wastes["inventory"]= wip/10
        if wait!="No": wastes["motion"]= 3 if wait=="Sometimes" else 6
        st.session_state.setdefault("stage_waste_scores", {})[name] = {"severity": sev, "wastes": wastes}
for s in stages: ask(s)
st.success("Captured. Feeds Observations & PACE.")
