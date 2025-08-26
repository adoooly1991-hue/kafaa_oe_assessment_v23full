
import streamlit as st, yaml
st.set_page_config(page_title="Kafaa OE Assessment — v25.3 Full", layout="wide")
st.sidebar.title("Kafaa OE Assessment"); st.sidebar.caption("v25.3 • Full‑fat (Defense profile)")

CARD_CSS = "<style>:root{--k-primary:#C00000;--k-muted:#f7f7f9}.block-container{padding-top:1.2rem}.k-card{background:white;border-radius:14px;padding:14px 16px;box-shadow:0 2px 10px rgba(0,0,0,.06);border:1px solid #eee}</style>"
st.markdown(CARD_CSS, unsafe_allow_html=True)

with st.sidebar.expander("✨ Tour & Glossary"):
    st.toggle("Tour mode (show explanations)", value=True, key="tour_mode")
    st.write("Takt = Available time / Demand; CT = Cycle time; WIP = Work‑in‑process.")

templates = yaml.safe_load(open("templates.yaml","r",encoding="utf-8"))
profiles = templates.get("profiles",{})
profile_key = st.sidebar.selectbox("Industry profile", options=list(profiles.keys()), format_func=lambda k: profiles[k]["label"], index=list(profiles.keys()).index("defense_unmanned") if "defense_unmanned" in profiles else 0)
st.session_state["profile_key"] = profile_key

st.title("Welcome")
st.write("Start with **01a — Data Collection**, **01b/01c — Questionnaires**. Then **02 — Product & Financials** and **03a/b/c** dashboards.")

with st.expander("Quick CT/Takt setup"):
    col1, col2 = st.columns(2)
    with col1:
        st.session_state["avail_sec"] = st.number_input("Available time per shift (sec)", value=28800, step=600)
        st.session_state["demand_units"] = st.number_input("Demand per shift (units)", value=100, step=10)
        st.session_state["takt_sec"] = st.session_state["avail_sec"] / max(st.session_state["demand_units"], 1)
    with col2:
        st.write("Enter CTs in **05 — VSM Diagram**; dashboards & exports will use them.")
