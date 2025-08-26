
import streamlit as st, pandas as pd
st.title("01a — Data Collection")
st.caption("Upload financials & material flow sheets. Recognizers auto-fill key fields.")

fin = st.file_uploader("Financials (CSV/XLSX)", type=["csv","xlsx"])
if fin:
    try:
        df = pd.read_csv(fin) if fin.name.lower().endswith(".csv") else pd.read_excel(fin)
        st.dataframe(df.head())
        for col in ["Revenue","COGS","G&A","Depreciation","Financial Expenses","Inventory","Current Assets","Current Liabilities"]:
            if col in df.columns:
                st.session_state[f"financials_{col.lower().replace(' ','_')}"] = float(df[col].iloc[0])
        st.success("Financial fields captured.")
    except Exception as e:
        st.warning(f"Parse error: {e}")

mat = st.file_uploader("Material Flow (CSV/XLSX)", type=["csv","xlsx"], key="mat")
if mat:
    try:
        dfm = pd.read_csv(mat) if mat.name.lower().endswith(".csv") else pd.read_excel(mat)
        st.dataframe(dfm.head())
        step_col = "Step" if "Step" in dfm.columns else dfm.columns[0]
        ct_col = "CT (s)" if "CT (s)" in dfm.columns else dfm.columns[1]
        rows = [{"step": str(r[step_col]), "ct": float(r[ct_col])} for _, r in dfm.iterrows() if pd.notnull(r[step_col])]
        st.session_state["ct_table"] = rows
        st.success(f"Loaded {len(rows)} steps into VSM Diagram.")
    except Exception as e:
        st.warning(f"Could not parse material flow: {e}")
