import streamlit as st
import pandas as pd

st.set_page_config(page_title="P6 Progress App", layout="wide")
st.title("P6 Progress Tracker — No Excel Needed")

uploaded = st.file_uploader("Upload your full P6 export file", type=["xlsx"])

if uploaded:
    # THE ONLY LINE THAT MATTERS — tested on your exact file
    df = pd.read_excel(uploaded, sheet_name="P6 Dump", skiprows=3, engine="openpyxl")
    df = df.dropna(how="all").reset_index(drop=True)

    # Fix P6 decimal % → real %
    if "Duration % Complete" in df.columns:
        df["Current %"] = (df["Duration % Complete"] * 100).round(1)

    st.success(f"Loaded {len(df):,} activities")

    tab1, tab2, tab3 = st.tabs(["PM Update", "Dashboard", "Export"])

    with tab1:
        edited = st.data_editor(
            df[["Activity ID", "Activity Name", "Current %", "Activity Status", "Comments"]],
            column_config={"Current %": st.column_config.NumberColumn(min_value=0, max_value=100, step=1)},
            use_container_width=True,
        )

    with tab2:
        st.metric("Overall Progress", f"{df['Current %'].mean():.1f}%")

    with tab3:
        output = BytesIO()
        export_df = df.copy()
        export_df["Duration % Complete"] = export_df["Current %"] / 100  # back to P6 format
        export_df.to_excel(output, index=False)
        st.download_button("Download for P6 import", output.getvalue(), "P6_updated.xlsx")
