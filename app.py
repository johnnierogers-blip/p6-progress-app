import streamlit as st
import pandas as pd

st.set_page_config(page_title="P6 Progress", layout="wide")
st.title("P6 Progress Tracker — No More Excel")

file = st.file_uploader("Upload your full P6 export (.xlsx)", type="xlsx")

if file:
    # Load P6 Dump, skip 4 rows (your file's structure)
    df = pd.read_excel(file, sheet_name="P6 Dump", skiprows=4, engine="openpyxl")
    df = df.dropna(how="all").reset_index(drop=True)

    # Fix P6 decimal % → real % (0.6991 → 69.9)
    if "Duration % Complete" in df.columns:
        df["Current % ↓↓↓"] = (df["Duration % Complete"] * 100).round(1)
    else:
        df["Current % ↓↓↓"] = 0.0

    st.success(f"Loaded {len(df):,} activities")

    tab1, tab2, tab3 = st.tabs(["PM Update", "Dashboard", "Export"])

    with tab1:
        # Only show columns that exist
        cols = ["Activity ID"]
        if "Activity Name " in df.columns:
            cols.append("Activity Name ")
        elif "Activity Name" in df.columns:
            cols.append("Activity Name")
        cols.append("Current % ↓↓↓")
        if "Activity Status" in df.columns:
            cols.append("Activity Status")

        edited = st.data_editor(
            df[cols],
            column_config={"Current % ↓↓↓": st.column_config.NumberColumn(min_value=0, max_value=100, step=1)},
            use_container_width=True,
        )

    with tab2:
        st.metric("Overall Progress", f"{df['Current % ↓↓↓'].mean():.1f}%")

    with tab3:
        output = BytesIO()
        export_df = df.copy()
        if "Duration % Complete" in export_df.columns:
            export_df["Duration % Complete"] = export_df["Current % ↓↓↓"] / 100
        export_df.to_excel(output, index=False)
        st.download_button("Download for P6 import", output.getvalue(), "P6_updated.xlsx")
