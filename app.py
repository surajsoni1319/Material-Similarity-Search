import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from io import BytesIO

# --------------------------------------------------
# CONFIG
# --------------------------------------------------
EXCEL_FILE = "MARA 30.12.25.xlsx"

MATERIAL_CODE_COL = "Material"
MATERIAL_NAME_COL = "Material Description"

# --------------------------------------------------
# PAGE SETUP
# --------------------------------------------------
st.set_page_config(
    page_title="Material Similarity Search",
    layout="wide"
)

st.title("🔍 Material Similarity Search")
st.caption("Search similar materials from SAP MARA master data")

# --------------------------------------------------
# CLEANING FUNCTIONS
# --------------------------------------------------
def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).upper()
    text = re.sub(r"[–\-/,()]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# --------------------------------------------------
# SIMILARITY FUNCTION
# --------------------------------------------------
def similarity_score(a, b):
    return fuzz.token_set_ratio(a, b)

# --------------------------------------------------
# LOAD DATA (CACHED)
# --------------------------------------------------
@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL_FILE)
    df["CLEAN_NAME"] = df[MATERIAL_NAME_COL].apply(clean_text)
    return df

try:
    df_master = load_data()
    st.success(f"Loaded {len(df_master)} materials successfully")
except Exception as e:
    st.error("❌ Failed to load Excel file. Check file name or column names.")
    st.stop()

# --------------------------------------------------
# SIDEBAR CONTROLS
# --------------------------------------------------
st.sidebar.header("⚙️ Controls")

min_similarity = st.sidebar.slider(
    "Minimum similarity (%)",
    min_value=0,
    max_value=100,
    value=50
)

max_results = st.sidebar.selectbox(
    "Max results to process",
    [50, 100, 200],
    index=1
)

# --------------------------------------------------
# SEARCH BAR
# --------------------------------------------------
search_term = st.text_input(
    "🔎 Search Material Name",
    placeholder="Example: PIN"
)

# --------------------------------------------------
# SEARCH LOGIC
# --------------------------------------------------
if search_term:

    clean_query = clean_text(search_term)

    df_master["SIMILARITY"] = df_master["CLEAN_NAME"].apply(
        lambda x: similarity_score(clean_query, x)
    )

    results = (
        df_master[df_master["SIMILARITY"] >= min_similarity]
        .sort_values("SIMILARITY", ascending=False)
        .head(max_results)
    )

    st.subheader("📊 Matching Results")
    st.caption(f"Cleaned search term: **{clean_query}**")

    # -------------------------------
    # TOP 10 RESULTS
    # -------------------------------
    top_10 = results.head(10)

    st.markdown("### 🔝 Top 10 Matches")
    st.dataframe(
        top_10[
            [MATERIAL_CODE_COL, MATERIAL_NAME_COL, "SIMILARITY"]
        ],
        use_container_width=True
    )

    # -------------------------------
    # SHOW MORE RESULTS
    # -------------------------------
    remaining = results.iloc[10:]

    if len(remaining) > 0:
        with st.expander(f"Show more results ({len(remaining)} items)"):
            st.dataframe(
                remaining[
                    [MATERIAL_CODE_COL, MATERIAL_NAME_COL, "SIMILARITY"]
                ],
                use_container_width=True
            )

    # -------------------------------
    # DOWNLOAD AS EXCEL
    # -------------------------------
    output_df = results[
        [MATERIAL_CODE_COL, MATERIAL_NAME_COL, "SIMILARITY"]
    ].copy()

    output_df.rename(columns={
        MATERIAL_CODE_COL: "Material Code",
        MATERIAL_NAME_COL: "Material Name",
        "SIMILARITY": "Similarity (%)"
    }, inplace=True)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False, sheet_name="Matches")

    st.download_button(
        label="⬇️ Download Matching List (Excel)",
        data=buffer.getvalue(),
        file_name="material_similarity_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Enter a material name to start searching")
