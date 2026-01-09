import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
import openpyxl  # ensures dependency is loaded

# --------------------------------------------------
# App Configuration
# --------------------------------------------------
st.set_page_config(
    page_title="Material Similarity Search | Star Cement",
    layout="wide"
)

st.title("🔍 Material Similarity Search")
st.caption("Duplicate material prevention & intelligent search")

# --------------------------------------------------
# Text Cleaning Function
# --------------------------------------------------
def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).upper()
    text = re.sub(r"[^A-Z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

# --------------------------------------------------
# Upload Excel File
# --------------------------------------------------
uploaded_file = st.file_uploader(
    "📤 Upload Material Master Excel File (.xlsx)",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("Please upload the material master Excel file to continue.")
    st.stop()

# --------------------------------------------------
# Load & Prepare Data (Cached)
# --------------------------------------------------
@st.cache_data(show_spinner=False)
def load_data(file):
    df = pd.read_excel(file, engine="openpyxl")

    DESC_COL = "Material Description"  # 🔴 update only if column name differs

    if DESC_COL not in df.columns:
        raise ValueError(f"Required column '{DESC_COL}' not found in file.")

    df["CLEAN_DESC"] = df[DESC_COL].apply(clean_text)
    return df, DESC_COL

try:
    df, DESC_COL = load_data(uploaded_file)
    st.success(f"✅ Loaded {len(df):,} materials successfully")
except Exception as e:
    st.error("Failed to load Excel file.")
    st.code(str(e))
    st.stop()

# --------------------------------------------------
# Similarity Logic
# --------------------------------------------------
def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio() * 100

def find_similar_materials(search_term, df, top_n=20, min_score=60):
    search_term = clean_text(search_term)
    results = []

    for _, row in df.iterrows():
        score = similarity(search_term, row["CLEAN_DESC"])
        if score >= min_score:
            results.append({
                "Material Description": row[DESC_COL],
                "Similarity %": round(score, 2)
            })

    result_df = pd.DataFrame(results)

    if not result_df.empty:
        result_df = result_df.sort_values(
            by="Similarity %",
            ascending=False
        ).head(top_n)

    return result_df

# --------------------------------------------------
# Sidebar Controls
# --------------------------------------------------
st.sidebar.header("⚙️ Search Settings")
top_n = st.sidebar.slider("Top Results", 5, 50, 20)
min_score = st.sidebar.slider("Minimum Similarity %", 50, 95, 60)

# --------------------------------------------------
# Search UI
# --------------------------------------------------
search_term = st.text_input(
    "🔎 Enter Material Name / Keyword",
    placeholder="Example: PIN, BOLT, MOTOR, BEARING"
)

if search_term:
    with st.spinner("Searching similar materials..."):
        result_df = find_similar_materials(
            search_term,
            df,
            top_n=top_n,
            min_score=min_score
        )

    if result_df.empty:
        st.warning("No similar materials found above the selected threshold.")
    else:
        st.success(f"Found {len(result_df)} similar materials")
        st.dataframe(result_df, use_container_width=True)

        csv = result_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Results as CSV",
            csv,
            "material_similarity_results.csv",
            "text/csv"
        )

# --------------------------------------------------
# Footer
# --------------------------------------------------
st.markdown("---")
st.caption("🚀 Star Cement | Material Master Similarity Automation")
