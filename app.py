import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher

# -------------------------------
# App Configuration
# -------------------------------
st.set_page_config(
    page_title="Material Similarity Search | Star Cement",
    layout="wide"
)

st.title("🔍 Material Similarity Search")
st.caption("Duplicate material prevention & intelligent search")

# -------------------------------
# Text Cleaning
# -------------------------------
def clean_text(text):
    if pd.isna(text):
        return ""
    text = text.upper()
    text = re.sub(r"[^A-Z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

# -------------------------------
# Load Dataset (Cached)
# -------------------------------
@st.cache_data
def load_data():
    df = pd.read_excel("MARA 30.12.25.xlsx")
    DESC_COL = "Material Description"
    df["CLEAN_DESC"] = df[DESC_COL].apply(clean_text)
    return df, DESC_COL

df, DESC_COL = load_data()

# -------------------------------
# Similarity Logic (Python Std Lib)
# -------------------------------
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

# -------------------------------
# Sidebar
# -------------------------------
st.sidebar.header("⚙️ Settings")
top_n = st.sidebar.slider("Top Results", 5, 50, 20)
min_score = st.sidebar.slider("Minimum Similarity %", 50, 90, 60)

# -------------------------------
# UI
# -------------------------------
search_term = st.text_input(
    "Enter Material Name / Keyword",
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
        st.warning("No similar materials found.")
    else:
        st.success(f"Found {len(result_df)} similar materials")
        st.dataframe(result_df, use_container_width=True)

        csv = result_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Results",
            csv,
            "material_similarity_results.csv",
            "text/csv"
        )

st.markdown("---")
st.caption("🚀 Star Cement | Material Master Intelligence")
