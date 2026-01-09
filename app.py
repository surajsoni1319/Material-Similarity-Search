import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz

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
# Text Cleaning Function
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
    file_path = "MARA 30.12.25.xlsx"  # Update path if required
    df = pd.read_excel(file_path)

    DESC_COL = "Material Description"  # verify once
    df["CLEAN_DESC"] = df[DESC_COL].apply(clean_text)

    return df, DESC_COL

df, DESC_COL = load_data()

# -------------------------------
# Similarity Logic
# -------------------------------
def calculate_similarity(search_term, material_text):
    return (
        0.4 * fuzz.token_set_ratio(search_term, material_text) +
        0.3 * fuzz.partial_ratio(search_term, material_text) +
        0.3 * fuzz.ratio(search_term, material_text)
    )

def find_similar_materials(
    search_term,
    df,
    top_n=20,
    min_score=60
):
    search_term = clean_text(search_term)

    results = []
    for _, row in df.iterrows():
        score = calculate_similarity(search_term, row["CLEAN_DESC"])
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
# Sidebar Controls
# -------------------------------
st.sidebar.header("⚙️ Search Settings")
top_n = st.sidebar.slider("Top Results", 5, 50, 20)
min_score = st.sidebar.slider("Minimum Similarity %", 50, 90, 60)

# -------------------------------
# Search UI
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
        st.warning("No similar materials found above the selected threshold.")
    else:
        st.success(f"Found {len(result_df)} similar materials")
        st.dataframe(result_df, use_container_width=True)

        # Download option
        csv = result_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️ Download Results as CSV",
            data=csv,
            file_name="material_similarity_results.csv",
            mime="text/csv"
        )

# -------------------------------
# Footer
# -------------------------------
st.markdown("---")
st.caption("🚀 Built for Star Cement | Material Master Intelligence")
