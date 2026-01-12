import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz

# --------------------------------------------------
# Streamlit Page Config
# --------------------------------------------------
st.set_page_config(
    page_title="Material Similarity Search",
    layout="wide"
)

st.title("🔍 Material Similarity Search")
st.write("Search similar materials from the master list with similarity percentage.")

# --------------------------------------------------
# Cleaning Functions
# --------------------------------------------------
def basic_clean(text):
    if pd.isna(text):
        return ""

    text = str(text).upper()
    text = re.sub(r"[–\-/,()]", " ", text)   # normalize separators
    text = re.sub(r"\s+", " ", text)         # remove extra spaces
    return text.strip()


def load_cleaning_rules(path):
    rules_df = pd.read_excel(path)
    return dict(zip(rules_df["FIND"], rules_df["REPLACE"]))


def apply_rules(text, rules_dict):
    for find, replace in rules_dict.items():
        text = re.sub(rf"\b{find}\b", replace, text)
    return text


def clean_text(text, rules_dict):
    text = basic_clean(text)
    text = apply_rules(text, rules_dict)
    return text


# --------------------------------------------------
# Similarity Function
# --------------------------------------------------
def similarity_score(a, b):
    return fuzz.token_set_ratio(a, b)


# --------------------------------------------------
# Data Loading (Cached)
# --------------------------------------------------
@st.cache_data
def load_master_data():
    return pd.read_excel("data/material_master.xlsx")


@st.cache_data
def load_rules():
    return load_cleaning_rules("config/cleaning_rules.xlsx")


@st.cache_data
def prepare_master(df, rules):
    df = df.copy()
    df["CLEAN_TEXT"] = df["MATERIAL_DESCRIPTION"].apply(
        lambda x: clean_text(x, rules)
    )
    return df


# --------------------------------------------------
# Load Everything
# --------------------------------------------------
try:
    df_raw = load_master_data()
    rules = load_rules()
    df_master = prepare_master(df_raw, rules)

    st.success(f"Loaded {len(df_master)} materials successfully")

except Exception as e:
    st.error("Error loading data. Check file paths and column names.")
    st.stop()


# --------------------------------------------------
# Search UI
# --------------------------------------------------
search_term = st.text_input(
    "Enter material name to search",
    placeholder="Example: PIN"
)

min_score = st.slider(
    "Minimum similarity (%)",
    min_value=0,
    max_value=100,
    value=30
)


# --------------------------------------------------
# Search Logic
# --------------------------------------------------
if search_term:
    clean_query = clean_text(search_term, rules)

    df_master["SIMILARITY"] = df_master["CLEAN_TEXT"].apply(
        lambda x: similarity_score(clean_query, x)
    )

    results = (
        df_master[df_master["SIMILARITY"] >= min_score]
        .sort_values("SIMILARITY", ascending=False)
        .head(20)
    )

    st.subheader("🔎 Matching Results")
    st.caption(f"Cleaned search term: **{clean_query}**")

    st.dataframe(
        results[["MATERIAL_DESCRIPTION", "SIMILARITY"]],
        use_container_width=True
    )

    # Download button
    csv = results.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download results",
        csv,
        "material_matches.csv",
        "text/csv"
    )
