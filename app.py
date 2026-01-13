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
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
    <style>
    .metric-container {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🔍 Material Similarity Search")
st.caption("Search similar materials from SAP MARA master data using fuzzy matching")

# --------------------------------------------------
# CLEANING FUNCTIONS
# --------------------------------------------------
def clean_text(text):
    """Clean and normalize text for comparison"""
    if pd.isna(text):
        return ""
    text = str(text).upper()
    text = re.sub(r"[–\-/,()]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# --------------------------------------------------
# SIMILARITY FUNCTIONS
# --------------------------------------------------
def similarity_score(a, b, method="token_set"):
    """Calculate similarity score using selected method"""
    if method == "token_set":
        return fuzz.token_set_ratio(a, b)
    elif method == "token_sort":
        return fuzz.token_sort_ratio(a, b)
    elif method == "partial":
        return fuzz.partial_ratio(a, b)
    else:
        return fuzz.ratio(a, b)

# --------------------------------------------------
# LOAD DATA (CACHED)
# --------------------------------------------------
@st.cache_data
def load_data():
    """Load and preprocess Excel data"""
    try:
        df = pd.read_excel(EXCEL_FILE)
        
        # Validate required columns
        if MATERIAL_CODE_COL not in df.columns or MATERIAL_NAME_COL not in df.columns:
            st.error(f"Required columns not found. Expected: '{MATERIAL_CODE_COL}' and '{MATERIAL_NAME_COL}'")
            return None
        
        # Clean and prepare data
        df["CLEAN_NAME"] = df[MATERIAL_NAME_COL].apply(clean_text)
        
        # Remove duplicates and empty names
        df = df[df["CLEAN_NAME"] != ""].drop_duplicates(subset=[MATERIAL_CODE_COL])
        
        return df
    except FileNotFoundError:
        st.error(f"❌ File '{EXCEL_FILE}' not found. Please ensure the file is in the same directory.")
        return None
    except Exception as e:
        st.error(f"❌ Error loading file: {str(e)}")
        return None

# --------------------------------------------------
# LOAD DATA
# --------------------------------------------------
with st.spinner("Loading material data..."):
    df_master = load_data()

if df_master is None:
    st.stop()

st.success(f"✅ Loaded {len(df_master):,} unique materials successfully")

# --------------------------------------------------
# SIDEBAR CONTROLS
# --------------------------------------------------
st.sidebar.header("⚙️ Search Settings")

# Matching method
match_method = st.sidebar.selectbox(
    "Matching Algorithm",
    ["token_set", "token_sort", "partial", "simple"],
    index=0,
    help="token_set: Best for word order differences\ntoken_sort: Good for sorted tokens\npartial: Finds substrings\nsimple: Direct character comparison"
)

# Similarity threshold
min_similarity = st.sidebar.slider(
    "Minimum similarity (%)",
    min_value=0,
    max_value=100,
    value=50,
    help="Lower values show more results but may be less relevant"
)

# Max results
max_results = st.sidebar.selectbox(
    "Max results to display",
    [50, 100, 200, 500],
    index=1
)

# Advanced options
with st.sidebar.expander("🔧 Advanced Options"):
    case_sensitive = st.checkbox("Case sensitive search", value=False)
    show_stats = st.checkbox("Show statistics", value=True)
    
st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Database Info")
st.sidebar.metric("Total Materials", f"{len(df_master):,}")
st.sidebar.metric("Unique Materials", f"{df_master[MATERIAL_CODE_COL].nunique():,}")

# --------------------------------------------------
# SEARCH BAR
# --------------------------------------------------
col1, col2 = st.columns([3, 1])

with col1:
    search_term = st.text_input(
        "🔎 Search Material Name",
        placeholder="Example: PIN, BOLT, BEARING, etc.",
        help="Enter material name or keywords to find similar items"
    )

with col2:
    st.markdown("<br>", unsafe_allow_html=True)
    search_button = st.button("🔍 Search", type="primary", use_container_width=True)

# --------------------------------------------------
# SEARCH LOGIC
# --------------------------------------------------
if search_term and (search_button or True):  # Auto-search on type
    
    clean_query = clean_text(search_term) if not case_sensitive else search_term.strip()
    
    with st.spinner("Searching for matches..."):
        # Calculate similarities
        df_master["SIMILARITY"] = df_master["CLEAN_NAME"].apply(
            lambda x: similarity_score(clean_query, x, match_method)
        )
        
        # Filter and sort results
        results = (
            df_master[df_master["SIMILARITY"] >= min_similarity]
            .sort_values("SIMILARITY", ascending=False)
            .head(max_results)
        )
    
    # --------------------------------------------------
    # DISPLAY RESULTS
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("📊 Search Results")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Matches Found", len(results))
    with col2:
        st.metric("Avg Similarity", f"{results['SIMILARITY'].mean():.1f}%" if len(results) > 0 else "N/A")
    with col3:
        st.metric("Best Match", f"{results['SIMILARITY'].max():.1f}%" if len(results) > 0 else "N/A")
    with col4:
        st.metric("Search Term", f"'{clean_query}'")
    
    if len(results) == 0:
        st.warning("⚠️ No matches found. Try lowering the similarity threshold or using different keywords.")
        st.stop()
    
    # --------------------------------------------------
    # STATISTICS CHART
    # --------------------------------------------------
    if show_stats and len(results) > 0:
        st.markdown("### 📈 Similarity Distribution")
        
        # Create histogram data
        bins = pd.cut(results["SIMILARITY"], bins=10)
        hist_data = results.groupby(bins, observed=True).size().reset_index()
        hist_data.columns = ["Range", "Count"]
        
        # Display as bar chart using Streamlit's native chart
        st.bar_chart(results["SIMILARITY"], height=250)
    
    # --------------------------------------------------
    # TOP 10 RESULTS
    # --------------------------------------------------
    st.markdown("### 🔝 Top 10 Matches")
    
    top_10 = results.head(10).copy()
    top_10["Rank"] = range(1, len(top_10) + 1)
    
    display_cols = ["Rank", MATERIAL_CODE_COL, MATERIAL_NAME_COL, "SIMILARITY"]
    
    # Style the dataframe
    styled_top_10 = top_10[display_cols].style.background_gradient(
        subset=["SIMILARITY"],
        cmap="RdYlGn",
        vmin=min_similarity,
        vmax=100
    )
    
    st.dataframe(styled_top_10, use_container_width=True, hide_index=True)
    
    # --------------------------------------------------
    # ADDITIONAL RESULTS
    # --------------------------------------------------
    remaining = results.iloc[10:]
    
    if len(remaining) > 0:
        with st.expander(f"📋 Show {len(remaining)} more results"):
            st.dataframe(
                remaining[[MATERIAL_CODE_COL, MATERIAL_NAME_COL, "SIMILARITY"]],
                use_container_width=True,
                hide_index=True
            )
    
    # --------------------------------------------------
    # EXPORT OPTIONS
    # --------------------------------------------------
    st.markdown("---")
    st.markdown("### 💾 Export Results")
    
    col1, col2 = st.columns(2)
    
    # Prepare export data
    output_df = results[[MATERIAL_CODE_COL, MATERIAL_NAME_COL, "SIMILARITY"]].copy()
    output_df.rename(columns={
        MATERIAL_CODE_COL: "Material Code",
        MATERIAL_NAME_COL: "Material Name",
        "SIMILARITY": "Similarity (%)"
    }, inplace=True)
    
    # Excel export
    with col1:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            output_df.to_excel(writer, index=False, sheet_name="Matches")
            
            # Add search info sheet
            info_df = pd.DataFrame({
                "Parameter": ["Search Term", "Cleaned Term", "Method", "Min Similarity", "Total Matches"],
                "Value": [search_term, clean_query, match_method, f"{min_similarity}%", len(results)]
            })
            info_df.to_excel(writer, index=False, sheet_name="Search Info")
        
        st.download_button(
            label="📥 Download as Excel",
            data=buffer.getvalue(),
            file_name=f"material_search_{clean_query.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # CSV export
    with col2:
        csv = output_df.to_csv(index=False)
        st.download_button(
            label="📥 Download as CSV",
            data=csv,
            file_name=f"material_search_{clean_query.replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True
        )

else:
    # --------------------------------------------------
    # INITIAL STATE
    # --------------------------------------------------
    st.info("👆 Enter a material name above to start searching")
    
    # Show sample data
    with st.expander("📖 View Sample Materials"):
        sample = df_master.sample(min(10, len(df_master)))[[MATERIAL_CODE_COL, MATERIAL_NAME_COL]]
        st.dataframe(sample, use_container_width=True, hide_index=True)
    
    # Usage instructions
    with st.expander("ℹ️ How to Use"):
        st.markdown("""
        **Instructions:**
        1. Enter a material name or keywords in the search box
        2. Adjust the similarity threshold if needed (lower = more results)
        3. Choose a matching algorithm based on your needs
        4. Review the top matches and download results
        
        **Tips:**
        - Use keywords rather than full descriptions for better results
        - Try different matching algorithms if results aren't satisfactory
        - Lower the similarity threshold to see more potential matches
        - The token_set method works best for materials with varying word orders
        """)

# --------------------------------------------------
# FOOTER
# --------------------------------------------------
st.markdown("---")
st.caption("💡 Tip: Use the sidebar to adjust search parameters for better results")
