import streamlit as st
import pandas as pd
from io import BytesIO

# ================================
# Streamlit Page Configuration
# ================================
st.set_page_config(page_title="📘 Excel Viewer & Column Sampler", layout="centered")
st.title("📘 Excel Column Selector & Sampler")

# ================================
# File Uploader
# ================================
uploaded_file = st.file_uploader("📂 Upload your Excel/CSV file", type=["xlsx", "xls", "xlsm", "xlsb", "csv"])

if uploaded_file:
    # ================================
    # Choose Header Row
    # ================================
    st.subheader("⚙️ Choose Header Row")
    header_row = st.number_input(
        "Select header row (starts with 1)", 
        min_value=1, value=1, step=1
    )
    
    # Load Data (Pandas uses 0-index, so subtract 1)
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, header=header_row-1)
        else:
            df = pd.read_excel(uploaded_file, header=header_row-1)
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        st.stop()
    
    # ================================
    # Show First 5 Rows (Preview)
    # ================================
    st.subheader("👀 Preview of First 5 Rows")
    st.dataframe(df.head())
    
    # ================================
    # Show Columns and Allow Selection
    # ================================
    st.subheader("📋 Select Columns to Work With")
    all_columns = df.columns.tolist()
    
    selected_columns = st.multiselect(
        "✅ Choose columns to include:",
        options=all_columns,
        default=all_columns  # default = all columns
    )
    
    if selected_columns:
        st.write("🔎 Data with Selected Columns (first 5 rows):")
        st.dataframe(df[selected_columns].head())
    
    # ================================
    # Sampling Viewer
    # ================================
    st.subheader("🎲 Sample Data Viewer")
    sample_size = st.number_input(
        "Enter number of random rows to sample", 
        min_value=1, max_value=len(df), value=5, step=1
    )
    
    if st.button("Get Sample"):
        # Take random sample and filter columns
        sample_df = df[selected_columns].sample(sample_size)
        
        st.write(f"✅ Showing {sample_size} random rows with selected columns")
        st.dataframe(sample_df)
        
        # ================================
        # Save as Excel Option (Fixed)
        # ================================
        def to_excel_bytes(dataframe):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                dataframe.to_excel(writer, index=False)
            processed_data = output.getvalue()
            return processed_data

        st.download_button(
            label="💾 Download Sample as Excel",
            data=to_excel_bytes(sample_df),
            file_name="sample_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
