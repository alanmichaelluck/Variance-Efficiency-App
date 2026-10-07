import streamlit as st
import pandas as pd
import re
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# WIDE MODE CONFIGURATION & CSS
st.set_page_config(layout="wide", page_title="Variance/Efficiency Report")

st.markdown("""
    <style>
    .stDataFrame {
        width: 100% !important;
    }
    div[data-testid="stHorizontalBlock"] {
        width: 100%;
    }
    /* Highlight input boxes yellow */
    div[data-baseweb="input"] {
        background-color: #fff3cd !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("Variance/Efficiency Report")
st.markdown("💡 **To download total food AvT:** Reports / Inventory / Actual/Theoretical cost. Change dates then click 'Retrieve'. Then click 'Total Food'. Click 'print' Icon and export as EXCEL file. Upload to variance report.")

# Input fields for Dates and Store Name
col1, col2 = st.columns(2)
with col1:
    audit_dates = st.text_input("Dates (from AvT report)", placeholder="e.g., 8/11-8/15")
with col2:
    store_name_input = st.text_input("Store Name", placeholder="e.g., Flower Child - Austin (2nd)")

st.write("Drag and drop your Excel variance report below.")

# Updated GL Dictionary (Dairy and Bakery GLs mapped correctly)
gl_mapping = {
    'P50100': 'Dry Goods',
    'P50200': 'Produce / Veg',
    'P50300': 'Bakery',
    'P50400': 'Fish',
    'P50450': 'Other Seafood / Shrimp',
    'P50500': 'Beef',
    'P50600': 'Poultry',
    'P50900': 'Dairy'
}

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
        
        # Determine store name from manual input or fallback
        store_name = store_name_input.strip() if store_name_input and store_name_input.strip() else "Variance Report"
        
        # Display store location header
        st.markdown(f"# {store_name}")
        st.write("---")
        
        prod_row_idx = df_raw[df_raw.apply(lambda r: r.astype(str).str.contains('Product Number', case=False).any(), axis=1)].index[0]
        prod_col_idx = df_raw.iloc[prod_row_idx][df_raw.iloc[prod_row_idx] == 'Product Number'].index[0]
        
        df = df_raw[df_raw[prod_col_idx].astype(str).str.match(r'^P\d+-[A-Za-z0-9]+', na=False)].copy()
        df = df.dropna(axis=1, how='all')
        
        if df.shape[1] == 10:
            df.columns = ['Product Number', 'Product Name', 'Inv. Unit', 
                          'Actual Value', 'Actual %', 'Theoretical Value', 'Theoretical %', 
                          'Variance Value', 'Variance %', 'Approx. Units']
            
            # Exclude U-12 rows
            df = df[~df['Product Name'].astype(str).str.contains('U-12', case=False, na=False)].copy()
            
            df['GL Code'] = df['Product Number'].apply(lambda x: str(x).split('-')[0].strip())
            df['Category'] = df['GL Code'].map(gl_mapping).fillna('Uncategorized')
            
            df['Actual Value'] = pd.to_numeric(df['Actual Value'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df['Theoretical Value'] = pd.to_numeric(df['Theoretical Value'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df['Variance %'] = pd.to_numeric(df['Variance %'].astype(str).str.replace(',', '').str.replace('%', ''), errors='coerce').fillna(0)
            df['Approx. Units'] = pd.to_numeric(df['Approx. Units'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            # --- SECTION 1: THE SUMMARY TABLE ---
            summary = df.groupby(['GL Code', 'Category']).agg({
                'Actual Value': 'sum',
                'Theoretical Value': 'sum',
                'Variance %': 'sum'
            }).reset_index()
            
            summary['Efficiency'] = summary['Theoretical Value'] / summary['Actual Value'].replace(0, 1)
            summary['Variance ($)'] = summary['Theoretical Value'] - summary['Actual Value']
            
            category_order = {
                'Produce / Veg': 1,
                'Dry Goods': 2,
                'Poultry': 3,
                'Beef': 4,
                'Other Seafood / Shrimp': 5,
                'Fish': 6,
                'Dairy': 7,
                'Bakery': 8
            }
            summary['SortOrder'] = summary['Category'].map(category_order).fillna(99)
            summary = summary.sort_values(by='SortOrder')
            
            tot_actual = summary['Actual Value'].sum()
            tot_theo = summary['Theoretical Value'].sum()
            
            total_row = pd.DataFrame({
                'GL Code': ['TOTAL'],
                'Category': ['Total'],
                'Actual Value': [tot_actual],
                'Theoretical Value': [tot_theo],
                'Variance %': [summary['Variance %'].sum()],
                'Efficiency': [tot_theo / tot_actual if tot_actual > 0 else 0],
                'Variance ($)': [tot_theo - tot_actual],
                'SortOrder': [100]
            })
            
            summary_with_total = pd.concat([summary, total_row], ignore_index=True)
            summary_with_total = summary_with_total.sort_values(by='SortOrder')
            
            summary_with_total = summary_with_total[['GL Code', 'Category', 'Actual Value', 'Theoretical Value', 'Variance ($)', 'Variance %', 'Efficiency', 'SortOrder']]
