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
    </style>
""", unsafe_allow_html=True)

st.title("Variance/Efficiency Report")
st.markdown("Download total food AvT: Reports / Inventory / Actual/Theoretical cost. Change dates then click 'Retrieve'. Then click 'Total Food'. Print as EXCEL file. Upload to variance report.")
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
        
        raw_title_text = str(df_raw.iloc[0, 1])
        store_title_parts = re.split(r'\s+\d+\s+[A-Za-z]', raw_title_text)
        store_name = store_title_parts[0].strip() if len(store_title_parts) > 0 else "Variance Report"
        
        # Display just the store location name
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
            
            # --- SECTION 1: THE SUMMARY TABLE ---
            summary = df.groupby(['GL Code', 'Category']).agg({
                'Actual Value': 'sum',
                'Theoretical Value': 'sum',
                'Variance %': 'sum'
            }).reset_index()
            
            summary['Efficiency'] = summary['Theoretical Value'] / summary['Actual Value'].replace(0, 1)
            summary['Variance ($)'] = summary['Actual Value'] - summary['Theoretical Value']
            
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
            
            total_row = pd.DataFrame({
                'GL Code': ['TOTAL'],
                'Category': ['Total'],
                'Actual Value': [summary['Actual Value'].sum()],
                'Theoretical Value': [summary['Theoretical Value'].sum()],
                'Variance %': [summary['Variance %'].sum()],
                'Efficiency': [summary['Theoretical Value'].sum() / summary['Actual Value'].sum() if summary['Actual Value'].sum() > 0 else 0],
                'Variance ($)s': [summary['Variance ($)'].sum()],
                'SortOrder': [100]
            })
            
            total_row = total_row.rename(columns={'Variance ($)s': 'Variance ($)'})
            summary_with_total = pd.concat([summary, total_row], ignore_index=True)
            summary_with_total = summary_with_total.sort_values(by='SortOrder')
            
            summary_with_total = summary_with_total[['GL Code', 'Category', 'Actual Value', 'Theoretical Value', 'Variance ($)', 'Variance %', 'Efficiency', 'SortOrder']]
            
            df_items = df[['GL Code', 'Category', 'Product Number', 'Product Name', 'Actual Value', 'Theoretical Value', 'Variance %']].copy()
            df_items['Efficiency'] = df_items['Theoretical Value'] / df_items['Actual Value'].replace(0, 1)
            df_items['Variance ($)'] = df_items['Actual Value'] - df_items['Theoretical Value']
            
            sorted_categories_for_loop = summary.sort_values(by='SortOrder')
            
            highlight_limits = {
                'Produce / Veg': 3,
                'Dry Goods': 3,
                'Dairy': 2,
                'Fish': 1,
                'Other Seafood / Shrimp': 1,
                'Beef': 1,
                'Poultry': 1,
                'Bakery': 1
            }

            st.write("### Total Food")
            
            def bold_total_row(row):
                if row['GL Code'] == 'TOTAL':
                    return ['font-weight: bold; background-color: #f8f9fa'] * len(row)
                return [''] * len(row)

            st.dataframe(summary_with_total.drop(columns=['SortOrder']).style.apply(bold_total_row, axis=1).format({
                'Actual Value': '${:,.2f}', 
                'Theoretical Value': '${:,.2f}',
                'Variance ($)': '${:,.2f}',
                'Variance %': '{:.2f}%',
                'Efficiency': '{:.2%}'
            }), use_container_width=True, hide_index=True)
            
            # --- SECTION 2: DETAILED ITEM BREAKDOWNS & ACTION PLAN FORM ---
            st.write("---")
            st.write("### Category Breakdowns")
            st.write("*💡 Tip: Click any column header to re-sort items. Log action plans in the form below and click 'Save Action Plans' before downloading.*")
            
            with st.form("action_plans_form"):
                for index, row in sorted_categories_for_loop.iterrows():
                    cat_name = row['Category']
                    gl_code = row['GL Code']
                    
                    st.write(f"#### {cat_name} ({gl_code})")
                    
                    cat_df = df_items[df_items['GL Code'] == gl_code].copy()
                    cat_df = cat_df[['Product Number', 'Product Name', 'Actual Value', 'Theoretical Value', 'Variance ($)', 'Variance %', 'Efficiency']]
                    
                    cat_df = cat_df.sort_values(by='Variance %', ascending=True).reset_index(drop=True)
                    
                    limit = highlight_limits.get(cat_name, 1)
                    
                    def make_highlight_func(n):
                        def highlight_top_n(df_subset):
                            if len(df_subset) == 0:
                                return pd.DataFrame('', index=df_subset.index, columns=df_subset.columns)
                            styled_res = pd.DataFrame('', index=df_subset.index, columns=df_subset.columns)
                            for i in range(min(n, len(df_subset))):
                                styled_res.iloc[i, :] = 'background-color: #fff3cd'
                            return styled_res
                        return highlight_top_n

                    st.dataframe(cat_df.style.apply(make_highlight_func(limit), axis=None).format({
                        'Actual Value': '${:,.2f}', 
                        'Theoretical Value': '${:,.2f}',
                        'Variance ($)': '${:,.2f}',
                        'Variance %': '{:.2f}%',
                        'Efficiency': '{:.2%}'
                    }), use_container_width=True, hide_index=True)
                    
                    highlighted_subset = cat_df.head(limit)
                    for _, item in highlighted_subset.iterrows():
                        prod_num = item['Product Number']
                        prod_name = item['Product Name']
                        var_pct = item['Variance %']
                        
                        input_label = f"Explanation / Action Plan for: {prod_name} ({prod_num}) [{var_pct:.2f}% Var]"
                        note_key = f"note_{prod_num}"
                        
                        st.text_input(
                            input_label,
                            placeholder="Type explanation and action plan here...",
                            key=note_key
                        )
                    st.write("")
                
                submitted = st.form_submit_button("💾 Save Action Plans")

            if submitted:
                st.success("Action plans saved successfully! You can now download your focus report.")

            # --- FOCUS REPORT PDF GENERATION ---
            def generate_focus_pdf():
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=18, leftMargin=18, topMargin=18, bottomMargin=18)
                elements = []
                styles = getSampleStyleSheet()
                
                title_style = ParagraphStyle(
                    'TitleStyle',
                    parent=styles['Heading1'],
                    fontSize=11,
                    textColor=colors.HexColor('#111111'),
                    spaceAfter=3
                )
                
                heading_style = ParagraphStyle(
                    'HeadingStyle',
                    parent=styles['Heading2'],
                    fontSize=8.5,
                    textColor=colors.HexColor('#222222'),
                    spaceBefore=4,
                    spaceAfter=2
                )
                
                note_style = ParagraphStyle(
                    'NoteStyle',
                    parent=styles['Normal'],
                    fontSize=6,
                    textColor=colors.HexColor('#d9534f'),
                    spaceBefore=1,
                    spaceAfter=3,
                    leftIndent=4
                )
                
                elements.append(Paragraph(f"Variance/Efficiency Report: {store_name}", title_style))
                elements.append(Paragraph("Total Food", heading_style))
                
                pdf_summary_data = [["GL", "Category", "Actual", "Theoretical", "Variance ($)", "Var %", "Efficiency"]]
                for idx, row in summary_with_total.iterrows():
                    pdf_summary_data.append([
                        str(row['GL Code']),
                        str(row['Category']),
                        f"${row['Actual Value']:,.2f}",
                        f"${row['Theoretical Value']:,.2f}",
                        f"${row['Variance ($)']:,.2f}",
                        f"{row['Variance %']:.2f}%",
                        f"{row['Efficiency']:.2%}"
                    ])
                
                t_summary = Table(pdf_summary_data, colWidths=[46, 100, 75, 75, 80, 55, 55], hAlign='LEFT')
                t_summary.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f0f2f6')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#111111')),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,0), 7),
                    ('BOTTOMPADDING', (0,0), (-1,0), 1.5),
                    ('TOPPADDING', (0,0), (-1,0), 1.5),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#dddddd')),
                    ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                    ('FONTSIZE', (0,1), (-1,-1), 6),
                    ('BOTTOMPADDING', (0,1), (-1,-1), 1),
                    ('TOPPADDING', (0,1), (-1,-1), 1),
                    ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
                    ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#f9f9f9')),
                ]))
                elements.append(t_summary)
                elements.append(Spacer(1, 2))
                
                elements.append(Paragraph("Top Highlighted Variance Drivers & Action Plans", heading_style))
                
                for index, row in sorted_categories_for_loop.iterrows():
                    cat_name = row['Category']
                    gl_code = row['GL Code']
                    limit = highlight_limits.get(cat_name, 1)
                    
                    cat_df = df_items[df_items['GL Code'] == gl_code].copy()
                    cat_df = cat_df.sort_values(by='Variance %', ascending=True).reset_index(drop=True)
                    focus_items = cat_df.head(limit)
                    
                    if len(focus_items) > 0:
                        elements.append(Paragraph(f"<b>{cat_name} ({gl_code})</b>", ParagraphStyle('SubHeading', parent=styles['Normal'], fontSize=7, fontName='Helvetica-Bold', spaceBefore=2, spaceAfter=1)))
                        
                        cat_table_data = [["Prod #", "Product Name", "Actual", "Theoretical", "Variance ($)", "Var %", "Efficiency"]]
                        for _, item in focus_items.iterrows():
                            cat_table_data.append([
                                str(item['Product Number']),
                                str(item['Product Name']),
                                f"${item['Actual Value']:,.2f}",
                                f"${item['Theoretical Value']:,.2f}",
                                f"${item['Variance ($)']:,.2f}",
                                f"{item['Variance %']:.2f}%",
                                f"{item['Efficiency']:.2%}"
                            ])
                        
                        t_cat = Table(cat_table_data, colWidths=[70, 206, 65, 65, 65, 55, 50], hAlign='LEFT')
                        t_cat.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#fff3cd')),
                            ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#111111')),
                            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0,0), (-1,0), 6),
                            ('BOTTOMPADDING', (0,0), (-1,0), 1),
                            ('TOPPADDING', (0,0), (-1,0), 1),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e2e2')),
                            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                            ('FONTSIZE', (0,1), (-1,-1), 5.5),
                            ('BOTTOMPADDING', (0,1), (-1,-1), 0.5),
                            ('TOPPADDING', (0,1), (-1,-1), 0.5),
                        ]))
                        elements.append(t_cat)
                        
                        for _, item in focus_items.iterrows():
                            p_num = item['Product Number']
                            p_name = item['Product Name']
                            user_note = st.session_state.get(f"note_{p_num}", "").strip()
                            if user_note:
                                elements.append(Paragraph(f"<b>{p_name}:</b> {user_note}", note_style))
                        
                        elements.append(Spacer(1, 2))
                
                doc.build(elements)
                buffer.seek(0)
                return buffer.getvalue()

            # Display Focus Report PDF Download Button at the very bottom
            st.write("---")
            st.download_button(
                label="📥 Download focus report",
                data=generate_focus_pdf(),
                file_name=f"Variance_Focus_Report_{store_name.replace(' ', '_')}.pdf",
                mime="application/pdf"
            )
            
        else:
            st.error(f"Layout mismatch: The script expected 10 data columns but found {df.shape[1]}.")

    except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")
