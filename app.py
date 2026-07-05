import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(layout="wide") 
st.title("🏗️ Material Unit Price Matrix Compare")
st.subheader("Upload multiple invoices to build a side-by-side unit price comparison sheet")

# Multi-File Upload Widget
uploaded_files = st.file_uploader(
    "Drag and drop your invoices here (CSV or TXT format)", 
    type=["csv", "txt"], 
    accept_multiple_files=True
)

def parse_invoice_lines(file_content, filename):
    """
    Reads every line item systematically from the invoice data, 
    tracking only the exact Material Description and the Unit Price Rate.
    """
    items = {}
    lines = file_content.split('\n')
    
    # Generate a clean column header using the invoice file name
    inv_name = filename.split('.')[0]
    column_header = f"{inv_name} (Rate)"
    
    for line in lines:
        if not line.strip():
            continue
        
        # Split by comma to separate the description and the price
        parts = line.split(',')
        if len(parts) >= 2:
            # The first part is the item name/description
            item_name = parts[0].strip()
            
            # Clean up the second part to extract the exact unit price
            price_raw = parts[1].strip()
            try:
                # Remove dollar signs or spaces, keeping only numbers and decimals
                price = float(re.sub(r"[^\d.]", "", price_raw))
                
                # Only save if we have a valid name and price
                if item_name and price > 0:
                    items[item_name] = price
            except ValueError:
                continue
                
    return column_header, items

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} files staged for processing.")
    
    if st.button(f"Generate Side-by-Side Comparison for {len(uploaded_files)} Invoices"):
        # Master dictionary: { "Exact Item Description": { "Invoice_1 (Rate)": 12.50, "Invoice_2 (Rate)": 13.00 } }
        master_matrix = {}
        all_invoice_columns = []
        
        # Loop through every single file uploaded
        for file in uploaded_files:
            try:
                bytes_data = file.read()
                string_data = bytes_data.decode("utf-8", errors="ignore")
                
                col_header, file_items = parse_invoice_lines(string_data, file.name)
                
                if file_items:
                    if col_header not in all_invoice_columns:
                        all_invoice_columns.append(col_header)
                    
                    # Map every single material line into our matrix
                    for item_desc, unit_price in file_items.items():
                        if item_desc not in master_matrix:
                            master_matrix[item_desc] = {}
                        master_matrix[item_desc][col_header] = unit_price
            except Exception as e:
                st.error(f"Error parsing file {file.name}: {str(e)}")
                continue
        
        # Re-compile the matrix into individual separate rows
        matrix_records = []
        for item_desc, tracking_cols in master_matrix.items():
            record = {"Item Description": item_desc}
            # Fill in the prices or leave blank if that invoice didn't contain the item
            for col in all_invoice_columns:
                record[col] = tracking_cols.get(col, None)
            matrix_records.append(record)
            
        if matrix_records:
            final_df = pd.DataFrame(matrix_records)
            
            # Calculate the difference if there are multiple columns to cross-reference
            if len(all_invoice_columns) >= 2:
                def calculate_change(row):
                    valid_prices = [row[col] for col in all_invoice_columns if pd.notnull(row[col])]
                    if len(valid_prices) >= 2:
                        old, new = valid_prices[0], valid_prices[-1]
                        if new > old:
                            return f"⚠️ Price went UP from ${old:.2f} to ${new:.2f}"
                        elif new < old:
                            return f"✅ Price went DOWN from ${old:.2f} to ${new:.2f}"
                    return "Stable / No Change"
                
                final_df["Price Shift Audit"] = final_df.apply(calculate_change, axis=1)
            else:
                final_df["Price Shift Audit"] = "Upload more invoices to see historical comparison."
                
            st.session_state["comparison_matrix"] = final_df
        else:
            st.error("No valid material descriptions or unit prices could be parsed. Check your data rows.")

# --- View and Download Output ---
if "comparison_matrix" in st.session_state:
    df_to_show = st.session_state["comparison_matrix"]
    
    st.write("---")
    st.subheader("📊 Unit Price Dynamic Grid")
    
    # Apply row coloring based on price changes
    def highlight_matrix(row):
        css = [''] * len(row)
        audit_val = str(row.get("Price Shift Audit", ""))
        if "UP" in audit_val:
            return ['background-color: #ffcccc; color: black;'] * len(row)
        elif "DOWN" in audit_val:
            return ['background-color: #ccffcc; color: black;'] * len(row)
        return css

    st.dataframe(df_to_show.style.apply(highlight_matrix, axis=1), use_container_width=True)

    # --- Excel File Export Block ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_to_show.to_excel(writer, index=False, sheet_name='Price Comparison Grid')
        
        workbook = writer.book
        worksheet = writer.sheets['Price Comparison Grid']
        
        # Clean column widths
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 16)
            
    st.download_button(
        label="📥 Download Clean Comparison Excel Sheet (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"Unit_Price_Comparison_Matrix_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
