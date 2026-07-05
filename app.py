import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import pdfplumber

st.set_page_config(layout="wide") 
st.title("🏗️ Material Unit Price Matrix Compare")
st.subheader("Upload multiple invoices to build a side-by-side unit price comparison sheet")

uploaded_files = st.file_uploader(
    "Drag and drop your invoices here (PDF, CSV, or TXT)", 
    type=["pdf", "csv", "txt"], 
    accept_multiple_files=True
)

def clean_price(val):
    """Safely extracts a valid decimal unit price and filters out non-price numbers."""
    if pd.isna(val) or val is None:
        return None
    cleaned = re.sub(r"[^\d.]", "", str(val).strip())
    if not cleaned or cleaned.count('.') > 1:
        return None
    try:
        price = float(cleaned)
        # Filter out random large numbers like phone numbers or tracking IDs
        if 0.05 <= price <= 50000.00:
            return price
    except ValueError:
        pass
    return None

def parse_pdf_tables(file_bytes):
    """Extracts structured table rows directly using pdfplumber."""
    extracted_items = {}
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Filter out empty rows or headers
                    if not row or len(row) < 2:
                        continue
                    
                    # Clean up descriptions and filter out short layout artifacts
                    desc = str(row[0]).strip() if row[0] else ""
                    if len(desc) < 4 or any(x in desc.lower() for x in ["total", "subtotal", "invoice", "date", "phone", "fax", "ship to", "bill to", "visa", "page"]):
                        continue
                        
                    # Search remaining columns to find a realistic unit price rate
                    for cell in row[1:]:
                        price = clean_price(cell)
                        if price is not None:
                            extracted_items[desc] = price
                            break # Found the rate column for this row
                            
    return extracted_items

def parse_text_lines(file_content):
    """Fallback parser for standard CSV or TXT rows."""
    items = {}
    for line in file_content.split('\n'):
        if ',' in line:
            parts = line.split(',')
            if len(parts) >= 2:
                desc = parts[0].strip()
                price = clean_price(parts[1])
                if desc and price:
                    items[desc] = price
    return items

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} files staged for processing.")
    
    if st.button(f"Generate Side-by-Side Comparison for {len(uploaded_files)} Invoices"):
        master_matrix = {}
        all_invoice_columns = []
        
        for file in uploaded_files:
            try:
                bytes_data = file.read()
                inv_name = file.name.split('.')[0]
                col_header = f"{inv_name} (Rate)"
                
                if file.name.lower().endswith('.pdf'):
                    file_items = parse_pdf_tables(bytes_data)
                else:
                    string_data = bytes_data.decode("utf-8", errors="ignore")
                    file_items = parse_text_lines(string_data)
                
                if file_items:
                    if col_header not in all_invoice_columns:
                        all_invoice_columns.append(col_header)
                    
                    for item_desc, unit_price in file_items.items():
                        if item_desc not in master_matrix:
                            master_matrix[item_desc] = {}
                        master_matrix[item_desc][col_header] = unit_price
            except Exception as e:
                st.error(f"Error parsing file {file.name}: {str(e)}")
                continue
        
        # Build side-by-side rows
        matrix_records = []
        for item_desc, tracking_cols in master_matrix.items():
            record = {"Item Description": item_desc}
            for col in all_invoice_columns:
                record[col] = tracking_cols.get(col, None)
            matrix_records.append(record)
            
        if matrix_records:
            final_df = pd.DataFrame(matrix_records)
            
            if len(all_invoice_columns) >= 2:
                def calculate_change(row):
                    valid_prices = [row[col] for col in all_invoice_columns if pd.notnull(row[col])]
                    if len(valid_prices) >= 2:
                        old, new = valid_prices[0], valid_prices[-1]
                        if new > old:
                            return f"⚠️ Up from ${old:.2f} to ${new:.2f}"
                        elif new < old:
                            return f"✅ Down from ${old:.2f} to ${new:.2f}"
                    return "Stable"
                final_df["Price Shift Audit"] = final_df.apply(calculate_change, axis=1)
            else:
                final_df["Price Shift Audit"] = "Upload more invoices to compare."

            st.session_state["comparison_matrix"] = final_df
        else:
            st.error("No valid tabular material descriptions or unit prices could be detected.")

if "comparison_matrix" in st.session_state:
    df_to_show = st.session_state["comparison_matrix"]
    st.write("---")
    st.subheader("📊 Unit Price Dynamic Grid")
    
    st.dataframe(df_to_show, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_to_show.to_excel(writer, index=False, sheet_name='Price Comparison Grid')
        workbook = writer.book
        worksheet = writer.sheets['Price Comparison Grid']
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
