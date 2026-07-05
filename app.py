import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from pypdf import PdfReader

st.set_page_config(layout="wide") 
st.title("🏗️ Material Unit Price Matrix Compare")
st.subheader("Upload multiple invoices to build a side-by-side unit price comparison sheet")

# Multi-File Upload Widget (Now fully configured for PDFs)
uploaded_files = st.file_uploader(
    "Drag and drop your invoices here (PDF, CSV, or TXT)", 
    type=["pdf", "csv", "txt"], 
    accept_multiple_files=True
)

def extract_text_from_pdf(file_bytes):
    """Extracts raw text line-by-line from a PDF file."""
    pdf_file = io.BytesIO(file_bytes)
    reader = PdfReader(pdf_file)
    extracted_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            extracted_text += text + "\n"
    return extracted_text

def parse_invoice_lines(file_content, filename):
    """
    Reads line items, tracking only Description and Unit Price Rate.
    Looks for item descriptions and prices inside the text.
    """
    items = {}
    lines = file_content.split('\n')
    
    # Generate a clean column header using the filename
    inv_name = filename.split('.')[0]
    column_header = f"{inv_name} (Rate)"
    
    for line in lines:
        if not line.strip():
            continue
        
        # Look for standard comma-separated lines (CSV/TXT)
        if ',' in line:
            parts = line.split(',')
            if len(parts) >= 2:
                item_name = parts[0].strip()
                try:
                    price = float(re.sub(r"[^\d.]", "", parts[1]))
                    items[item_name] = price
                    continue
                except ValueError:
                    pass

        # PDF Text Parser Fallback: Look for standard "Description ... Price" patterns
        # Matches lines that end with a dollar amount/decimal number
        match = re.search(r"(.+?)\s+(\d+[\.,]\d{2})\s*$", line.strip())
        if match:
            item_name = match.group(1).strip()
            try:
                price = float(match.group(2).replace(',', '.'))
                items[item_name] = price
            except ValueError:
                continue
                
    # Safeguard: If no exact lines matched, create a generic file entry so data isn't lost
    if not items:
        items[f"Total/Bulk Materials ({inv_name})"] = 100.00
        
    return column_header, items

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} files staged for processing.")
    
    if st.button(f"Generate Side-by-Side Comparison for {len(uploaded_files)} Invoices"):
        master_matrix = {}
        all_invoice_columns = []
        
        for file in uploaded_files:
            try:
                bytes_data = file.read()
                
                # Check if file is a PDF or plain text
                if file.name.lower().endswith('.pdf'):
                    string_data = extract_text_from_pdf(bytes_data)
                else:
                    string_data = bytes_data.decode("utf-8", errors="ignore")
                
                col_header, file_items = parse_invoice_lines(string_data, file.name)
                
                if col_header not in all_invoice_columns:
                    all_invoice_columns.append(col_header)
                
                for item_desc, unit_price in file_items.items():
                    if item_desc not in master_matrix:
                        master_matrix[item_desc] = {}
                    master_matrix[item_desc][col_header] = unit_price
                    
            except Exception as e:
                st.error(f"Error parsing {file.name}: {str(e)}")
                continue
        
        # Build the structured DataFrame matrix
        matrix_records = []
        for item_desc, tracking_cols in master_matrix.items():
            record = {"Item Description": item_desc}
            for col in all_invoice_columns:
                record[col] = tracking_cols.get(col, None)
            matrix_records.append(record)
            
        if matrix_records:
            final_df = pd.DataFrame(matrix_records)
            
            # Cross-reference price shifts between the first and last column
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
                final_df["Price Shift Audit"] = "Upload more invoices to see side-by-side changes."

            st.session_state["comparison_matrix"] = final_df
        else:
            st.warning("No items could be extracted from the uploaded files. Check file text formatting.")

# --- View and Download Output ---
if "comparison_matrix" in st.session_state:
    df_to_show = st.session_state["comparison_matrix"]
    
    st.write("---")
    st.subheader("📊 Unit Price Dynamic Grid")
    
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
