import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from pypdf import PdfReader

st.set_page_config(layout="wide") 
st.title("🏗️ Material Unit Price Matrix Compare")
st.subheader("Upload multiple invoices (including PDFs) to build a side-by-side unit price comparison sheet")

# Multi-File Upload Widget (Accepts PDF, CSV, and TXT)
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
    Reads every line item from text or PDF, extracting the exact 
    Material Description and its corresponding Unit Price Rate.
    """
    items = {}
    lines = file_content.split('\n')
    
    # Generate a clean column header using the filename
    inv_name = filename.split('.')[0]
    column_header = f"{inv_name} (Rate)"
    
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
        
        # Scenario A: Handle standard Comma-Separated structured lines (CSV/TXT)
        if ',' in cleaned_line:
            parts = cleaned_line.split(',')
            if len(parts) >= 2:
                item_name = parts[0].strip()
                price_raw = parts[1].strip()
                try:
                    price = float(re.sub(r"[^\d.]", "", price_raw))
                    if item_name and price > 0:
                        items[item_name] = price
                        continue
                except ValueError:
                    pass

        # Scenario B: Handle PDF/Unstructured Lines (e.g., "3/4 IN EMT Conduit  12.50")
        # Captures the text description on the left, and a trailing decimal number on the right
        match = re.search(r"(.+?)\s+(\d+[\.,]\d{2})\s*$", cleaned_line)
        if match:
            item_name = match.group(1).strip()
            price_raw = match.group(2).replace(',', '.')
            try:
                price = float(price_raw)
                # Filter out system artifacts or page numbers
                if item_name and price > 0 and len(item_name) > 2:
                    items[item_name] = price
            except ValueError:
                continue
                
    return column_header, items

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} files staged for processing.")
    
    if st.button(f"Generate Side-by-Side Comparison for {len(uploaded_files)} Invoices"):
        # Master dictionary: { "Item Description": { "Invoice_1 (Rate)": 12.50 } }
        master_matrix = {}
        all_invoice_columns = []
        
        for file in uploaded_files:
            try:
                bytes_data = file.read()
                
                # Step 1: Handle text extraction based on file format
                if file.name.lower().endswith('.pdf'):
                    string_data = extract_text_from_pdf(bytes_data)
                else:
                    string_data = bytes_data.decode("utf-8", errors="ignore")
                
                # Step 2: Parse out the descriptions and unit prices
                col_header, file_items = parse_invoice_lines(string_data, file.name)
                
                if file_items:
                    if col_header not in all_invoice_columns:
                        all_invoice_columns.append(col_header)
                    
                    # Step 3: Insert item description as a separate row row in the matrix
                    for item_desc, unit_price in file_items.items():
                        if item_desc not in master_matrix:
                            master_matrix[item_desc] = {}
                        master_matrix[item_desc][col_header] = unit_price
            except Exception as e:
                st.error(f"Error parsing file {file.name}: {str(e)}")
                continue
        
        # Build layout rows from the matrix dictionary
        matrix_records = []
        for item_desc, tracking_cols in master_matrix.items():
            record = {"Item Description": item_desc}
            for col in all_invoice_columns:
                record[col] = tracking_cols.get(col, None)
            matrix_records.append(record)
            
        if matrix_records:
            final_df = pd.DataFrame(matrix_records)
            
            # Cross-reference price shifts between columns
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
            st.error("Could not find or extract distinct item descriptions and unit prices from these files.")

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
