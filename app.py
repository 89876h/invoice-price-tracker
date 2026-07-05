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

def clean_price_string(price_str):
    """Cleans up fragmented price segments from split columns."""
    cleaned = re.sub(r"[^\d.]", "", price_str.strip())
    try:
        val = float(cleaned)
        if 0.01 <= val <= 100000.00:
            return val
    except ValueError:
        pass
    return None

def parse_invoice_content(text_content, filename):
    """
    Robust line-by-line parser tailored for complex multi-column layouts.
    Extracts explicit material descriptions and unit rates.
    """
    items = {}
    lines = text_content.split('\n')
    
    inv_name = filename.split('.')[0]
    column_header = f"{inv_name} (Rate)"
    
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
            
        # Ignore obvious headers, metadata, and totals
        if any(x in cleaned_line.lower() for x in ["sub total:", "total:", "tax:", "remit to:", "ship to:", "sold to:", "page:"]):
            continue

        # Handle pipe-separated layouts (like Elliott Electric)
        if '|' in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split('|')]
            
            # Find the segment containing descriptions (typically long string blocks)
            desc_part = ""
            for p in parts:
                if any(m in p.upper() for m in ["PVC", "CONDUIT", "ELBOW", "COUPLING", "STRAP", "CAP", "BOX", "WIRE"]):
                    desc_part = p
                    break
            
            if desc_part:
                # Look for a valid decimal rate within the adjacent columns
                for p in parts:
                    # Clean out noise characters to check for standalone numbers
                    digits_only = re.sub(r"[^\d.]", "", p)
                    if digits_only and '.' in digits_only:
                        price = clean_price_string(p)
                        if price is not None:
                            # Avoid misidentifying the long Extended Price at the far right
                            # Unit prices typically map to smaller amounts before totals
                            items[desc_part] = price
                            break
            continue

        # Fallback for standard comma-separated lines (CSV/TXT)
        if ',' in cleaned_line:
            parts = cleaned_line.split(',')
            if len(parts) >= 2:
                desc = parts[0].strip()
                price = clean_price_string(parts[1])
                if desc and price:
                    items[desc] = price
                    continue

        # Fallback regex for standard text line structures
        match = re.search(r"(.+?)\s+(\d+[\.,]\d{2})\s*$", cleaned_line)
        if match:
            desc = match.group(1).strip()
            price = clean_price_string(match.group(2))
            if desc and price and len(desc) > 3:
                items[desc] = price

    return column_header, items

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} files staged for processing.")
    
    if st.button(f"Generate Side-by-Side Comparison for {len(uploaded_files)} Invoices"):
        master_matrix = {}
        all_invoice_columns = []
        
        for file in uploaded_files:
            try:
                bytes_data = file.read()
                
                # Extract text reliably from either PDF or text files
                if file.name.lower().endswith('.pdf'):
                    with pdfplumber.open(io.BytesIO(bytes_data)) as pdf:
                        string_data = "\n".join([page.extract_text() or "" for page in pdf.pages])
                else:
                    string_data = bytes_data.decode("utf-8", errors="ignore")
                
                col_header, file_items = parse_invoice_content(string_data, file.name)
                
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
        
        # Format the matrix data into clean row-by-row outputs
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
                final_df["Price Shift Audit"] = "Upload more invoices to compare historical tracking."

            st.session_state["comparison_matrix"] = final_df
        else:
            st.error("No valid material lines or unit price structures could be parsed. Check text layouts.")

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
