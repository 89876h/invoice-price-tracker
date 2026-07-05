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
    Parses structural pipes to perfectly isolate Catalog Number, 
    Item Description, and Unit Price while completely stripping out 
    quantities, vendor codes, and totals.
    """
    items = {}
    lines = text_content.split('\n')
    
    inv_name = filename.split('.')[0]
    column_header = f"{inv_name} (Rate)"
    
    # Common vendor classification codes to explicitly drop
    vendor_codes = {"PVC", "PVF", "BRI", "CON", "ALF"}
    
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
            
        # Ignore structural invoice layout noise, headers, and totals
        if any(x in cleaned_line.lower() for x in [
            "sub total:", "total:", "tax:", "remit to:", "ship to:", "sold to:", 
            "visa credit:", "coppell", "waco", "farmers branch", "ticket #", 
            "original invoice", "invoice date:", "date ordered:", "date shipped:", 
            "customer job", "signed by:", "cash sale", "page:", "shipping from:", 
            "amount:", "freight:", "ship via:", "salesman:", "item number", "continued on"
        ]):
            continue

        # Process structural pipe-separated lines
        if '|' in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split('|')]
            
            if len(parts) >= 6:
                catalog_num = ""
                description = ""
                unit_price = None
                
                # 1. Isolate Catalog Number (Filter out vendor codes and raw digit rows)
                for part in parts[2:5]:
                    if part and not re.search(r"^\d+$", part) and part not in vendor_codes:
                        # Clean off any text prefixes or trailing numbers if glued together by text reader
                        cleaned_part = re.sub(r"^\d+\s*", "", part).split()[0]
                        if cleaned_part not in vendor_codes:
                            catalog_num = cleaned_part
                            break
                
                # 2. Isolate Description
                for part in parts:
                    if any(k in part.upper() for k in ["CONDUIT", "ELBOW", "COUPLING", "STRAP", "CAP", "BOX", "WIRE"]):
                        description = part.strip()
                        break
                
                # 3. Isolate Unit Price (Avoid capturing the final Extended Price column)
                for part in parts[4:-1]:
                    matches = re.findall(r"\d{1,5}\.\d{2}", part)
                    if matches:
                        test_price = clean_price_string(matches[0])
                        if test_price:
                            unit_price = test_price
                            break
                
                # Commit if valid product structural data exists
                if catalog_num and description and unit_price is not None:
                    items[(catalog_num, description)] = unit_price

    return column_header, items

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} files staged for processing.")
    
    if st.button(f"Generate Side-by-Side Comparison for {len(uploaded_files)} Invoices"):
        master_matrix = {}
        all_invoice_columns = []
        
        for file in uploaded_files:
            try:
                bytes_data = file.read()
                
                if file.name.lower().endswith('.pdf'):
                    with pdfplumber.open(io.BytesIO(bytes_data)) as pdf:
                        string_data = "\n".join([page.extract_text() or "" for page in pdf.pages])
                else:
                    string_data = bytes_data.decode("utf-8", errors="ignore")
                
                col_header, file_items = parse_invoice_content(string_data, file.name)
                
                if file_items:
                    if col_header not in all_invoice_columns:
                        all_invoice_columns.append(col_header)
                    
                    for (cat_num, item_desc), unit_price in file_items.items():
                        item_key = (cat_num, item_desc)
                        if item_key not in master_matrix:
                            master_matrix[item_key] = {}
                        master_matrix[item_key][col_header] = unit_price
            except Exception as e:
                st.error(f"Error parsing file {file.name}: {str(e)}")
                continue
        
        matrix_records = []
        for (cat_num, item_desc), tracking_cols in master_matrix.items():
            record = {
                "Catalog Number": cat_num,
                "Item Description": item_desc
            }
            for col in all_invoice_columns:
                # Forces 0.00 as a clean default if the item is missing on this specific document
                record[col] = tracking_cols.get(col, 0.00)
            matrix_records.append(record)
            
        if matrix_records:
            final_df = pd.DataFrame(matrix_records)
            final_df[all_invoice_columns] = final_df[all_invoice_columns].fillna(0.00)
            
            # Dynamic price tracking
            if len(all_invoice_columns) >= 2:
                def calculate_change(row):
                    valid_prices = [row[col] for col in all_invoice_columns if row[col] > 0.00]
                    if len(valid_prices) >= 2:
                        old, new = valid_prices[0], valid_prices[-1]
                        if new > old:
                            return f"⚠️ Up from ${old:.2f} to ${new:.2f}"
                        elif new < old:
                            return f"✅ Down from ${old:.2f} to ${new:.2f}"
                    return "Stable / New Item"
                final_df["Price Shift Audit"] = final_df.apply(calculate_change, axis=1)
            else:
                final_df["Price Shift Audit"] = "Upload more invoices to run comparative tracking."

            st.session_state["comparison_matrix"] = final_df
        else:
            st.error("No valid material lines matching structural specifications were found.")

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
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 18)
            
    st.download_button(
        label="📥 Download Clean Comparison Excel Sheet (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"Unit_Price_Comparison_Matrix_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
