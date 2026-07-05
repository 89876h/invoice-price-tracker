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

def parse_invoice_content(text_content, filename):
    """
    Normalizes and splits lines into predictable tokens.
    Extracts Catalog Number, Description, and Unit Price completely layout-free.
    """
    items = {}
    lines = text_content.split('\n')
    
    inv_name = filename.split('.')[0]
    column_header = f"{inv_name} (Rate)"
    
    # Standard vendor codes to automatically drop from the catalog field
    vendor_codes = {"PVC", "PVF", "BRI", "CON", "ALF", "LIT"}
    
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
            
        # Ignore invoice metadata, locations, and totals
        if any(x in cleaned_line.lower() for x in [
            "sub total:", "total:", "tax:", "remit to:", "ship to:", "sold to:", 
            "visa credit:", "coppell", "waco", "farmers branch", "ticket #", 
            "original invoice", "invoice date:", "date ordered:", "date shipped:", 
            "customer job", "signed by:", "cash sale", "page:", "shipping from:", 
            "amount:", "freight:", "ship via:", "salesman:", "item number", "continued on"
        ]):
            continue

        # 1. Normalize line layout by replacing pipes with spaces and stripping dollar signs
        normalized = cleaned_line.replace('|', ' ').replace('$', ' ')
        tokens = [t.strip() for t in normalized.split() if t.strip()]
        
        # A valid item line must have an item#, qty, backorder, catalog, vendor code, description, and price
        if len(tokens) >= 6:
            # Find all decimal prices in the tokens list
            decimal_indices = [i for i, t in enumerate(tokens) if re.match(r"^\d{1,5}\.\d{2}$", t)]
            
            if decimal_indices:
                # The first decimal match is always our Unit Price
                unit_price_idx = decimal_indices[0]
                try:
                    unit_price = float(tokens[unit_price_idx])
                except ValueError:
                    continue
                
                # Look backwards from the unit price to identify structural segments
                # Positions relative to standard invoice token sequences:
                # tokens[0]=Item#, tokens[1]=Qty, tokens[2]=BO, tokens[3]=Catalog#, tokens[4]=VendorCode
                
                catalog_num = ""
                # Determine catalog number position safely near the beginning of the tokens list
                for idx in [3, 2, 4]:
                    if idx < len(tokens) and tokens[idx] not in vendor_codes and not re.match(r"^\d+$", tokens[idx]):
                        catalog_num = tokens[idx]
                        break
                
                # Fallback if catalog number is glued directly to the backorder 0 (e.g. "0PVC1")
                if not catalog_num and len(tokens) > 2:
                    for idx in [2, 3]:
                        if idx < len(tokens):
                            m = re.match(r"^\d+([A-Za-z0-9\-]+)$", tokens[idx])
                            if m and m.group(1) not in vendor_codes:
                                catalog_num = m.group(1)
                                break
                
                # Isolate the remaining tokens between the vendor code and the unit price as the item description
                desc_tokens = []
                for t in tokens[3:unit_price_idx]:
                    if t not in vendor_codes and t != catalog_num and not re.match(r"^\d+$", t):
                        desc_tokens.append(t)
                
                description = " ".join(desc_tokens).strip()
                
                # Commit item row data if we have a valid catalog number and description
                if catalog_num and description:
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
                # Add 0.00 fallback if item is missing on specific invoice columns
                record[col] = tracking_cols.get(col, 0.00)
            matrix_records.append(record)
            
        if matrix_records:
            final_df = pd.DataFrame(matrix_records)
            final_df[all_invoice_columns] = final_df[all_invoice_columns].fillna(0.00)
            
            # Audit changes across items
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
