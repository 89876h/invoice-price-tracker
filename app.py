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
    Robust line-by-line parser that captures ALL material items without skipping.
    Cleans up descriptions and matches them with their true unit rates.
    """
    items = {}
    lines = text_content.split('\n')
    
    inv_name = filename.split('.')[0]
    column_header = f"{inv_name} (Rate)"
    
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
            
        # Ignore structural invoice metadata, totals, and location information
        if any(x in cleaned_line.lower() for x in [
            "sub total:", "total:", "tax:", "remit to:", "ship to:", "sold to:", 
            "visa credit:", "coppell", "waco", "farmers branch", "ticket #", 
            "original invoice", "invoice date:", "date ordered:", "date shipped:", 
            "customer job", "signed by:", "cash sale", "page:", "shipping from:", 
            "amount:", "freight:", "ship via:", "salesman:", "item number", "continued on"
        ]):
            continue

        # Handle structural pipe-separated lines (e.g., Elliott Electric)
        if '|' in cleaned_line:
            parts = [p.strip() for p in cleaned_line.split('|')]
            
            # A valid material line typically splits into 5+ segments including descriptions and rates
            if len(parts) >= 4:
                # Find the description field (usually the long text segment in the middle)
                desc = ""
                for part in parts:
                    # Filter out short tracking codes, quantities, and flags
                    if len(part) > 5 and not re.search(r"^\d+$", part) and not part.startswith("$"):
                        desc = part
                        break
                
                if desc:
                    price = None
                    # Search through columns for the true decimal unit rate
                    for part in parts:
                        matches = re.findall(r"\d{1,5}\.\d{2}", part)
                        # Ensure we don't grab the 'Extended Price' total at the very end of the row
                        if matches and part != parts[-1]:
                            test_price = clean_price_string(matches[0])
                            if test_price:
                                price = test_price
                                break
                    
                    if price is not None:
                        # Clean off leading quantity data prefixes from the text string
                        desc_clean = re.sub(r"^[A-Za-z0-9\/_\-]+\s+", "", desc)
                        items[desc_clean.strip()] = price
            continue

        # Fallback for standard text or comma-separated tables
        if ',' in cleaned_line:
            parts = cleaned_line.split(',')
            if len(parts) >= 2:
                desc = parts[0].strip()
                price = clean_price_string(parts[1])
                if desc and price:
                    items[desc] = price
                    continue

        # Pattern lookup for plain text row arrangements
        match = re.search(r"([A-Za-z0-9\"'\/\s\-]{6,})\s+\$?\s*(\d{1,5}\.\d{2})", cleaned_line)
        if match:
            desc_clean = match.group(1).strip()
            price = clean_price_string(match.group(2))
            if desc_clean and price:
                items[desc_clean] = price

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
                    
                    for item_desc, unit_price in file_items.items():
                        if item_desc not in master_matrix:
                            master_matrix[item_desc] = {}
                        master_matrix[item_desc][col_header] = unit_price
            except Exception as e:
                st.error(f"Error parsing file {file.name}: {str(e)}")
                continue
        
        matrix_records = []
        for item_desc, tracking_cols in master_matrix.items():
            record = {"Item Description": item_desc}
            for col in all_invoice_columns:
                # Fallback to 0.00 if the invoice is completely missing this specific item
                record[col] = tracking_cols.get(col, 0.00)
            matrix_records.append(record)
            
        if matrix_records:
            final_df = pd.DataFrame(matrix_records)
            
            # Replace any hidden NaN artifacts explicitly with 0.00
            final_df[all_invoice_columns] = final_df[all_invoice_columns].fillna(0.00)
            
            # Dynamically calculate price movements across the sheet
            if len(all_invoice_columns) >= 2:
                def calculate_change(row):
                    # Filter out items that are 0.00 (not present on that specific invoice)
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
            st.error("No valid material line items or price entries could be compiled.")

if "comparison_matrix" in st.session_state:
    df_to_show = st.session_state["comparison_matrix"]
    st.write("---")
    st.subheader("📊 Unit Price Dynamic Grid")
    
    # Format grid view to display crisp currency numbers
    st.dataframe(df_to_show, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_to_show.to_excel(writer, index=False, sheet_name='Price Comparison Grid')
        workbook = writer.book
        worksheet = writer.sheets['Price Comparison Grid']
        
        # Apply standard column padding formatting for clean Excel outputs
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 22)
            
    st.download_button(
        label="📥 Download Clean Comparison Excel Sheet (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"Unit_Price_Comparison_Matrix_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
