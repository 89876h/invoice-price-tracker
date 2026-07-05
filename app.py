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
    Scans invoice text line by line using specific patterns to find item lines.
    Extracts the clean material description and its unit price.
    """
    items = {}
    lines = text_content.split('\n')
    
    inv_name = filename.split('.')[0]
    column_header = f"{inv_name} (Rate)"
    
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
            
        # Skip summary totals or layout noise
        if any(x in cleaned_line.lower() for x in ["sub total:", "total:", "tax:", "remit to:", "ship to:", "sold to:", "visa credit:"]):
            continue

        # Look for typical item lines containing numbers and material keywords
        # Example: '1 | 60 | 0 PVC1 | PVC | SCH 40 1" 10' PVC CONDUIT | $ 40.88 C'
        # Example: '15 | 8 | 0 CPL1 | PVF | 1" PVC COUPLING | C 35.00 $'
        if any(keyword in cleaned_line.upper() for keyword in ["CONDUIT", "ELBOW", "COUPLING", "STRAP", "CAP", "PVC", "PVF", "BRI"]):
            
            # 1. Clean out the pipe characters to look at the whole string
            segments = [s.strip() for s in cleaned_line.split('|') if s.strip()]
            
            # Find the segment that represents the description text
            desc = ""
            for seg in segments:
                if any(k in seg.upper() for k in ["CONDUIT", "ELBOW", "COUPLING", "STRAP", "CAP"]):
                    desc = seg
                    break
            
            if desc:
                # 2. Look for the decimal price near the description inside the line segments
                price = None
                for seg in segments:
                    # Find numbers with a decimal point (e.g., 40.88, 35.00, 1258.94)
                    matches = re.findall(r"\d{1,4}\.\d{2}", seg)
                    if matches:
                        try:
                            # Use the first valid decimal match in the segment
                            test_price = float(matches[0])
                            # Filter out extended totals or quantities if they leak in
                            if 0.01 <= test_price <= 50000.00:
                                price = test_price
                        except ValueError:
                            continue
                
                if price is not None:
                    items[desc] = price
                    continue

        # Fallback for plain CSV/TXT formats
        if ',' in cleaned_line:
            parts = cleaned_line.split(',')
            if len(parts) >= 2:
                desc = parts[0].strip()
                try:
                    price_val = float(re.sub(r"[^\d.]", "", parts[1]))
                    if desc and price_val:
                        items[desc] = price_val
                except ValueError:
                    pass

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
                        # Extract clean plain layout text from all pages
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
        
        # Structure parsed keys into rows
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
            st.error("No valid material lines could be compiled. Please check formatting rules.")

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
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 20)
            
    st.download_button(
        label="📥 Download Clean Comparison Excel Sheet (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"Unit_Price_Comparison_Matrix_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
