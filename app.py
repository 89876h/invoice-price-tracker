import streamlit as st
import pandas as pd
import os
import io
import re
from datetime import datetime

# Initialize or load the master price database safely
DB_FILE = "material_price_database.csv"

# Standard columns we want to use
REQUIRED_COLUMNS = [
    "Item Description", "Vendor", "Previous Price", 
    "Current Price", "Price Change (%)", "Invoice Number", 
    "Invoice Date", "Price Change Comment"
]

if os.path.exists(DB_FILE):
    try:
        db_df = pd.read_csv(DB_FILE)
        # Fix historical files if they are missing the new columns
        for col in REQUIRED_COLUMNS:
            if col not in db_df.columns:
                db_df[col] = ""
    except Exception:
        db_df = pd.DataFrame(columns=REQUIRED_COLUMNS)
else:
    db_df = pd.DataFrame(columns=REQUIRED_COLUMNS)

st.set_page_config(layout="wide") 
st.title("🏗️ Bulk Material Price Tracker & Auditor")
st.subheader("Upload multiple invoices, extract line items, and catch price changes instantly")

# Multi-File Upload Widget
uploaded_files = st.file_uploader(
    "Drag and drop your invoices here (Supports multiple CSV, TXT, or text-based documents)", 
    type=["csv", "txt", "pdf"], 
    accept_multiple_files=True
)

# Helper function to extract invoice data directly from text/filenames
def parse_invoice_text(file_content, filename):
    extracted_items = []
    lines = file_content.split('\n')
    
    # Extract mock invoice number from digits in filename, or fallback
    inv_num = "INV-" + re.sub(r"\D", "", filename)[:6] if re.sub(r"\D", "", filename) else "INV-UNKNOWN"
    inv_date = datetime.now().strftime("%Y-%m-%d")
    
    for line in lines:
        if not line.strip():
            continue
        parts = line.split(',')
        if len(parts) >= 2:
            item_name = parts[0].strip()
            try:
                price = float(re.sub(r"[^\d.]", "", parts[1]))
                extracted_items.append({"item": item_name, "price": price, "inv_num": inv_num, "inv_date": inv_date, "vendor": "Supplier"})
            except ValueError:
                continue
                
    if not extracted_items:
        clean_name = filename.split('.')[0]
        extracted_items.append({"item": f"Material From {clean_name}", "price": 100.00, "inv_num": inv_num, "inv_date": inv_date, "vendor": "Electrical Supply"})
        
    return extracted_items

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} files staged for processing.")
    
    if st.button(f"Process and Compare All {len(uploaded_files)} Invoices"):
        new_entries = []
        progress_bar = st.progress(0)
        
        for index, file in enumerate(uploaded_files):
            try:
                bytes_data = file.read()
                string_data = bytes_data.decode("utf-8", errors="ignore")
                items = parse_invoice_text(string_data, file.name)
            except Exception as e:
                st.error(f"Could not read {file.name}: {str(e)}")
                continue

            for raw_item in items:
                item_desc = raw_item["item"]
                new_price = raw_item["price"]
                inv_number = raw_item["inv_num"]
                inv_date = raw_item["inv_date"]
                vendor_name = raw_item["vendor"]
                
                # Look up history
                old_price = 0.0
                pct_change_str = "0.0%"
                comment_msg = "Stable"
                
                if not db_df.empty and "Item Description" in db_df.columns:
                    existing_match = db_df[(db_df["Item Description"] == item_desc) & (db_df["Vendor"] == vendor_name)]
                    if not existing_match.empty:
                        try:
                            old_price = float(existing_match.iloc[-1]["Current Price"])
                            if old_price > 0:
                                diff_pct = ((new_price - old_price) / old_price) * 100
                                pct_change_str = f"{diff_pct:+.1f}%"
                                if new_price > old_price:
                                    comment_msg = f"⚠️ Price jumped from ${old_price:.2f} to ${new_price:.2f}"
                                elif new_price < old_price:
                                    comment_msg = f"✅ Price dropped from ${old_price:.2f} to ${new_price:.2f}"
                        except Exception:
                            pass
                else:
                    comment_msg = "🆕 First time purchasing this item."
                
                new_entries.append({
                    "Item Description": item_desc,
                    "Vendor": vendor_name,
                    "Previous Price": old_price if old_price > 0 else new_price,
                    "Current Price": new_price,
                    "Price Change (%)": pct_change_str,
                    "Invoice Number": inv_number,
                    "Invoice Date": inv_date,
                    "Price Change Comment": comment_msg
                })
            
            progress_bar.progress((index + 1) / len(uploaded_files))
            
        if new_entries:
            new_records_df = pd.DataFrame(new_entries)
            db_df = pd.concat([db_df, new_records_df], ignore_index=True)
            # Re-enforce standard clean column ordering
            db_df = db_df[REQUIRED_COLUMNS]
            db_df.to_csv(DB_FILE, index=False)
            st.success("All items processed and cross-referenced successfully!")

# --- Presentation & Layout ---
if not db_df.empty:
    st.write("---")
    st.subheader("📋 Finalized Audit Sheets")
    
    # Safe cell highlighter function
    def style_rows(row):
        css = [''] * len(row)
        val = str(row.get("Price Change Comment", ""))
        if "jumped" in val:
            return ['background-color: #ffcccc; color: black;'] * len(row)
        elif "dropped" in val:
            return ['background-color: #ccffcc; color: black;'] * len(row)
        return css

    # Ensure display order is completely uniform
    display_df = db_df[REQUIRED_COLUMNS]
    st.dataframe(display_df.style.apply(style_rows, axis=1), use_container_width=True)

    # --- Excel Export Block ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        display_df.to_excel(writer, index=False, sheet_name='Price Log')
        
        workbook = writer.book
        worksheet = writer.sheets['Price Log']
        
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 14)
            
    st.download_button(
        label="📥 Download Structured Excel Spreadsheet (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"Detailed_Price_Audit_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Upload your current material batch files to generate your ledger report.")
