import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

# Initialize or load the master price database
DB_FILE = "material_price_database.csv"
if os.path.exists(DB_FILE):
    db_df = pd.read_csv(DB_FILE)
else:
    db_df = pd.DataFrame(columns=["Date", "File Source", "Item Description", "Unit Price", "Vendor", "Price Change Comment"])

st.set_page_config(layout="wide") 
st.title("🏗️ Bulk Material Price Tracker & Reporter")
st.subheader("Upload multiple invoice files, compare history, and export to Excel")

# 1. Multi-File Upload Widget
uploaded_files = st.file_uploader(
    "Drag and drop your material invoices here (Up to 100+ files)", 
    type=["pdf", "png", "jpg", "jpeg", "csv", "txt"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} files staged for processing.")
    
    if st.button(f"Process All {len(uploaded_files)} Invoices"):
        current_date = datetime.now().strftime("%Y-%m-%d")
        new_records = []
        
        progress_bar = st.progress(0)
        
        for index, file in enumerate(uploaded_files):
            # --- SIMULATED AI EXTRACTION LOGIC ---
            # In production, this parses text using an OCR or LLM API
            simulated_vendor = "Apex Material Supplies"
            simulated_item = f"Material Item Type {index % 6}" # Simulates recurring items
            simulated_price = 15.00 + (index * 0.75 % 5.00)     # Simulates price changes
            # -------------------------------------

            # Database Comparison
            previous_purchases = db_df[(db_df["Item Description"] == simulated_item) & (db_df["Vendor"] == simulated_vendor)]
            
            comment = ""
            if not previous_purchases.empty:
                old_price = float(previous_purchases.iloc[-1]["Unit Price"])
                
                if simulated_price > old_price:
                    diff = ((simulated_price - old_price) / old_price) * 100
                    comment = f"⚠️ Price INCREASED by {diff:.1f}% (Was ${old_price:.2f})"
                elif simulated_price < old_price:
                    diff = ((old_price - simulated_price) / old_price) * 100
                    comment = f"✅ Price DECREASED by {diff:.1f}% (Was ${old_price:.2f})"
                else:
                    comment = "Stable" # No change notification
            else:
                comment = "🆕 New item added"

            new_records.append({
                "Date": current_date,
                "File Source": file.name,
                "Item Description": simulated_item,
                "Unit Price": simulated_price,
                "Vendor": simulated_vendor,
                "Price Change Comment": comment
            })
            
            progress_bar.progress((index + 1) / len(uploaded_files))
            
        if new_records:
            new_df = pd.DataFrame(new_records)
            db_df = pd.concat([db_df, new_df], ignore_index=True)
            db_df.to_csv(DB_FILE, index=False)
            st.success(f"Successfully processed all {len(uploaded_files)} files!")

# --- Data View & Export Section ---
if not db_df.empty:
    st.write("---")
    st.subheader("📊 Price Logs & Reporting Database")
    
    tab1, tab2 = st.tabs(["All Checked Items", "Price Fluctuations Only"])
    
    def highlight_rows(val):
        if "INCREASED" in str(val):
            return 'background-color: #ffcccc; color: black;'
        elif "DECREASED" in str(val):
            return 'background-color: #ccffcc; color: black;'
        return ''

    with tab1:
        st.dataframe(db_df.style.applymap(highlight_rows, subset=['Price Change Comment']), use_container_width=True)
        
    with tab2:
        filtered_df = db_df[db_df["Price Change Comment"].str.contains("⚠️|✅", na=False)]
        if not filtered_df.empty:
            st.dataframe(filtered_df.style.applymap(highlight_rows, subset=['Price Change Comment']), use_container_width=True)
        else:
            st.info("No items with price updates found yet.")

    # --- Excel Generation Button ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        db_df.to_excel(writer, index=False, sheet_name='Price Master Log')
        
        # Pull worksheet to apply basic formatting
        workbook = writer.book
        worksheet = writer.sheets['Price Master Log']
        # Autofit column widths dynamically
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    st.download_button(
        label="📥 Download Excel Report (.xlsx)",
        data=buffer.getvalue(),
        file_name=f"Price_Analysis_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.write("Database is empty. Upload your first batch of invoices to begin.")
