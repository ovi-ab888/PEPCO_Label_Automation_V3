import streamlit as st
import fitz
import pandas as pd
import re
from datetime import datetime
from io import StringIO
import csv as pycsv
import os

st.set_page_config(page_title="PEPCO Label Automation V2", layout="wide")

st.title("PEPCO Label Automation V2")

# =============================
# Extract General Information (Old file structure)
# =============================
def extract_general_data(text):
    # Order ID extraction (improved pattern from old file)
    order_id = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)", text, re.IGNORECASE)
    
    # Style extraction (6 digit number)
    style = re.search(r"\b\d{6}\b", text)
    
    # Supplier product code
    supplier_code = re.search(r"Supplier product code\s*\.{2,}\s*(.+)", text, re.IGNORECASE)
    
    # Item classification
    item_class = re.search(r"Item classification\s*\.{2,}\s*(.+)", text, re.IGNORECASE)
    
    # Supplier name
    supplier = re.search(r"Supplier name\s*\.{2,}\s*(.+)", text, re.IGNORECASE)
    
    return {
        "Order_ID": order_id.group(1).strip() if order_id else "",
        "Style": style.group(0) if style else "",
        "Supplier_product_code": supplier_code.group(1).strip() if supplier_code else "",
        "Item_classification": item_class.group(1).strip() if item_class else "",
        "Supplier_name": supplier.group(1).strip() if supplier else "",
        "today_date": datetime.today().strftime('%d-%m-%Y')
    }

# =============================
# Extract Label Data (New file structure)
# =============================
def extract_label_data(text):
    # TC Number
    tc = re.search(r"TC\s*-\s*(T\d+)", text, re.IGNORECASE)
    if not tc:
        tc = re.search(r"TC\s*[:.]?\s*(T\d+)", text, re.IGNORECASE)
    
    # Product name
    product = re.search(r"ITEM\s*\d+\s*\n\s*(.+)", text, re.IGNORECASE)
    if not product:
        product = re.search(r"Product\s*name\s*[:.]?\s*(.+)", text, re.IGNORECASE)
    
    # Barcode (13 digits)
    barcode = re.search(r"\b\d{13}\b", text)
    
    # Inner quantity
    inner_qty = re.search(r"(\d+)\s*Pcs", text, re.IGNORECASE)
    
    # Outer quantity
    outer_qty = re.search(r"(\d+)\s*Inner", text, re.IGNORECASE)
    if not outer_qty:
        outer_qty = re.search(r"Outer\s*qty\s*[:.]?\s*(\d+)", text, re.IGNORECASE)
    
    # Inner kg
    inner_kg = re.search(r"MAX\.?\s*(\d+)\s*kg", text, re.IGNORECASE)
    if not inner_kg:
        inner_kg = re.search(r"(\d+)\s*kg", text, re.IGNORECASE)
    
    # Season
    season = re.search(r"\b(AW|SS|FW|SW)\d{2}\b", text, re.IGNORECASE)
    
    return {
        "TC_Number": tc.group(1) if tc else "",
        "Product_name": product.group(1).strip() if product else "",
        "Barcode": barcode.group(0) if barcode else "",
        "Inner_kg": f"MAX. {inner_kg.group(1)} kg" if inner_kg else "",
        "Season": season.group(0).upper() if season else "",
        "Inner_qty": f"{inner_qty.group(1)} Pcs" if inner_qty else "",
        "Outer_qty": f"{outer_qty.group(1)} Inner" if outer_qty else ""
    }

# =============================
# Extract Colour (Improved from old file)
# =============================
def extract_colour(text, page_number=1):
    try:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        
        # Keywords to skip (from old file)
        skip_keywords = [
            "PURCHASE", "COLOUR", "TOTAL", "PANTONE", "SUPPLIER", "PRICE",
            "ORDERED", "SIZES", "TPG", "TPX", "USD", "NIP", "PEPCO",
            "Poland", "ul. Strzeszyńska 73A, 60-479 Poznań", "NIP 782-21-31-157",
            "BARCODE", "INNER", "OUTER", "MAX", "TC", "ITEM", "PRODUCT"
        ]
        
        # Filter out lines with skip keywords or just numbers
        filtered = [
            line for line in lines
            if all(k.lower() not in line.lower() for k in skip_keywords)
            and not re.match(r"^[\d\s,./-]+$", line)
            and len(line) > 2  # At least 3 characters
        ]
        
        colour = "UNKNOWN"
        if filtered:
            # Take the first meaningful line as colour
            colour = filtered[0]
            # Clean up the colour
            colour = re.sub(r'[\d\.\)\(]+', '', colour).strip().upper()
            
            # Handle manual input if needed
            if "MANUAL" in colour:
                st.warning(f"⚠️ Page {page_number}: 'MANUAL' detected in colour field")
                manual = st.text_input(f"Enter Colour (Page {page_number}):", key=f"colour_manual_{page_number}")
                return manual.upper() if manual else "UNKNOWN"
            
            # Check if it's a valid colour name (not too long)
            if len(colour) > 50:
                colour = "UNKNOWN"
            
            return colour if colour else "UNKNOWN"
        
        # If no colour found, try to find any line that might be a colour
        for line in lines[:10]:  # Check first 10 lines
            if len(line) < 30 and not re.search(r'\d', line):  # Short line with no numbers
                return line.strip().upper()
        
        st.warning(f"⚠️ Page {page_number}: Colour information not found in PDF")
        manual = st.text_input(f"Enter Colour (Page {page_number}):", key=f"colour_missing_{page_number}")
        return manual.upper() if manual else "UNKNOWN"
    
    except Exception as e:
        st.error(f"Error extracting colour: {str(e)}")
        return "UNKNOWN"

# =============================
# Extract Order ID only (for multiple files)
# =============================
def extract_order_id_only(file):
    # Save current position
    pos = None
    try:
        pos = file.tell()
    except Exception:
        pass
    
    # Reset to beginning
    try:
        file.seek(0)
    except Exception:
        pass
    
    try:
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            if len(doc) < 1:
                # Restore position
                try:
                    file.seek(0 if pos is None else pos)
                except Exception:
                    pass
                return None
            
            page1_text = doc[0].get_text()
    
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        try:
            file.seek(0 if pos is None else pos)
        except Exception:
            pass
        return None
    
    # Restore position
    try:
        file.seek(0 if pos is None else pos)
    except Exception:
        pass
    
    # Extract Order ID
    match = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)", page1_text, re.IGNORECASE)
    return match.group(1).strip() if match else None


# =============================
# Process PDF (Combined old and new structure)
# =============================
def process_pdf(uploaded_file, page_for_colour=1):
    try:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        
        # Check if PDF has at least 2 pages for colour extraction
        if len(doc) < 2:
            st.warning("PDF has less than 2 pages. Some data may not extract properly.")
        
        # Combine text from all pages
        full_text = ""
        for page in doc:
            full_text += page.get_text()
        
        # Extract general data (from old file structure)
        general = extract_general_data(full_text)
        
        # Extract label data (from new file structure)
        label = extract_label_data(full_text)
        
        # Extract colour from page 2 if available, otherwise from page 1
        colour_text = doc[1].get_text() if len(doc) > 1 else doc[0].get_text()
        colour = extract_colour(colour_text, page_for_colour)
        
        # Combine all data (14 fields total)
        data = {
            # Old file fields (7)
            "Order_ID": general["Order_ID"],
            "Style": general["Style"],
            "Colour": colour,
            "Supplier_product_code": general["Supplier_product_code"],
            "Item_classification": general["Item_classification"],
            "Supplier_name": general["Supplier_name"],
            "today_date": general["today_date"],
            
            # New file fields (7)
            "TC_Number": label["TC_Number"],
            "Product_name": label["Product_name"],
            "Barcode": label["Barcode"],
            "Inner_kg": label["Inner_kg"],
            "Season": label["Season"],
            "Inner_qty": label["Inner_qty"],
            "Outer_qty": label["Outer_qty"]
        }
        
        return data
    
    except Exception as e:
        st.error(f"Error processing PDF: {str(e)}")
        return None


# =============================
# Validate extracted data
# =============================
def validate_data(df):
    issues = []
    
    # Check required fields
    if df["Order_ID"].iloc[0] == "":
        issues.append("❌ Order ID not found")
    if df["Style"].iloc[0] == "":
        issues.append("⚠️ Style not found")
    if df["Barcode"].iloc[0] == "":
        issues.append("⚠️ Barcode not found")
    
    return issues


# =============================
# UI - Upload PDF
# =============================
uploaded_files = st.file_uploader(
    "Upload PEPCO PO PDFs",
    type="pdf",
    accept_multiple_files=True
)


# =============================
# Batch Processing
# =============================
if uploaded_files:
    
    all_rows = []
    order_ids = []
    first_file = True
    
    # Progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, pdf in enumerate(uploaded_files):
        status_text.text(f"Processing file {idx+1}/{len(uploaded_files)}...")
        
        if first_file:
            # Process first file normally (extract all data)
            row = process_pdf(pdf, page_for_colour=1)
            first_file = False
        else:
            # For additional files, just get Order ID
            order_id = extract_order_id_only(pdf)
            if order_id:
                order_ids.append(order_id)
            progress_bar.progress((idx + 1) / len(uploaded_files))
            continue
        
        if row:
            all_rows.append(row)
            
            if row["Order_ID"] and row["Order_ID"] not in order_ids:
                order_ids.append(row["Order_ID"])
        
        # Update progress
        progress_bar.progress((idx + 1) / len(uploaded_files))
    
    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()
    
    if all_rows:
        # Create DataFrame from all rows
        df = pd.DataFrame(all_rows)
        
        # Combine Order IDs if there are multiple
        if len(order_ids) > 1:
            df["Order_ID"] = "+".join(order_ids)
        
        st.success(f"✅ Successfully processed {len(all_rows)} main file(s) and combined {len(order_ids)} Order ID(s)")
        
        # Define all 14 columns in the exact order
        all_columns = [
            "Order_ID",
            "Style",
            "Colour",
            "Supplier_product_code",
            "Item_classification",
            "Supplier_name",
            "today_date",
            "TC_Number",
            "Product_name",
            "Barcode",
            "Inner_kg",
            "Season",
            "Inner_qty",
            "Outer_qty"
        ]
        
        # Ensure all columns exist
        for col in all_columns:
            if col not in df.columns:
                df[col] = ""
        
        # Reorder columns
        df = df[all_columns]
        
        # Validate data
        validation_issues = validate_data(df)
        if validation_issues:
            st.warning("Data Quality Issues:")
            for issue in validation_issues:
                st.write(issue)
        
        # Display the data
        st.subheader("📋 Extracted Data (14 Fields)")
        st.dataframe(df, use_container_width=True)
        
        # =============================
        # CSV Export
        # =============================
        csv_buffer = StringIO()
        
        writer = pycsv.writer(
            csv_buffer,
            delimiter=';',
            quoting=pycsv.QUOTE_ALL
        )
        
        # Write headers
        writer.writerow(all_columns)
        
        # Write data
        for row in df.itertuples(index=False):
            writer.writerow(row)
        
        # Generate filename based on Order_ID and date
        order_id_str = df.iloc[0]["Order_ID"] if not df.empty else "data"
        # Clean filename (remove invalid characters)
        order_id_str = re.sub(r'[^\w\-_\. ]', '_', order_id_str)
        today_str = datetime.today().strftime('%Y%m%d')
        
        # Create columns for download and summary
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col1:
            st.download_button(
                "📥 Download Complete CSV (14 Fields)",
                csv_buffer.getvalue().encode('utf-8-sig'),
                file_name=f"pepco_{order_id_str}_{today_str}.csv",
                mime="text/csv"
            )
        
        with col2:
            st.info(f"📊 Records: {len(df)}")
        
        with col3:
            # Count non-empty fields
            non_empty = df[all_columns].replace('', pd.NA).count().sum()
            total_fields = len(df) * len(all_columns)
            st.info(f"✅ Data Points: {non_empty}/{total_fields}")
        
        # Show summary of extracted data
        with st.expander("📊 Data Summary"):
            summary_data = []
            for col in all_columns:
                non_empty_count = df[col].astype(str).str.len().gt(0).sum()
                empty_count = len(df) - non_empty_count
                sample_value = df[col].iloc[0] if non_empty_count > 0 else "N/A"
                summary_data.append({
                    "Field": col,
                    "Non-empty": non_empty_count,
                    "Empty": empty_count,
                    "Sample": str(sample_value)[:50] + "..." if len(str(sample_value)) > 50 else str(sample_value)
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True)
        
    else:
        st.error("❌ No data could be extracted from the uploaded files.")

# =============================
# Instructions
# =============================
with st.expander("ℹ️ How to use"):
    st.markdown("""
    ### 📋 Instructions:
    1. **Upload** one or more PEPCO PDF files
    2. The first file will be processed to extract all 14 fields
    3. Additional files will only contribute their **Order_ID** (combined with +)
    4. **Download** the CSV file with all data
    
    ### 📊 Extracted Fields:
    | Source | Fields |
    |--------|--------|
    | **Old File** | Order_ID, Style, Colour, Supplier_product_code, Item_classification, Supplier_name, today_date |
    | **New File** | TC_Number, Product_name, Barcode, Inner_kg, Season, Inner_qty, Outer_qty |
    
    ### ⚠️ Notes:
    - Colour is extracted from page 2 of the PDF
    - Empty fields will be shown as blank
    - Multiple Order IDs are combined with "+"
    """)

# =============================
# Footer
# =============================
st.markdown("---")
st.caption("PEPCO Label Automation Tool - Combined Version | Extracts 14 fields from PDF labels")
