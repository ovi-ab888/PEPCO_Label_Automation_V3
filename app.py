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
# Extract General Information
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
# Extract Colour (Improved from old file)
# =============================
def extract_colour(text, page_number=1):
    try:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        
        # Keywords to skip (from old file)
        skip_keywords = [
            "PURCHASE", "COLOUR", "TOTAL", "PANTONE", "SUPPLIER", "PRICE",
            "ORDERED", "SIZES", "TPG", "TPX", "USD", "NIP", "PEPCO",
            "Poland", "ul. Strzeszyńska 73A, 60-479 Poznań", "NIP 782-21-31-157"
        ]
        
        # Filter out lines with skip keywords or just numbers
        filtered = [
            line for line in lines
            if all(k.lower() not in line.lower() for k in skip_keywords)
            and not re.match(r"^[\d\s,./-]+$", line)
        ]
        
        colour = "UNKNOWN"
        if filtered:
            colour = filtered[0]
            # Clean up the colour
            colour = re.sub(r'[\d\.\)\(]+', '', colour).strip().upper()
            
            # Handle manual input if needed (like old file)
            if "MANUAL" in colour:
                st.warning(f"⚠️ Page {page_number}: 'MANUAL' detected in colour field")
                manual = st.text_input(f"Enter Colour (Page {page_number}):", key=f"colour_manual_{page_number}")
                return manual.upper() if manual else "UNKNOWN"
            
            return colour if colour else "UNKNOWN"
        
        st.warning(f"⚠️ Page {page_number}: Colour information not found in PDF")
        manual = st.text_input(f"Enter Colour (Page {page_number}):", key=f"colour_missing_{page_number}")
        return manual.upper() if manual else "UNKNOWN"
    
    except Exception as e:
        st.error(f"Error extracting colour: {str(e)}")
        return "UNKNOWN"


# =============================
# Extract Label Data
# =============================
def extract_label_data(text):
    # TC Number
    tc = re.search(r"TC\s*-\s*(T\d+)", text)
    
    # Product name
    product = re.search(r"ITEM\s*\d+\s*\n\s*(.+)", text)
    
    # Barcode (13 digits)
    barcode = re.search(r"\b\d{13}\b", text)
    
    # Inner quantity
    inner_qty = re.search(r"\d+\s*Pcs", text, re.IGNORECASE)
    
    # Outer quantity
    outer_qty = re.search(r"(\d+)\s*Inner", text, re.IGNORECASE)
    
    # Inner kg
    inner_kg = re.search(r"MAX\.\s*\d+\s*kg", text, re.IGNORECASE)
    
    # Season
    season = re.search(r"\b(AW|SS)\d{2}\b", text)
    
    return {
        "TC_Number": tc.group(1) if tc else "",
        "Product_name": product.group(1).strip() if product else "",
        "Barcode": barcode.group(0) if barcode else "",
        "Inner_kg": inner_kg.group(0) if inner_kg else "",
        "Season": season.group(0) if season else "",
        "Inner_qty": inner_qty.group(0) if inner_qty else "",
        "Outer_qty": outer_qty.group(0) if outer_qty else ""
    }

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
# Process PDF (Updated with old file structure)
# =============================
def process_pdf(uploaded_file, page_for_colour=1):
    try:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        
        # Check if PDF has at least 2 pages for colour extraction
        if len(doc) < 2:
            st.warning("PDF has less than 2 pages. Colour extraction may not work properly.")
        
        # Combine text from all pages
        full_text = ""
        for page in doc:
            full_text += page.get_text()
        
        # Extract general data (from all pages)
        general = extract_general_data(full_text)
        
        # Extract label data
        label = extract_label_data(full_text)
        
        # Extract colour from page 2 if available, otherwise from page 1
        colour_text = doc[1].get_text() if len(doc) > 1 else doc[0].get_text()
        colour = extract_colour(colour_text, page_for_colour)
        
        # Combine all data
        data = {
            "Order_ID": general["Order_ID"],
            "Style": general["Style"],
            "Colour": colour,
            "Supplier_product_code": general["Supplier_product_code"],
            "Item_classification": general["Item_classification"],
            "Supplier_name": general["Supplier_name"],
            "today_date": general["today_date"],
            # Additional fields (optional)
            "TC_Number": label["TC_Number"],
            "Product_name": label["Product_name"],
            "Barcode": label["Barcode"],
            "Season": label["Season"],
        }
        
        return data
    
    except Exception as e:
        st.error(f"Error processing PDF: {str(e)}")
        return None


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
    
    for idx, pdf in enumerate(uploaded_files):
        with st.spinner(f"Processing file {idx+1}/{len(uploaded_files)}..."):
            
            if first_file:
                # Process first file normally
                row = process_pdf(pdf, page_for_colour=1)
                first_file = False
            else:
                # For additional files, just get Order ID
                order_id = extract_order_id_only(pdf)
                if order_id:
                    order_ids.append(order_id)
                continue
            
            if row:
                all_rows.append(row)
                
                if row["Order_ID"] and row["Order_ID"] not in order_ids:
                    order_ids.append(row["Order_ID"])
    
    if all_rows:
        # Create DataFrame from all rows
        df = pd.DataFrame(all_rows)
        
        # Combine Order IDs if there are multiple
        if len(order_ids) > 1:
            df["Order_ID"] = "+".join(order_ids)
        
        st.success(f"✅ Successfully processed {len(all_rows)} file(s)")
        
        # Display only the required 7 columns
        required_columns = [
            "Order_ID",
            "Style",
            "Colour",
            "Supplier_product_code",
            "Item_classification",
            "Supplier_name",
            "today_date"
        ]
        
        # Ensure all required columns exist
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""
        
        # Show only required columns
        display_df = df[required_columns]
        
        st.subheader("Extracted Data")
        st.dataframe(display_df, use_container_width=True)
        
        # =============================
        # CSV Export (with only required columns)
        # =============================
        csv_buffer = StringIO()
        
        writer = pycsv.writer(
            csv_buffer,
            delimiter=';',
            quoting=pycsv.QUOTE_ALL
        )
        
        # Write headers
        writer.writerow(required_columns)
        
        # Write data
        for row in display_df.itertuples(index=False):
            writer.writerow(row)
        
        # Generate filename based on Order_ID
        order_id_str = display_df.iloc[0]["Order_ID"] if not display_df.empty else "data"
        # Clean filename (remove invalid characters)
        order_id_str = re.sub(r'[^\w\-_\. ]', '_', order_id_str)
        
        st.download_button(
            "📥 Download CSV",
            csv_buffer.getvalue().encode('utf-8-sig'),
            file_name=f"pepco_{order_id_str}.csv",
            mime="text/csv"
        )
        
        # Show summary
        st.info(f"📊 Total Records: {len(df)}")
        
    else:
        st.error("❌ No data could be extracted from the uploaded files.")

# =============================
# Footer
# =============================
st.markdown("---")
st.caption("PEPCO Label Automation Tool - Extracts Order_ID, Style, Colour, Supplier_product_code, Item_classification, Supplier_name, and Date")
