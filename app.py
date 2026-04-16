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
    order_id = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)", text, re.IGNORECASE)
    style = re.search(r"\b\d{6}\b", text)
    supplier_code = re.search(r"Supplier product code\s*\.{2,}\s*(.+)", text, re.IGNORECASE)
    item_class = re.search(r"Item classification\s*\.{2,}\s*(.+)", text, re.IGNORECASE)
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
# Outer Qty FIXED (NEW)
# =============================
def extract_outer_qty(text):
    patterns = [
        r"(\d+)\s*Inner\s*OUTER",
        r"(\d+)\s*OUTER",
        r"OUTER\s*[:.]?\s*(\d+)",
        r"(\d+)\s*X\s*INNER\s*OUTER",
        r"OUTER\s*QTY\s*[:.]?\s*(\d+)"
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return f"{match.group(1)} Inner"

    return ""

# =============================
# Extract Label Data
# =============================
def extract_label_data(text):
    tc = re.search(r"TC\s*-\s*(T\d+)", text, re.IGNORECASE)
    if not tc:
        tc = re.search(r"TC\s*[:.]?\s*(T\d+)", text, re.IGNORECASE)

    product = re.search(r"ITEM\s*\d+\s*\n\s*(.+)", text, re.IGNORECASE)
    if not product:
        product = re.search(r"Product\s*name\s*[:.]?\s*(.+)", text, re.IGNORECASE)

    barcode = re.search(r"\b\d{13}\b", text)
    inner_qty = re.search(r"(\d+)\s*Pcs", text, re.IGNORECASE)
    inner_kg = re.search(r"MAX\.?\s*(\d+)\s*kg", text, re.IGNORECASE)
    if not inner_kg:
        inner_kg = re.search(r"(\d+)\s*kg", text, re.IGNORECASE)

    season = re.search(r"\b(AW|SS|FW|SW)\d{2}\b", text, re.IGNORECASE)

    return {
        "TC_Number": tc.group(1) if tc else "",
        "Product_name": product.group(1).strip() if product else "",
        "Barcode": barcode.group(0) if barcode else "",
        "Inner_kg": f"MAX. {inner_kg.group(1)} kg" if inner_kg else "",
        "Season": season.group(0).upper() if season else "",
        "Inner_qty": f"{inner_qty.group(1)} Pcs" if inner_qty else "",
        "Outer_qty": extract_outer_qty(text)   # ✅ FIXED HERE
    }

# =============================
# Extract Colour
# =============================
def extract_colour(text, page_number=1):
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    skip_keywords = [
        "PURCHASE", "COLOUR", "TOTAL", "PANTONE", "SUPPLIER", "PRICE",
        "ORDERED", "SIZES", "TPG", "TPX", "USD", "PEPCO",
        "BARCODE", "INNER", "OUTER", "MAX", "TC", "ITEM", "PRODUCT"
    ]

    filtered = [
        line for line in lines
        if all(k.lower() not in line.lower() for k in skip_keywords)
        and not re.match(r"^[\d\s,./-]+$", line)
        and len(line) > 2
    ]

    if filtered:
        colour = re.sub(r'[\d\.\)\(]+', '', filtered[0]).strip().upper()
        return colour if colour else "UNKNOWN"

    return "UNKNOWN"

# =============================
# Process PDF
# =============================
def process_pdf(uploaded_file, page_for_colour=1):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")

    full_text = ""
    for page in doc:
        full_text += page.get_text()

    general = extract_general_data(full_text)
    label = extract_label_data(full_text)

    colour_text = doc[1].get_text() if len(doc) > 1 else doc[0].get_text()
    colour = extract_colour(colour_text, page_for_colour)

    return {
        **general,
        "Colour": colour,
        **label
    }

# =============================
# Validate
# =============================
def validate_data(df):
    issues = []
    if df["Order_ID"].iloc[0] == "":
        issues.append("❌ Order ID not found")
    if df["Barcode"].iloc[0] == "":
        issues.append("⚠️ Barcode not found")
    return issues

# =============================
# UI
# =============================
uploaded_files = st.file_uploader("Upload PEPCO PO PDFs", type="pdf", accept_multiple_files=True)

if uploaded_files:

    all_rows = []
    order_ids = []
    first_file = True

    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, pdf in enumerate(uploaded_files):
        status_text.text(f"Processing file {idx+1}/{len(uploaded_files)}...")

        if first_file:
            row = process_pdf(pdf)
            first_file = False
        else:
            order_id = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)", pdf.read().decode("latin1"), re.IGNORECASE)
            if order_id:
                order_ids.append(order_id.group(1))
            progress_bar.progress((idx + 1) / len(uploaded_files))
            continue

        if row:
            all_rows.append(row)
            if row["Order_ID"]:
                order_ids.append(row["Order_ID"])

        progress_bar.progress((idx + 1) / len(uploaded_files))

    progress_bar.empty()
    status_text.empty()

    if all_rows:

        df = pd.DataFrame(all_rows)

        if len(order_ids) > 1:
            df["Order_ID"] = "+".join(order_ids)

        all_columns = [
            "Order_ID","Style","Colour","Supplier_product_code",
            "Item_classification","Supplier_name","today_date",
            "TC_Number","Product_name","Barcode","Inner_kg",
            "Season","Inner_qty","Outer_qty"
        ]

        for col in all_columns:
            if col not in df.columns:
                df[col] = ""

        df = df[all_columns]

        st.subheader("📋 Extracted Data")
        st.dataframe(df, use_container_width=True)

        # CSV export
        csv_buffer = StringIO()
        writer = pycsv.writer(csv_buffer, delimiter=';', quoting=pycsv.QUOTE_ALL)

        writer.writerow(all_columns)
        for row in df.itertuples(index=False):
            writer.writerow(row)

        st.download_button(
            "📥 Download CSV",
            csv_buffer.getvalue().encode('utf-8-sig'),
            file_name=f"pepco_{datetime.today().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )

        with st.expander("📊 Data Summary"):
            st.dataframe(df.describe(include='all'))

    else:
        st.error("❌ No data extracted")

st.markdown("---")
st.caption("PEPCO Label Automation V2 | Updated Outer Qty Logic")
