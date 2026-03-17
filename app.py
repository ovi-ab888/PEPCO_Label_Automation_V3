import streamlit as st
import fitz
import pandas as pd
import re
from datetime import datetime
from io import StringIO
import csv as pycsv

st.set_page_config(page_title="PEPCO Label Automation V2", layout="wide")

st.title("PEPCO Label Automation V2")

# ------------------------------
# Extract General Info
# ------------------------------
def extract_general_data(text):

    order_id = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_]+)", text)
    style = re.search(r"\b\d{6}\b", text)
    supplier_code = re.search(r"Supplier product code\s*\.{2,}\s*(.+)", text)
    item_class = re.search(r"Item classification\s*\.{2,}\s*(.+)", text)
    supplier = re.search(r"Supplier name\s*\.{2,}\s*(.+)", text)

    return {
        "Order_ID": order_id.group(1) if order_id else "",
        "Style": style.group(0) if style else "",
        "Supplier_product_code": supplier_code.group(1).strip() if supplier_code else "",
        "Item_classification": item_class.group(1).strip() if item_class else "",
        "Supplier_name": supplier.group(1).strip() if supplier else "",
        "today_date": datetime.today().strftime('%d-%m-%Y')
    }


# ------------------------------
# Extract Label Data
# ------------------------------
def extract_label_data(text):

    tc = re.search(r"TC\s*-\s*(T\d+)", text)
    item = re.search(r"ITEM\s*(\d+)", text)
    product = re.search(r"ITEM\s*\d+\s*\n\s*(.+)", text)
    barcode = re.search(r"\b\d{13}\b", text)

    inner_qty = re.search(r"(\d+)\s*Pcs\s*Inner", text, re.IGNORECASE)
    outer_qty = re.search(r"(\d+)\s*Inner\s*OUTER", text, re.IGNORECASE)

    inner_kg = re.search(r"MAX\.\s*(\d+\s*kg)", text)
    season = re.search(r"\b(AW|SS)\d{2}\b", text)

    return {
        "TC_Number": tc.group(1) if tc else "",
        "ITEM_Number": item.group(1) if item else "",
        "Product_name": product.group(1).strip() if product else "",
        "Barcode": barcode.group(0) if barcode else "",
        "Inner_kg": inner_kg.group(1) if inner_kg else "",
        "Season": season.group(0) if season else "",
        "Inner_qty": inner_qty.group(1) if inner_qty else "",
        "Outer_qty": outer_qty.group(1) if outer_qty else ""
    }


# ------------------------------
# Extract Colour
# ------------------------------
def extract_colour(text):

    match = re.search(r"Multicolor\s+([\d\-]+)", text)

    if match:
        return "MULTICOLOR"

    return "UNKNOWN"


# ------------------------------
# Process PDF
# ------------------------------
def process_pdf(uploaded_file):

    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")

    full_text = ""

    for page in doc:
        full_text += page.get_text()

    general = extract_general_data(full_text)
    label = extract_label_data(full_text)

    colour = extract_colour(full_text)

    data = {
        "Order_ID": general["Order_ID"],
        "Style": general["Style"],
        "Colour": colour,
        "Supplier_product_code": general["Supplier_product_code"],
        "Item_classification": general["Item_classification"],
        "Supplier_name": general["Supplier_name"],
        "today_date": general["today_date"],
        "TC_Number": label["TC_Number"],
        "Product_name": label["Product_name"],
        "Barcode": label["Barcode"],
        "Inner_kg": label["Inner_kg"],
        "Season": label["Season"],
        "Inner_qty": label["Inner_qty"],
        "Outer_qty": label["Outer_qty"]
    }

    return pd.DataFrame([data])


# ------------------------------
# UI
# ------------------------------

uploaded_pdf = st.file_uploader("Upload PEPCO PO PDF", type="pdf")

if uploaded_pdf:

    df = process_pdf(uploaded_pdf)

    st.success("Extraction Complete")

    st.dataframe(df)

    # CSV Export
    csv_buffer = StringIO()
    writer = pycsv.writer(csv_buffer, delimiter=';', quoting=pycsv.QUOTE_ALL)

    cols = list(df.columns)

    writer.writerow(cols)

    for row in df.itertuples(index=False):
        writer.writerow(row)

    st.download_button(
        "Download CSV",
        csv_buffer.getvalue().encode('utf-8-sig'),
        file_name="pepco_label_data.csv",
        mime="text/csv"
    )
