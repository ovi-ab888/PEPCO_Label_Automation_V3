import streamlit as st
import fitz
import pandas as pd
import re
from datetime import datetime
from io import StringIO
import csv as pycsv

st.set_page_config(page_title="PEPCO Label Automation V2", layout="wide")

st.title("PEPCO Label Automation V2")


# =============================
# Extract General Information
# =============================
def extract_general_data(text):

    order_id = re.search(r"ORD\d+_\d+", text)
    style = re.search(r"\b\d{6}\b", text)

    supplier_code = re.search(r"Supplier product code\s*\.{2,}\s*(.+)", text)
    item_class = re.search(r"Item classification\s*\.{2,}\s*(.+)", text)
    supplier = re.search(r"Supplier name\s*\.{2,}\s*(.+)", text)

    return {
        "Order_ID": order_id.group(0) if order_id else "",
        "Style": style.group(0) if style else "",
        "Supplier_product_code": supplier_code.group(1).strip() if supplier_code else "",
        "Item_classification": item_class.group(1).strip() if item_class else "",
        "Supplier_name": supplier.group(1).strip() if supplier else "",
        "today_date": datetime.today().strftime('%d-%m-%Y')
    }


# =============================
# Extract Label Data
# =============================
def extract_label_data(text):

    tc = re.search(r"TC\s*-\s*(T\d+)", text)

    product = re.search(r"ITEM\s*\d+\s*\n\s*(.+)", text)

    barcode = re.search(r"\b\d{13}\b", text)

    inner_qty = re.search(r"\d+\s*Pcs", text, re.IGNORECASE)

    outer_qty = re.search(r"\d+\s*Inner", text, re.IGNORECASE)

    inner_kg = re.search(r"MAX\.\s*\d+\s*kg", text, re.IGNORECASE)

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
# Extract Colour
# =============================
def extract_colour(text):

    match = re.search(r"Multicolor", text, re.IGNORECASE)

    if match:
        return "MULTICOLOR"

    return "UNKNOWN"


# =============================
# Process PDF
# =============================
def process_pdf(uploaded_file):

    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")

    text = ""

    for page in doc:
        text += page.get_text()

    general = extract_general_data(text)

    label = extract_label_data(text)

    colour = extract_colour(text)

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

    return data


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

    order_ids = []

    first_row = None

    for pdf in uploaded_files:

        row = process_pdf(pdf)

        if row:

            if row["Order_ID"]:
                order_ids.append(row["Order_ID"])

            if first_row is None:
                first_row = row


    if first_row:

        combined_order_id = "+".join(order_ids)

        first_row["Order_ID"] = combined_order_id

        df = pd.DataFrame([first_row])

        st.success("Batch Processing Complete")

        st.dataframe(df)


        # =============================
        # CSV Export
        # =============================
        csv_buffer = StringIO()

        writer = pycsv.writer(
            csv_buffer,
            delimiter=';',
            quoting=pycsv.QUOTE_ALL
        )

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
