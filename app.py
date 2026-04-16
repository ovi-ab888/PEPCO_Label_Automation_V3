import streamlit as st
import fitz
import pandas as pd
import re
from datetime import datetime
from io import StringIO
import csv as pycsv

st.set_page_config(page_title="PEPCO Label Automation V3", layout="wide")
st.title("PEPCO Label Automation V3")

# ================================================================
# GENERAL DATA
# ================================================================
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

# ================================================================
# OUTER QTY (FIXED)
# ================================================================
def extract_outer_qty(text):
    patterns = [
        r"(\d+)\s*Inner\s*OUTER",
        r"(\d+)\s*OUTER",
        r"OUTER\s*[:.]?\s*(\d+)",
        r"(\d+)\s*X\s*INNER\s*OUTER",
        r"OUTER\s*QTY\s*[:.]?\s*(\d+)"
    ]

    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return f"{m.group(1)} Inner"

    return ""

# ================================================================
# LABEL DATA
# ================================================================
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
        "Outer_qty": extract_outer_qty(text)
    }

# ================================================================
# COLOUR EXTRACTION (FINAL FIXED)
# ================================================================
def extract_colour_from_page2(text):
    try:
        m = re.search(
            r"Colour[^\n]*?\n\s*([A-Za-z ]+)\s+([0-9]{2}-[0-9]{4}[A-Za-z]*)",
            text,
            re.IGNORECASE
        )
        if m:
            return m.group(1).strip().upper()
    except:
        pass
    return None


def extract_colour_from_pdf_pages(pages_text, page_number=1):

    # 1️⃣ Standard format
    for txt in pages_text:
        m = re.search(
            r"Colour.*?\n.*?\n\s*([A-Za-z ]+)\s+[0-9]{2}-[0-9]{4}",
            txt,
            re.IGNORECASE | re.DOTALL
        )
        if m:
            return m.group(1).strip().upper()

    # 2️⃣ Purchase price block
    for txt in pages_text:
        m = re.search(
            r"Purchase price.*?\n\s*([A-Za-z ]+)\s+[0-9]{2}-[0-9]{4}",
            txt,
            re.IGNORECASE | re.DOTALL
        )
        if m:
            return m.group(1).strip().upper()

    # 3️⃣ Broken line format
    for txt in pages_text:
        if "colour" in txt.lower():
            for line in txt.splitlines():
                m = re.search(r"([A-Za-z ]+)\s+[0-9]{2}-[0-9]{4}", line)
                if m:
                    return m.group(1).strip().upper()

    # 4️⃣ Page 2 fallback
    if len(pages_text) > 1:
        legacy = extract_colour_from_page2(pages_text[1])
        if legacy:
            return legacy

    # 5️⃣ Manual input
    st.warning("⚠️ Colour not found in PDF")

    key = f"manual_colour_{page_number}"
    if key not in st.session_state:
        st.session_state[key] = ""

    manual = st.text_input("Enter Colour:", key=key)

    return manual.strip().upper() if manual else "UNKNOWN"

# ================================================================
# PROCESS PDF
# ================================================================
def process_pdf(uploaded_file):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")

    pages_text = [page.get_text() for page in doc]
    full_text = "\n".join(pages_text)

    general = extract_general_data(full_text)
    label = extract_label_data(full_text)
    colour = extract_colour_from_pdf_pages(pages_text)

    return {
        **general,
        "Colour": colour,
        **label
    }

# ================================================================
# UI
# ================================================================
uploaded_files = st.file_uploader(
    "Upload PEPCO PO PDFs",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:

    all_rows = []
    order_ids = []
    first = True

    progress = st.progress(0)
    status = st.empty()

    for i, pdf in enumerate(uploaded_files):

        status.text(f"Processing file {i+1}/{len(uploaded_files)}...")

        if first:
            row = process_pdf(pdf)
            first = False
        else:
            order_id = re.search(
                r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)",
                pdf.read().decode("latin1"),
                re.IGNORECASE
            )
            if order_id:
                order_ids.append(order_id.group(1))

            progress.progress((i + 1) / len(uploaded_files))
            continue

        if row:
            all_rows.append(row)
            if row["Order_ID"]:
                order_ids.append(row["Order_ID"])

        progress.progress((i + 1) / len(uploaded_files))

    progress.empty()
    status.empty()

    if all_rows:

        df = pd.DataFrame(all_rows)

        if len(order_ids) > 1:
            df["Order_ID"] = "+".join(order_ids)

        columns = [
            "Order_ID","Style","Colour","Supplier_product_code",
            "Item_classification","Supplier_name","today_date",
            "TC_Number","Product_name","Barcode","Inner_kg",
            "Season","Inner_qty","Outer_qty"
        ]

        for c in columns:
            if c not in df.columns:
                df[c] = ""

        df = df[columns]

        st.subheader("📋 Extracted Data")
        st.dataframe(df, use_container_width=True)

        # CSV Export
        buffer = StringIO()
        writer = pycsv.writer(buffer, delimiter=';', quoting=pycsv.QUOTE_ALL)

        writer.writerow(columns)
        for r in df.itertuples(index=False):
            writer.writerow(r)

        st.download_button(
            "📥 Download CSV",
            buffer.getvalue().encode("utf-8-sig"),
            file_name=f"pepco_{datetime.today().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )

    else:
        st.error("❌ No data extracted")

st.markdown("---")
st.caption("PEPCO Label Automation V3 | Final Stable Version")
