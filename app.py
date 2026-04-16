import streamlit as st
import fitz
import pandas as pd
import re
from datetime import datetime
from io import StringIO
import csv as pycsv

st.set_page_config(page_title="PEPCO Label Automation V3", layout="wide")

st.title("PEPCO Label Automation V3 (Final)")

# =============================
# General Data (OLD BEST)
# =============================
def extract_general_data(text):
    order_match = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)", text, re.IGNORECASE)

    return {
        "Order_ID": order_match.group(1).strip() if order_match else "",
        "Style": re.search(r"Item No\s*\.{2,}\s*(\d+)", text).group(1) if re.search(r"Item No", text) else "",
        "Supplier_product_code": re.search(r"Supplier product code\s*\.{2,}\s*(.+)", text, re.IGNORECASE).group(1).strip() if re.search(r"Supplier product code", text) else "",
        "Item_classification": re.search(r"Item classification\s*\.{2,}\s*(.+)", text, re.IGNORECASE).group(1).strip() if re.search(r"Item classification", text) else "",
        "Supplier_name": re.search(r"Supplier name\s*\.{2,}\s*(.+)", text, re.IGNORECASE).group(1).strip() if re.search(r"Supplier name", text) else "",
        "today_date": datetime.today().strftime('%d-%m-%Y')
    }

# =============================
# Inner Qty
# =============================
def extract_inner_qty(text):
    match = re.search(r"(\d+)\s*Pcs\s*Inner?", text, re.IGNORECASE)
    return f"{match.group(1)} Pcs" if match else ""

# =============================
# Outer Qty (FIXED)
# =============================
def extract_outer_qty(text):
    patterns = [
        r"(\d+)\s*Inner\s*OUTER",
        r"(\d+)\s*OUTER",
        r"OUTER\s*[:.]?\s*(\d+)"
    ]
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return f"{m.group(1)} Inner"
    return ""

# =============================
# Label Data
# =============================
def extract_label_data(text):
    return {
        "TC_Number": re.search(r"TC\s*-\s*(T\d+)", text, re.IGNORECASE).group(1) if re.search(r"TC", text) else "",
        "Product_name": re.search(r"ITEM\s*\d+\s*\n\s*(.+)", text, re.IGNORECASE).group(1).strip() if re.search(r"ITEM", text) else "",
        "Barcode": re.search(r"(\d{13})", text).group(1) if re.search(r"\d{13}", text) else "",
        "Inner_kg": re.search(r"MAX\.?\s*(\d+)\s*kg", text, re.IGNORECASE).group(1) if re.search(r"kg", text) else "",
        "Season": re.search(r"\b(AW|SS)\d{2}\b", text).group(0) if re.search(r"(AW|SS)", text) else "",
        "Inner_qty": extract_inner_qty(text),
        "Outer_qty": extract_outer_qty(text)
    }

# =============================
# Colour (OLD SMART)
# =============================
def extract_colour(text):
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    skip_keywords = [
        "PURCHASE", "COLOUR", "TOTAL", "PANTONE", "SUPPLIER", "PRICE",
        "ORDERED", "SIZES", "TPG", "TPX", "USD", "NIP", "PEPCO",
        "Poland", "BARCODE", "INNER", "OUTER", "MAX", "TC", "ITEM", "PRODUCT"
    ]

    filtered = [
        line for line in lines
        if all(k.lower() not in line.lower() for k in skip_keywords)
        and not re.match(r"^[\d\s,./-]+$", line)
        and len(line) > 2
    ]

    if filtered:
        colour = filtered[0]
        colour = re.sub(r'[\d\.\)\(]+', '', colour).strip().upper()

        if len(colour) > 50:
            return "UNKNOWN"

        return colour

    return "UNKNOWN"

# =============================
# Process PDF
# =============================
def process_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    full_text = "".join([p.get_text() for p in doc])

    general = extract_general_data(full_text)
    label = extract_label_data(full_text)
    colour = extract_colour(doc[1].get_text() if len(doc) > 1 else full_text)

    return {
        **general,
        "Colour": colour,
        **label
    }

# =============================
# Upload UI
# =============================
uploaded_files = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True)

if uploaded_files:
    rows = []

    for pdf in uploaded_files:
        data = process_pdf(pdf)
        if data:
            rows.append(data)

    df = pd.DataFrame(rows)

    st.success(f"Processed {len(df)} files")
    st.dataframe(df, use_container_width=True)

    # CSV Export
    buffer = StringIO()
    writer = pycsv.writer(buffer, delimiter=';', quoting=pycsv.QUOTE_ALL)

    writer.writerow(df.columns)
    for r in df.itertuples(index=False):
        writer.writerow(r)

    st.download_button("Download CSV", buffer.getvalue(), "pepco_data.csv")

# Footer
st.markdown("---")
st.caption("PEPCO Automation Final Version | Stable & Optimized")
```
