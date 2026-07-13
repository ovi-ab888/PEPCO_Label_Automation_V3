# ================================================================
# PART 1 — PAGE CONFIG + IMPORTS + THEME + PASSWORD + CONSTANTS
# ================================================================

# ---------- PAGE CONFIG (must be at top) ----------
import streamlit as st
st.set_page_config(
    page_title="PEPCO SL",
    page_icon="🧾",
    layout="wide"
)

# ---------- Imports ----------
import fitz  # PyMuPDF
import pandas as pd
import re
from io import StringIO
import csv as pycsv
from datetime import datetime
import os


# ================================================================
#  MAPPING DICTIONARIES
# ================================================================

PICTOGRAM_MAPPING = {
    "PIC00033": "a",
    "PIC00034": "f",
    "PIC00032": "c",
    "PIC00181": "g",
    "PIC00030": "m",
    "PIC00183": "v",
    "PIC00182": "i",
    "PIC00031": "e",
    "PIC00028": "t",
    "PIC00185": "o",
    "PIC00027": "p",
    "PIC00029": "r"
}

PROMOTIONAL_MAPPING = {
    "PROMO": "p",
    "KVI": "K",
    "HS": "H"
}


# ================================================================
#  LOGO & THEME
# ================================================================
LOGO_PNG = "logo.png"
LOGO_SVG = "logo.svg"

THEME_CSS = """
<style>
:root{
  --card-bg: rgba(255,255,255,.04);
  --card-br: rgba(255,255,255,.12);
  --input-bg: rgba(255,255,255,.08);
  --input-br: rgba(255,255,255,.25);
  --txt:      #E9ECF6;
  --muted:    #C2C8DF;
}

.block-container{max-width:1120px; padding-top:1rem; padding-bottom:3rem;}
h1,h2,h3{font-weight:700;}
h1{letter-spacing:.2px;} h2,h3{letter-spacing:.1px;}

section[data-testid="stFileUploader"],
div[data-testid="stDataFrameContainer"],
div[data-testid="stVerticalBlock"]:has(> div[data-testid="stDataEditor"]){
  background:var(--card-bg)!important;
  border:1px solid var(--card-br)!important;
  border-radius:14px!important;
  padding:12px 14px;
  box-shadow:0 1px 8px rgba(0,0,0,.12);
}

label, .stMultiSelect label, .stSelectbox label, .stNumberInput label, .stTextInput label{
  color:var(--txt)!important; font-weight:500;
}

input, textarea{
  color:var(--txt)!important;
  background:var(--input-bg)!important;
  border-color:var(--input-br)!important;
}
input::placeholder, textarea::placeholder{
  color:var(--muted)!important; opacity:.95;
}

div[data-baseweb="select"] > div{
  background:var(--input-bg)!important;
  border-color:var(--input-br)!important;
  border-radius:12px!important;
}
div[data-baseweb="select"] input{ color:var(--txt)!important; }
div[data-baseweb="select"] svg{ opacity:.9; }

div[data-testid="stNumberInput"] input{
  color:var(--txt)!important;
  background:var(--input-bg)!important;
  border-color:var(--input-br)!important;
}

.stButton > button{
  border-radius:12px; padding:.55rem 1rem;
}

[data-testid="stTable"] td,[data-testid="stTable"] th{
  padding:.45rem .6rem;
}
</style>
"""


# ================================================================
#  PASSWORD CHECK SYSTEM
# ================================================================
def check_password():
    """Simple password gate using secrets or environment."""
    expected = None

    try:
        expected = st.secrets.get("app_password", None)
    except Exception:
        expected = None

    if expected is None:
        expected = os.environ.get("PEPCO_APP_PASSWORD")

    if expected is None:
        st.error("App password not configured. Please set 'app_password' in secrets or PEPCO_APP_PASSWORD env var.")
        return False

    def _password_entered():
        if st.session_state.get("password") == expected:
            st.session_state["password_correct"] = True
            try:
                del st.session_state["password"]
            except Exception:
                pass
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", None) is True:
        return True

    st.text_input("Enter Your Access Code", type="password", key="password", on_change=_password_entered)

    if st.session_state.get("password_correct") is False:
        st.error("Your password Incorrect, Please contact Mr. Ovi")

    return False


# ================================================================
#  EXTRACTION FUNCTIONS
# ================================================================

def extract_all_tc_numbers_from_page4_plus(pages_text):
    """
    Extract ALL TC numbers ONLY from PAGE 4 and onwards.
    Pages 1, 2, 3 are SKIPPED completely.
    Returns max 7 unique TC numbers.
    """
    tc_list = []
    
    if len(pages_text) >= 4:
        for i in range(3, len(pages_text)):
            page_text = pages_text[i]
            
            patterns = [
                r"TC\s*-\s*(T\d+)",
                r"TC\s*[:.]?\s*(T\d+)"
            ]
            
            for pattern in patterns:
                matches = re.findall(pattern, page_text, re.IGNORECASE)
                for m in matches:
                    if m not in tc_list:
                        tc_list.append(m)
    
    return tc_list[:7]


def extract_all_barcodes_from_page4_plus(pages_text):
    """
    Extract ALL barcodes (13 digits) ONLY from PAGE 4 and onwards.
    Pages 1, 2, 3 are SKIPPED completely.
    Returns max 7 unique barcodes.
    """
    barcode_list = []
    
    if len(pages_text) >= 4:
        for i in range(3, len(pages_text)):
            page_text = pages_text[i]
            barcodes_on_page = re.findall(r"\b\d{13}\b", page_text)
            barcode_list.extend(barcodes_on_page)
    
    unique_barcodes = []
    for b in barcode_list:
        if b not in unique_barcodes:
            unique_barcodes.append(b)
    
    return unique_barcodes[:7]


def extract_product_name_from_page4_plus(pages_text):
    """Extract product name ONLY from PAGE 4 and onwards."""
    if len(pages_text) >= 4:
        for i in range(3, len(pages_text)):
            text = pages_text[i]
            m = re.search(r"ITEM\s*\d+\s*\n\s*(.+)", text, re.IGNORECASE)
            if not m:
                m = re.search(r"Product\s*name\s*[:.]?\s*(.+)", text, re.IGNORECASE)
            if m:
                return m.group(1).strip()
    return ""


def extract_inner_kg_from_page4_plus(pages_text):
    """Extract inner kg ONLY from PAGE 4 and onwards."""
    if len(pages_text) >= 4:
        for i in range(3, len(pages_text)):
            text = pages_text[i]
            m = re.search(r"MAX\.?\s*(\d+)\s*kg", text, re.IGNORECASE)
            if not m:
                m = re.search(r"(\d+)\s*kg", text, re.IGNORECASE)
            if m:
                return f"MAX. {m.group(1)} kg"
    return ""


def extract_season_from_page4_plus(pages_text):
    """Extract season code ONLY from PAGE 4 and onwards."""
    if len(pages_text) >= 4:
        for i in range(3, len(pages_text)):
            text = pages_text[i]
            m = re.search(r"\b(AW|SS|FW|SW)\d{2}\b", text, re.IGNORECASE)
            if m:
                return m.group(0).upper()
    return ""


def extract_inner_qty_from_page4_plus(pages_text):
    """Extract inner quantity ONLY from PAGE 4 and onwards."""
    if len(pages_text) >= 4:
        for i in range(3, len(pages_text)):
            text = pages_text[i]
            m = re.search(r"(\d+)\s*Pcs", text, re.IGNORECASE)
            if m:
                return f"{m.group(1)} Pcs"
    return ""


def extract_outer_qty_from_page4_plus(pages_text):
    """Extract outer quantity ONLY from PAGE 4 and onwards."""
    if len(pages_text) >= 4:
        patterns = [
            r"(\d+)\s*Inner\s*OUTER",
            r"(\d+)\s*OUTER",
            r"OUTER\s*[:.]?\s*(\d+)",
            r"(\d+)\s*X\s*INNER\s*OUTER",
            r"OUTER\s*QTY\s*[:.]?\s*(\d+)"
        ]
        for i in range(3, len(pages_text)):
            text = pages_text[i]
            for p in patterns:
                m = re.search(p, text, re.IGNORECASE)
                if m:
                    return f"{m.group(1)} Inner"
    return ""


def clean_item_name_english(name: str) -> str:
    """Clean Item_name_EN by removing prefixes."""
    if not isinstance(name, str):
        return ""
    
    text = name.strip()
    lower = text.lower()
    
    prefixes = ["xxxxx", "xxxxx", "xxxxx", "xxxxx", "xxxxx", "xxxxx", "xxxxx", "xxxxx"]
    
    for p in prefixes:
        if lower.startswith(p):
            cut_len = len(p)
            text = text[cut_len:].strip(" -_,./").strip()
            break
    
    return text.upper()


def extract_colour_from_pdf_pages(pages_text):
    """Extract colour from PDF pages."""
    for txt in pages_text:
        m = re.search(r"Colour.*?\n.*?\n\s*([A-Za-z ]+)\s+[0-9]{2}-[0-9]{4}", txt, re.IGNORECASE | re.DOTALL)
        if m:
            return m.group(1).strip().upper()
    
    for txt in pages_text:
        m2 = re.search(r"Purchase price.*?\n\s*([A-Za-z ]+)\s+[0-9]{2}-[0-9]{4}", txt, re.IGNORECASE | re.DOTALL)
        if m2:
            return m2.group(1).strip().upper()
    
    for txt in pages_text:
        if "colour" in txt.lower():
            for line in txt.splitlines():
                if re.search(r"[A-Za-z ]+\s+[0-9]{2}-[0-9]{4}", line):
                    name = line.split()[0:-1]
                    if name:
                        return " ".join(name).upper()
    
    st.warning("⚠️ Colour not found in PDF. Enter colour manually:")
    manual = st.text_input("Colour (e.g. WHITE):", key="manual_colour_fix")
    return manual.strip().upper() if manual else "UNKNOWN"


def extract_order_id_only(file):
    """Extract only Order ID from a PDF file."""
    pos = None
    try:
        pos = file.tell()
    except Exception:
        pass
    
    try:
        file.seek(0)
    except Exception:
        pass
    
    try:
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            page1_text = doc[0].get_text() if len(doc) > 0 else ""
    except Exception:
        try:
            file.seek(0 if pos is None else pos)
        except Exception:
            pass
        return None
    
    try:
        file.seek(0 if pos is None else pos)
    except Exception:
        pass
    
    m = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)", page1_text, re.IGNORECASE)
    return m.group(1).strip() if m else None


# ================================================================
#  NEW EXTRACTION FUNCTIONS FOR PICTOGRAM AND PROMOTIONAL
# ================================================================

def extract_pictogram_from_page1(page1_text):
    """
    Extract pictogram code from Page 1 and apply mapping.
    Example: "Pictogram no PIC00030" -> "m"
    """
    # Multiple patterns try করব
    patterns = [
        r"Pictogram\s*no\s*(PIC\d+)",  # "Pictogram no PIC00030"
        r"Pictogram\s*no\s*\.{0,2}\s*(\w+)",  # "Pictogram no PIC00030"
        r"Pictogram\s*[:.]?\s*(PIC\d+)",  # "Pictogram: PIC00030"
    ]
    
    for pattern in patterns:
        m = re.search(pattern, page1_text, re.IGNORECASE)
        if m:
            pictogram_code = m.group(1).strip().upper()
            # Apply mapping
            mapped_value = PICTOGRAM_MAPPING.get(pictogram_code, "")
            if mapped_value:
                return mapped_value
            # যদি mapping এ না থাকে, তবুও code return করব
            return pictogram_code
    
    return ""


def extract_promotional_from_page1(page1_text):
    """
    Extract promotional code from Page 1 and apply mapping.
    Example: "Promotional product KVI" -> "K"
    """
    # Multiple patterns try করব
    patterns = [
        r"Promotional\s*product\s*(\w+)",  # "Promotional product KVI"
        r"Promotional\s*product\s*\.{0,2}\s*(\w+)",  # "Promotional product KVI"
        r"Promotional\s*[:.]?\s*(\w+)",  # "Promotional: KVI"
    ]
    
    for pattern in patterns:
        m = re.search(pattern, page1_text, re.IGNORECASE)
        if m:
            promo_code = m.group(1).strip().upper()
            # Apply mapping
            mapped_value = PROMOTIONAL_MAPPING.get(promo_code, "")
            if mapped_value:
                return mapped_value
            # যদি mapping এ না থাকে, তবুও code return করব
            return promo_code
    
    return ""

# ================================================================
#  MAIN PDF EXTRACTION ENGINE
# ================================================================

def extract_data_from_pdf(file):
    """Main PDF extractor with all fields."""
    try:
        raw = file.read()
        if not raw:
            st.error("Empty PDF uploaded.")
            return None
        
        doc = fitz.open(stream=raw, filetype="pdf")
        
        if len(doc) < 1:
            st.error("PDF must have at least 1 page.")
            return None
        
        pages_text = [doc[i].get_text() for i in range(len(doc))]
        full_text = "\n".join(pages_text)
        page1 = pages_text[0]
        
        # Extract from page 4 onwards
        all_tc_numbers = extract_all_tc_numbers_from_page4_plus(pages_text)
        all_barcodes = extract_all_barcodes_from_page4_plus(pages_text)
        
        # Extract other fields (ONLY from page 4+)
        product_name = extract_product_name_from_page4_plus(pages_text)
        inner_kg = extract_inner_kg_from_page4_plus(pages_text)
        season_st = extract_season_from_page4_plus(pages_text)
        inner_qty = extract_inner_qty_from_page4_plus(pages_text)
        outer_qty = extract_outer_qty_from_page4_plus(pages_text)
        
        # Extract Pictogram and Promotional from Page 1
        pictogram = extract_pictogram_from_page1(page1)
        promotional = extract_promotional_from_page1(page1)
        
        # Item Name EN
        item_name_en = None
        m_item = re.search(r"Item\s*name\s*English\s*[:\.]{1,}\s*(.+)", full_text, re.IGNORECASE)
        if not m_item:
            m_item = re.search(r"Item\s*name\s*[:\.]{1,}\s*(.+?)\n", full_text, re.IGNORECASE)
        if m_item:
            item_name_en = m_item.group(1).strip()
        
        # Identifiers
        style_code = re.search(r"\b\d{6}\b", page1)
        order_id = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*(.+)", page1)
        item_class = re.search(r"Item classification\s*\.{2,}\s*(.+)", page1)
        supplier_code = re.search(r"Supplier product code\s*\.{2,}\s*(.+)", page1)
        supplier_name = re.search(r"Supplier name\s*\.{2,}\s*(.+)", page1)
        season = re.search(r"Season\s*\.{2,}\s*(\w+)?\s*(\d{2})", page1)
        
        item_class_value = item_class.group(1).strip() if item_class else "UNKNOWN"
        colour = extract_colour_from_pdf_pages(pages_text)
        
        # SKU extraction for filename
        skus = []
        for txt in pages_text:
            skus.extend(re.findall(r"\b\d{8}\b", txt))
        
        def _dedupe(seq):
            seen = set()
            out = []
            for x in seq:
                if x not in seen:
                    seen.add(x)
                    out.append(x)
            return out
        
        skus = _dedupe(skus)
        
        if not skus:
            st.error("SKU missing from PDF.")
            return None
        
        sku_for_filename = "_".join(skus) if skus else "UNKNOWN"
        season_value = f"{season.group(1)}{season.group(2)}" if season else "UNKNOWN"
        
        # Build results
        results = []
        row_data = {
            "Order_ID": order_id.group(1).strip() if order_id else "UNKNOWN",
            "Style": style_code.group() if style_code else "UNKNOWN",
            "Colour": colour,
            "Supplier_product_code": supplier_code.group(1).strip() if supplier_code else "UNKNOWN",
            "Item_classification": item_class_value,
            "Supplier_name": supplier_name.group(1).strip() if supplier_name else "UNKNOWN",
            "today_date": datetime.today().strftime('%d-%m-%Y'),
            "Item_name_EN": item_name_en or "",
            "Season": season_value,
            "Product_name": product_name,
            "Inner_kg": inner_kg,
            "Season_st": season_st,
            "Inner_qty": inner_qty,
            "Outer_qty": outer_qty,
            "Pictogram": pictogram,
            "Promotional": promotional,
            "_temp_sku_for_filename": sku_for_filename
        }
        
        # Add TC numbers (st1 to st7)
        for i in range(7):
            col_name = f"TC_Number_st{i+1}"
            row_data[col_name] = all_tc_numbers[i] if i < len(all_tc_numbers) else ""
        
        # Add Barcodes (st1 to st7)
        for i in range(7):
            col_name = f"Barcode_st{i+1}"
            row_data[col_name] = all_barcodes[i] if i < len(all_barcodes) else ""
        
        results.append(row_data)
        
        return results
    
    except Exception as e:
        st.error(f"PDF error: {str(e)}")
        return None


# ================================================================
#  MAIN PROCESSOR FUNCTION
# ================================================================

def process_pepco_pdf(uploaded_pdf, extra_order_ids: str | None = None):
    """Main pipeline: parse PDF, build DF, export CSV."""
    
    if not uploaded_pdf:
        return
    
    result_data = extract_data_from_pdf(uploaded_pdf)
    if not result_data:
        return
    
    df = pd.DataFrame(result_data)
    
    # Get SKU for filename
    sku_for_filename = df['_temp_sku_for_filename'].iloc[0] if '_temp_sku_for_filename' in df.columns else "UNKNOWN"
    
    if '_temp_sku_for_filename' in df.columns:
        df = df.drop(columns=['_temp_sku_for_filename'])
    
    # Merge extra Order IDs
    if extra_order_ids:
        try:
            df['Order_ID'] = df['Order_ID'].astype(str) + "+" + extra_order_ids
        except Exception:
            pass
    
    # Clean Item_name_English
    df["Item_name_English"] = df["Item_name_EN"].apply(clean_item_name_english)
    
    # Dynamic final columns with Pictogram and Promotional
    final_cols = [
        "Order_ID", "Style", "Colour", "Supplier_product_code",
        "Item_classification", "Supplier_name", "today_date",
        "Item_name_English", "Season", "Product_name", "Inner_kg",
        "Season_st", "Inner_qty", "Outer_qty",
        "Pictogram", "Promotional"
    ]
    
    # Add TC Number columns
    tc_cols = [f"TC_Number_st{i+1}" for i in range(7)]
    for col in tc_cols:
        if col in df.columns:
            final_cols.append(col)
    
    # Add Barcode columns
    barcode_cols = [f"Barcode_st{i+1}" for i in range(7)]
    for col in barcode_cols:
        if col in df.columns:
            final_cols.append(col)
    
    # Ensure all columns exist
    for col in final_cols:
        if col not in df.columns:
            df[col] = ""
    
    st.success("✅ Done!")
    st.subheader("Edit Before Download")
    
    edited_df = st.data_editor(df[final_cols])
    
    # Build CSV
    csv_buffer = StringIO()
    writer = pycsv.writer(csv_buffer, delimiter=';', quoting=pycsv.QUOTE_ALL)
    writer.writerow(final_cols)
    
    for row in edited_df.itertuples(index=False):
        writer.writerow(row)
    
    # Generate filename
    first_row_df = df.iloc[0]
    season_val = first_row_df.get("Season", "UNKNOWN").upper()
    supplier_code = first_row_df.get("Supplier_product_code", "UNKNOWN")
    style_val = first_row_df.get("Style", "UNKNOWN")
    
    custom_filename = f"PEPCO_{season_val}_{sku_for_filename}_Sticker {supplier_code}_00_{style_val}.csv"
    
    st.download_button(
        "📥 Download CSV",
        csv_buffer.getvalue().encode('utf-8-sig'),
        file_name=custom_filename,
        mime="text/csv"
    )


# ================================================================
#  PEPCO SECTION
# ================================================================

def pepco_section():
    """Main PEPCO UI section."""
    st.subheader("PEPCO Data Processing")
    
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0
    
    cols = st.columns([1, 6])
    
    with cols[0]:
        def _reset_all():
            for k in list(st.session_state.keys()):
                if k.startswith(("ui_", "pepco_", "colour_")):
                    st.session_state.pop(k, None)
            st.session_state.uploader_key += 1
        
        st.button("🆕 Upload New File", on_click=_reset_all)
    
    uploaded_pdfs = st.file_uploader(
        "Upload PEPCO Data file",
        type=["pdf"],
        key=f"pepco_uploader_{st.session_state.uploader_key}",
        accept_multiple_files=True
    )
    
    if uploaded_pdfs:
        if not isinstance(uploaded_pdfs, list):
            uploaded_pdfs = [uploaded_pdfs]
        
        primary_pdf = uploaded_pdfs[0]
        others = uploaded_pdfs[1:]
        
        other_ids = []
        for f in others:
            try:
                f.seek(0)
            except Exception:
                pass
            
            oid = extract_order_id_only(f)
            if oid:
                other_ids.append(oid)
            
            try:
                f.seek(0)
            except Exception:
                pass
        
        concatenated_ids = "+".join(other_ids) if other_ids else ""
        process_pepco_pdf(primary_pdf, extra_order_ids=concatenated_ids)


# ================================================================
#  HEADER RENDER
# ================================================================

def render_header():
    """Render logo or fallback icon."""
    left, _ = st.columns([3, 10], vertical_alignment="center")
    with left:
        if os.path.exists(LOGO_SVG):
            st.image(LOGO_SVG, width=300)
        elif os.path.exists(LOGO_PNG):
            st.image(LOGO_PNG, width=300)
        else:
            st.markdown("<div style='font-size:40px'>🏷️</div>", unsafe_allow_html=True)


# ================================================================
#  MAIN APP
# ================================================================

def main():
    st.markdown(THEME_CSS, unsafe_allow_html=True)
    render_header()
    st.title("PEPCO Automation App")
    
    if not check_password():
        st.stop()
    
    pepco_section()
    st.markdown("---")
    st.caption("This app developed by Ovi")


if __name__ == "__main__":
    main()
