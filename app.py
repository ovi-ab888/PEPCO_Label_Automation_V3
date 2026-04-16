# ================================================================
# PART 1 — PAGE CONFIG + IMPORTS + THEME + PASSWORD + CONSTANTS
# ================================================================

# ---------- PAGE CONFIG (must be at top) ----------
import streamlit as st
st.set_page_config(
    page_title="PEPCO",
    page_icon="🧾",
    layout="wide"
)

# ---------- Imports ----------
import fitz  # PyMuPDF
import pandas as pd
import re
from io import StringIO
import csv as pycsv
from datetime import datetime, timedelta
import os
import requests


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

    # Prefer streamlit secrets
    try:
        expected = st.secrets.get("app_password", None)
    except Exception:
        expected = None

    # Fallback env variable
    if expected is None:
        expected = os.environ.get("PEPCO_APP_PASSWORD")

    # If not found → error
    if expected is None:
        st.error("App password not configured. Please set 'app_password' in secrets or PEPCO_APP_PASSWORD env var.")
        return False

    # When password typed
    def _password_entered():
        if st.session_state.get("password") == expected:
            st.session_state["password_correct"] = True
            try:
                del st.session_state["password"]
            except Exception:
                pass
        else:
            st.session_state["password_correct"] = False

    # Already correct?
    if st.session_state.get("password_correct", None) is True:
        return True

    # Input box
    st.text_input("Enter Your Access Code", type="password", key="password", on_change=_password_entered)

    # Wrong
    if st.session_state.get("password_correct") is False:
        st.error("Your password Incorrect,  Please contact Mr. Ovi")

    return False


# ================================================================
#  PRODUCT TRANSLATION LOADER (only needed for DEPT mapping)
# ================================================================
@st.cache_data(ttl=600)
def load_product_translations():
    """Load product name translations from Google Sheet."""
    try:
        sheet_id = "1ue68TSJQQedKa7sVBB4syOc0OXJNaLS7p9vSnV52mKA"
        sheet_name = "SS26 Product_Name"
        encoded = requests.utils.quote(sheet_name)

        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={encoded}"
        df = pd.read_csv(url)

        if df.empty:
            st.error("Loaded translations but sheet appears empty")

        return df

    except Exception as e:
        st.error(f"❌ Failed to load translations: {str(e)}")
        return pd.DataFrame()


# ================================================================
#  NEW EXTRACTION FUNCTIONS FOR B PART
# ================================================================

def extract_tc_number(text):
    """Extract TC number from PDF text."""
    m = re.search(r"TC\s*-\s*(T\d+)", text, re.IGNORECASE)
    if not m:
        m = re.search(r"TC\s*[:.]?\s*(T\d+)", text, re.IGNORECASE)
    return m.group(1) if m else ""


def extract_product_name(text):
    """Extract product name from PDF text."""
    m = re.search(r"ITEM\s*\d+\s*\n\s*(.+)", text, re.IGNORECASE)
    if not m:
        m = re.search(r"Product\s*name\s*[:.]?\s*(.+)", text, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def extract_barcode(text):
    """Extract barcode (13 digits) from PDF text."""
    m = re.search(r"\b\d{13}\b", text)
    return m.group(0) if m else ""


def extract_inner_kg(text):
    """Extract inner kg from PDF text."""
    m = re.search(r"MAX\.?\s*(\d+)\s*kg", text, re.IGNORECASE)
    if not m:
        m = re.search(r"(\d+)\s*kg", text, re.IGNORECASE)
    return f"MAX. {m.group(1)} kg" if m else ""


def extract_season(text):
    """Extract season code (AW/SS/FW/SW + year) from PDF text."""
    m = re.search(r"\b(AW|SS|FW|SW)\d{2}\b", text, re.IGNORECASE)
    return m.group(0).upper() if m else ""


def extract_inner_qty(text):
    """Extract inner quantity from PDF text."""
    m = re.search(r"(\d+)\s*Pcs", text, re.IGNORECASE)
    return f"{m.group(1)} Pcs" if m else ""


def extract_outer_qty(text):
    """Extract outer quantity from PDF text."""
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
#  HELPER FUNCTIONS
# ================================================================

# ---------- Classification → mapping ----------
def get_classification_type(item_class):
    """Determine class type key."""
    if not item_class:
        return None

    ic = item_class.lower()

    if 'younger girls outerwear' in ic:
        return 'yg'
    if 'older girls outerwear' in ic:
        return 'og'
    if 'younger boys outerwear' in ic:
        return 'yb'
    if 'older boys outerwear' in ic:
        return 'ob'
    if 'baby girls outerwear' in ic:
        return 'a'
    if 'baby boys outerwear' in ic:
        return 'b'
    if 'baby girls essentials' in ic:
        return 'd_girls'
    if 'baby boys essentials' in ic:
        return 'd'
    if 'ladies outerwear' in ic:
        return 'l'
    if 'mens outerwear' in ic:
        return 'm'

    return None


# ---------- Map Item_classification → Dept label ----------
def map_item_class_to_dept_label(item_class):
    """Map item_class text to UI Department names."""
    if not item_class:
        return None

    ic = item_class.lower()

    if 'baby boys outerwear' in ic or 'baby boys essentials' in ic:
        return "Baby Boy"
    if 'baby girls outerwear' in ic or 'baby girls essentials' in ic:
        return "Baby Girl"
    if 'younger boys outerwear' in ic or 'older boys outerwear' in ic:
        return "Boys"
    if 'younger girls outerwear' in ic or 'older girls outerwear' in ic:
        return "Girls"
    if 'ladies outerwear' in ic:
        return "Women"
    if 'mens outerwear' in ic:
        return "Mens"

    return None


# ---------- Map Item_classification → DEPT column ----------
def get_dept_value(item_class):
    """Maps classification → BABY / KIDS / TEENS / WOMEN / MEN."""
    if not item_class:
        return ""

    ic = item_class.lower()

    if any(x in ic for x in ['baby boys', 'baby girls']):
        return "BABY"
    if any(x in ic for x in ['younger boys', 'younger girls']):
        return "KIDS"
    if any(x in ic for x in ['older girls', 'older boys']):
        return "TEENS"
    if 'ladies outerwear' in ic:
        return "WOMEN"
    if 'mens outerwear' in ic:
        return "MEN"

    return ""


# ---------- Item_name_EN cleaning ----------
def clean_item_name_english(name: str) -> str:
    """
    Item_name_EN থেকে prefix গুলো বাদ দিয়ে বাকি অংশ CAPITAL LETTERS এ রিটার্ন করবে।
    """
    if not isinstance(name, str):
        return ""

    text = name.strip()
    lower = text.lower()

    prefixes = [
        "xxxxx",
        "xxxxx",
        "xxxxx",
        "xxxxx",
        "xxxxx",
        "xxxxx",
        "xxxxx",
        "xxxxx",
    ]

    for p in prefixes:
        if lower.startswith(p):
            cut_len = len(p)
            text = text[cut_len:].strip(" -_,./").strip()
            break

    return text.upper()


# ================================================================
# PART 2 — PDF EXTRACTION + COLOUR SYSTEM + NEW EXTRACTIONS
# ================================================================

# ================================================================
#  COLOUR EXTRACTION (multiple PDF layout compatible)
# ================================================================
def extract_colour_from_page2(text, page_number=1):
    """Old function: Extract colour from page2."""
    try:
        m = re.search(
            r"Colour[^\n]*?\n\s*([A-Za-z]+)\s+([0-9]{2}-[0-9]{4}[A-Za-z]*)",
            text,
            re.IGNORECASE
        )
        if m:
            colour_name = m.group(1).strip().upper()
            pantone = m.group(2).strip().upper()
            return f"{colour_name} {pantone}"
        return "UNKNOWN"
    except Exception:
        return "UNKNOWN"


def extract_colour_from_pdf_pages(pages_text):
    """
    Ultra-robust PEPCO Colour Detection
    Supports:
        ✔ Old 6-page PDF format
        ✔ New 5-page PDF format
        ✔ Broken layout (Colour row + size row merged)
        ✔ Missing pantone
    """
    # -------- 1️⃣ Standard Colour Table --------
    for txt in pages_text:
        m = re.search(
            r"Colour.*?\n.*?\n\s*([A-Za-z ]+)\s+[0-9]{2}-[0-9]{4}",
            txt,
            re.IGNORECASE | re.DOTALL
        )
        if m:
            return m.group(1).strip().upper()

    # -------- 2️⃣ Purchase Price block --------
    for txt in pages_text:
        m2 = re.search(
            r"Purchase price.*?\n\s*([A-Za-z ]+)\s+[0-9]{2}-[0-9]{4}",
            txt,
            re.IGNORECASE | re.DOTALL
        )
        if m2:
            return m2.group(1).strip().upper()

    # -------- 3️⃣ Generic fallback using "colour" keyword --------
    for txt in pages_text:
        if "colour" in txt.lower():
            for line in txt.splitlines():
                if re.search(r"[A-Za-z ]+\s+[0-9]{2}-[0-9]{4}", line):
                    name = line.split()[0:-1]
                    if name:
                        return " ".join(name).upper()

    # -------- 4️⃣ Manual input fallback --------
    st.warning("⚠️ Colour not found in PDF. Enter colour manually:")
    manual = st.text_input("Colour (e.g. WHITE):", key="manual_colour_fix")
    return manual.strip().upper() if manual else "UNKNOWN"


# ================================================================
#  EXTRACT ORDER ID FROM PDF (for multiple uploads)
# ================================================================
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

    m = re.search(
        r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)",
        page1_text,
        re.IGNORECASE
    )
    return m.group(1).strip() if m else None


# ================================================================
# MAIN PDF EXTRACTION ENGINE (UPDATED - SKU removed from CSV)
# ================================================================
def extract_data_from_pdf(file):
    """Robust PEPCO extractor with SKU for filename only."""
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

        # ---------------- EXTRACT NEW FIELDS (B PART) ----------------
        tc_number = extract_tc_number(full_text)
        product_name = extract_product_name(full_text)
        barcode = extract_barcode(full_text)
        inner_kg = extract_inner_kg(full_text)
        season_st = extract_season(full_text)
        inner_qty = extract_inner_qty(full_text)
        outer_qty = extract_outer_qty(full_text)

        # ---------------- Item Name EN ----------------
        item_name_en = None

        m_item = re.search(
            r"Item\s*name\s*English\s*[:\.]{1,}\s*(.+)",
            full_text,
            re.IGNORECASE
        )
        if not m_item:
            m_item = re.search(
                r"Item\s*name\s*[:\.]{1,}\s*(.+?)\n",
                full_text,
                re.IGNORECASE
            )
        if m_item:
            item_name_en = m_item.group(1).strip()

        # ---------------- Identifiers ----------------
        merch_code = re.search(r"Merch\s*code\s*\.{2,}\s*([\w/]+)", page1)
        season = re.search(r"Season\s*\.{2,}\s*(\w+)?\s*(\d{2})", page1)
        style_code = re.search(r"\b\d{6}\b", page1)

        order_id = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*(.+)", page1)
        item_class = re.search(r"Item classification\s*\.{2,}\s*(.+)", page1)
        supplier_code = re.search(r"Supplier product code\s*\.{2,}\s*(.+)", page1)
        supplier_name = re.search(r"Supplier name\s*\.{2,}\s*(.+)", page1)

        item_class_value = item_class.group(1).strip() if item_class else "UNKNOWN"

        # ---------------- AUTO COLOUR EXTRACTION ----------------
        colour = extract_colour_from_pdf_pages(pages_text)

        # ---------------- SKU extraction (ONLY FOR FILENAME, NOT IN CSV) ----------------
        skus = []
        for txt in pages_text:
            skus.extend(re.findall(r"\b\d{8}\b", txt))

        # Dedupe SKUs
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

        # Store SKUs for filename (will be removed from final CSV)
        sku_for_filename = "_".join(skus) if skus else "UNKNOWN"

        season_value = (
            f"{season.group(1)}{season.group(2)}"
            if season else "UNKNOWN"
        )

        # ---------------- BUILD RESULT (WITHOUT SKU in output) ----------------
        results = []
        for sku in skus:
            results.append({
                "Order_ID": order_id.group(1).strip() if order_id else "UNKNOWN",
                "Style": style_code.group() if style_code else "UNKNOWN",
                "Colour": colour,
                "Supplier_product_code": supplier_code.group(1).strip() if supplier_code else "UNKNOWN",
                "Item_classification": item_class_value,
                "Supplier_name": supplier_name.group(1).strip() if supplier_name else "UNKNOWN",
                "today_date": datetime.today().strftime('%d-%m-%Y'),
                "Item_name_EN": item_name_en or "",
                "Season": season_value,
                # NEW FIELDS (B PART)
                "TC_Number_st1": tc_number,
                "Product_name": product_name,
                "Barcode_st1": barcode,
                "Inner_kg": inner_kg,
                "Season_st": season_st,
                "Inner_qty": inner_qty,
                "Outer_qty": outer_qty,
                # SKU stored temporarily for filename (will be removed)
                "_temp_sku_for_filename": sku_for_filename
            })

        return results

    except Exception as e:
        st.error(f"PDF error: {str(e)}")
        return None


# ================================================================
#  MAIN PROCESSOR + UI SECTION + APP ENTRY (UPDATED)
# ================================================================

def process_pepco_pdf(uploaded_pdf, extra_order_ids: str | None = None):
    """Main pipeline: parse PDF, build DF, export CSV."""
    # ----- Load reference data -----
    translations_df = load_product_translations()

    if not uploaded_pdf:
        return

    # ----- Parse PDF to structured data -----
    result_data = extract_data_from_pdf(uploaded_pdf)
    if not result_data:
        return

    df = pd.DataFrame(result_data)

    # ----- Get SKU for filename before removing it -----
    sku_for_filename = df['_temp_sku_for_filename'].iloc[0] if '_temp_sku_for_filename' in df.columns else "UNKNOWN"
    
    # ----- Remove temporary SKU column (not needed in CSV) -----
    if '_temp_sku_for_filename' in df.columns:
        df = df.drop(columns=['_temp_sku_for_filename'])

    # ----- Base values from first row -----
    first_row = result_data[0] if len(result_data) > 0 else {}
    pdf_item_class = first_row.get("Item_classification", "")
    pdf_item_name_en = (first_row.get("Item_name_EN") or "").strip()

    # ----- Merge extra Order IDs from other PDFs -----
    if extra_order_ids:
        try:
            df['Order_ID'] = df['Order_ID'].astype(str) + "+" + extra_order_ids
        except Exception:
            pass

    # ============================================================
    #  UI Controls (Department only)
    # ============================================================
    c1, _ = st.columns([2, 3])

    # -- Department select (default from item_class) --
    depts = translations_df['DEPARTMENT'].dropna().unique().tolist() if not translations_df.empty else []
    default_dept_label = map_item_class_to_dept_label(pdf_item_class)
    default_dept_index = 0

    if default_dept_label and depts:
        for i, d in enumerate(depts):
            if str(d).strip().lower() == str(default_dept_label).strip().lower():
                default_dept_index = i
                break

    with c1:
        selected_dept = st.selectbox(
            "Select Department",
            options=depts if depts else ["No Data"],
            index=default_dept_index if depts else 0,
            key="ui_dept"
        )

    # ============================================================
    #  DataFrame enrichment (Dept, Item_name_English)
    # ============================================================
    df['Dept'] = df['Item_classification'].apply(get_dept_value)

    # Clean Item_name_English
    df["Item_name_English"] = df["Item_name_EN"].apply(clean_item_name_english)

    # ============================================================
    #  FINAL COLUMNS (NO SKU COLUMN)
    # ============================================================
    final_cols = [
        "Order_ID",
        "Style",
        "Colour",
        "Supplier_product_code",
        "Item_classification",
        "Supplier_name",
        "today_date",
        "Item_name_English",
        "Season",
        "Dept",
        # NEW FIELDS (B PART)
        "TC_Number_st1",
        "Product_name",
        "Barcode_st1",
        "Inner_kg",
        "Season_st",
        "Inner_qty",
        "Outer_qty"
    ]

    # Ensure all columns exist
    for col in final_cols:
        if col not in df.columns:
            df[col] = ""

    st.success("✅ Done!")
    st.subheader("Edit Before Download")

    edited_df = st.data_editor(df[final_cols])

    # Build CSV with ; separator & quoted fields
    csv_buffer = StringIO()
    writer = pycsv.writer(
        csv_buffer,
        delimiter=';',
        quoting=pycsv.QUOTE_ALL
    )
    writer.writerow(final_cols)

    for row in edited_df.itertuples(index=False):
        writer.writerow(row)

    # ---------- Custom CSV filename (USING SKU) ----------
    first_row_df = df.iloc[0]
    season_val = first_row_df.get("Season", "UNKNOWN").upper()

    # Use the SKU we saved earlier for filename
    sku_val = sku_for_filename

    supplier_code = first_row_df.get("Supplier_product_code", "UNKNOWN")
    style_val = first_row_df.get("Style", "UNKNOWN")

    custom_filename = (
        f"PEPCO_{season_val}_{sku_val}_Sticker "
        f"{supplier_code}_00_{style_val}.csv"
    )

    st.download_button(
        "📥 Download CSV",
        csv_buffer.getvalue().encode('utf-8-sig'),
        file_name=custom_filename,
        mime="text/csv"
    )

# ================================================================
#  PEPCO SECTION (Uploader + Reset)
# ================================================================
def pepco_section():
    """Main PEPCO UI section (upload + reset + extra order IDs merge)."""
    st.subheader("PEPCO Data Processing")

    # One-time init for uploader key
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    cols = st.columns([1, 6])

    # Reset / new upload button
    with cols[0]:
        def _reset_all():
            # Clear only app-related session keys
            for k in list(st.session_state.keys()):
                if k.startswith(("ui_", "mat_", "pepco_", "colour_", "colour_manual_", "colour_missing_")):
                    st.session_state.pop(k, None)

            # Force uploader refresh
            st.session_state.uploader_key += 1

        st.button("🆕 Upload New File", on_click=_reset_all)

    # File uploader (multi PDF)
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

        # Collect Order_ID from additional PDFs
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
            st.markdown(
                "<div style='font-size:40px'>🏷️</div>",
                unsafe_allow_html=True
            )


# ================================================================
#  MAIN APP
# ================================================================
def main():
    # Apply theme
    st.markdown(THEME_CSS, unsafe_allow_html=True)

    # Header + Title
    render_header()
    st.title("PEPCO Automation App")

    # Password gate
    if not check_password():
        st.stop()

    # Main content
    pepco_section()

    st.markdown("---")
    st.caption("This app developed by Ovi")


# ================================================================
#  ENTRY POINT
# ================================================================
if __name__ == "__main__":
    main()
