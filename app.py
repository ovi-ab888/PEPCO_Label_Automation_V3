import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import StringIO
import csv as pycsv
from datetime import datetime, timedelta
import os
import requests

# ================================================================
# PART 1 — PAGE CONFIG & THEME
# ================================================================
st.set_page_config(page_title="PEPCO Unified Automation", page_icon="🧾", layout="wide")

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
section[data-testid="stFileUploader"], div[data-testid="stDataFrameContainer"],
div[data-testid="stVerticalBlock"]:has(> div[data-testid="stDataEditor"]){
  background:var(--card-bg)!important; border:1px solid var(--card-br)!important;
  border-radius:14px!important; padding:12px 14px;
}
label, .stMultiSelect label, .stSelectbox label, .stNumberInput label, .stTextInput label{
  color:var(--txt)!important; font-weight:500;
}
input, textarea, div[data-baseweb="select"] > div{
  color:var(--txt)!important; background:var(--input-bg)!important; border-color:var(--input-br)!important;
}
</style>
"""

# ================================================================
# PART 2 — COMMON UTILITIES (Password, Cleaning, Header)
# ================================================================
def check_password():
    expected = st.secrets.get("app_password", os.environ.get("PEPCO_APP_PASSWORD"))[cite: 1, 2]
    if not expected:
        st.error("App password not configured.")[cite: 1, 2]
        return False
    def _entered():
        if st.session_state.get("password") == expected:
            st.session_state["password_correct"] = True[cite: 1, 2]
            try: del st.session_state["password"]
            except: pass
        else:
            st.session_state["password_correct"] = False[cite: 1, 2]
    if st.session_state.get("password_correct"): return True[cite: 1, 2]
    st.text_input("Enter Your Access Code", type="password", key="password", on_change=_entered)[cite: 1, 2]
    if st.session_state.get("password_correct") is False:
        st.error("Your password Incorrect, Please contact Mr. Ovi")[cite: 1, 2]
    return False

def clean_item_name_english(name: str) -> str:
    if not isinstance(name, str): return ""[cite: 1, 2]
    text = name.strip()
    prefixes = ["xxxxx"] # Add your prefixes here[cite: 1, 2]
    for p in prefixes:
        if text.lower().startswith(p):
            text = text[len(p):].strip(" -_,./").strip()[cite: 1, 2]
            break
    return text.upper()[cite: 1, 2]

def render_header():
    left, _ = st.columns([3, 10], vertical_alignment="center")[cite: 1, 2]
    with left:
        if os.path.exists(LOGO_SVG): st.image(LOGO_SVG, width=300)[cite: 1, 2]
        elif os.path.exists(LOGO_PNG): st.image(LOGO_PNG, width=300)[cite: 1, 2]
        else: st.markdown("<div style='font-size:40px'>🏷️</div>", unsafe_allow_html=True)[cite: 1, 2]

def extract_order_id_only(file):
    try:
        file.seek(0)
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            page1_text = doc[0].get_text() if len(doc) > 0 else ""[cite: 1, 2]
        file.seek(0)
        m = re.search(r"Order\s*-\s*ID\s*\.{2,}\s*([A-Z0-9_+-]+)", page1_text, re.IGNORECASE)[cite: 1, 2]
        return m.group(1).strip() if m else None[cite: 1, 2]
    except: return None

# ================================================================
# PART 3 — STICKER MODULE (Source 1 Logic)
# ================================================================
def sticker_extract_logic(pages_text):
    tc_list = []
    barcode_list = []
    if len(pages_text) >= 4:[cite: 1]
        for i in range(3, len(pages_text)):[cite: 1]
            page_text = pages_text[i]
            tc_list.extend(re.findall(r"TC\s*[-:.]?\s*(T\d+)", page_text, re.IGNORECASE))[cite: 1]
            barcode_list.extend(re.findall(r"\b\d{13}\b", page_text))[cite: 1]
    return list(dict.fromkeys(tc_list))[:7], list(dict.fromkeys(barcode_list))[:7][cite: 1]

def sticker_pdf_engine(file):
    try:
        raw = file.read()
        doc = fitz.open(stream=raw, filetype="pdf")[cite: 1]
        pages_text = [doc[i].get_text() for i in range(len(doc))][cite: 1]
        full_text, page1 = "\n".join(pages_text), pages_text[0][cite: 1]
        
        tc_nums, b_codes = sticker_extract_logic(pages_text)[cite: 1]
        
        row_data = {
            "Order_ID": re.search(r"Order\s*-\s*ID\s*\.{2,}\s*(.+)", page1).group(1).strip() if re.search(r"Order\s*-\s*ID\s*\.{2,}\s*(.+)", page1) else "UNKNOWN",[cite: 1]
            "Style": re.search(r"\b\d{6}\b", page1).group() if re.search(r"\b\d{6}\b", page1) else "UNKNOWN",[cite: 1]
            "today_date": datetime.today().strftime('%d-%m-%Y'),[cite: 1]
            "Item_name_English": clean_item_name_english(re.search(r"Item\s*name\s*English\s*[:\.]{1,}\s*(.+)", full_text, re.IGNORECASE).group(1).strip() if re.search(r"Item\s*name\s*English\s*[:\.]{1,}\s*(.+)", full_text, re.IGNORECASE) else ""),[cite: 1]
            "Season": re.search(r"Season\s*\.{2,}\s*(\w+)?\s*(\d{2})", page1).group(0) if re.search(r"Season\s*\.{2,}\s*(\w+)?\s*(\d{2})", page1) else "UNKNOWN"[cite: 1]
        }
        
        for i in range(7):
            row_data[f"TC_Number_st{i+1}"] = tc_nums[i] if i < len(tc_nums) else ""[cite: 1]
            row_data[f"Barcode_st{i+1}"] = b_codes[i] if i < len(b_codes) else ""[cite: 1]
            
        return [row_data]
    except Exception as e:
        st.error(f"Sticker PDF Error: {e}")
        return None

# ================================================================
# PART 4 — SWINGTAG MODULE (Source 2 Logic)
# ================================================================
WASHING_CODES = {'1': '১২৩৪৫', '2': '১৪৭৮৫', '3': 'djnst', '4': 'djnpt', '5': 'djnqt', '6': 'djnqt', '7': 'gjnpt', '8': 'gjnpu', '9': 'gjnqt', '10': 'gjnqu', '11': 'ijnst', '12': 'ijnsu', '13': 'ijnpu', '14': 'ijnsv', '15': 'djnsw'}[cite: 2]

@st.cache_data(ttl=600)
def load_swingtag_data(url):
    try: return pd.read_csv(url)[cite: 2]
    except: return pd.DataFrame()

def swingtag_pdf_engine(file):
    try:
        raw = file.read()
        doc = fitz.open(stream=raw, filetype="pdf")[cite: 2]
        pages_text = [doc[i].get_text() for i in range(len(doc))][cite: 2]
        page1 = pages_text[0][cite: 2]
        
        skus = list(dict.fromkeys(re.findall(r"\b\d{8}\b", "\n".join(pages_text))))[cite: 2]
        barcodes = list(dict.fromkeys(re.findall(r"\b\d{13}\b", "\n".join(pages_text))))[cite: 2]
        
        results = []
        for sku, barcode in zip(skus, barcodes):[cite: 2]
            results.append({
                "Order_ID": re.search(r"Order\s*-\s*ID\s*\.{2,}\s*(.+)", page1).group(1).strip() if re.search(r"Order\s*-\s*ID\s*\.{2,}\s*(.+)", page1) else "UNKNOWN",[cite: 2]
                "Style": re.search(r"\b\d{6}\b", page1).group() if re.search(r"\b\d{6}\b", page1) else "UNKNOWN",[cite: 2]
                "today_date": datetime.today().strftime('%d-%m-%Y'),[cite: 2]
                "barcode": barcode,[cite: 2]
                "SKU_Info": sku,[cite: 2]
                "Item_classification": re.search(r"Item classification\s*\.{2,}\s*(.+)", page1).group(1).strip() if re.search(r"Item classification\s*\.{2,}\s*(.+)", page1) else "UNKNOWN"[cite: 2]
            })
        return results
    except Exception as e:
        st.error(f"Swingtag PDF Error: {e}")
        return None

# ================================================================
# PART 5 — MAIN APP NAVIGATION
# ================================================================
def main():
    st.markdown(THEME_CSS, unsafe_allow_html=True)[cite: 1, 2]
    render_header()[cite: 1, 2]
    st.title("PEPCO Automation Unified App")[cite: 1, 2]
    
    if not check_password(): st.stop()[cite: 1, 2]
    
    st.sidebar.title("App Mode")
    mode = st.sidebar.radio("Select Automation Type:", ["Sticker Automation (Source 1)", "Swingtag Automation (Source 2)"])[cite: 1, 2]
    
    if "uploader_key" not in st.session_state: st.session_state.uploader_key = 0
    if st.sidebar.button("🆕 Reset & New Upload"):
        for k in list(st.session_state.keys()): 
            if k not in ["password_correct"]: st.session_state.pop(k, None)
        st.session_state.uploader_key += 1
        st.rerun()

    uploaded_files = st.file_uploader("Upload PDF Data Files", type=["pdf"], key=f"up_{st.session_state.uploader_key}", accept_multiple_files=True)[cite: 1, 2]

    if uploaded_files:
        if mode == "Sticker Automation (Source 1)":
            st.subheader("Sticker Mode Processing")
            data = sticker_pdf_engine(uploaded_files[0])[cite: 1]
            if data:
                df = pd.DataFrame(data)
                edited_df = st.data_editor(df)[cite: 1]
                # Add CSV Download Logic here same as Source 1[cite: 1]
        
        else:
            st.subheader("Swingtag Mode Processing")
            # Implement Material composition and Price Ladder Logic from Source 2 here[cite: 2]
            data = swingtag_pdf_engine(uploaded_files[0])[cite: 2]
            if data:
                df = pd.DataFrame(data)
                st.data_editor(df)[cite: 2]

    st.markdown("---")
    st.caption("Developed by Ovi")[cite: 1, 2]

if __name__ == "__main__":
    main()
