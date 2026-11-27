import io
import re
import zipfile
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
from pypdf import PdfReader
import streamlit as st

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Invoice PDF to Excel Extractor (Fast)",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ Building Material Invoice PDF to Excel Extractor")
st.caption(
    "Zero API Key • Zero Rate Limit • 100+ PDFs seconds me extract karein, purane Master Excel me append karein, aur missing invoices ka alert payein."
)

# ==========================================
# REGEX PATTERNS & HELPERS
# ==========================================
DATE_PAT = r"(?:\d{1,2}[-/\.](?:[A-Za-z]{3,9}|\d{1,2})[-/\.]\d{2,4})"
NUM_PAT = r"\d+(?:,\d+)*(?:\.\d+)?"
# Supports both full (MH04 MH 0777) and short (KU 3900) vehicle numbers
VEH_PAT = r"(?:[A-Z]{2}\s*-?\s*\d{1,2}\s*-?\s*[A-Z]{1,3}\s*-?\s*\d{1,4}|[A-Z]{2,3}\s*-?\s*\d{3,4})"


def parse_date_ddmmyyyy(raw: str) -> str:
    if not raw:
        return ""
    s = raw.strip().replace(".", "-").replace("/", "-")
    for fmt in ("%d-%m-%Y", "%d-%m-%y", "%d-%b-%y", "%d-%b-%Y", "%d-%B-%y", "%d-%B-%Y"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return raw.strip()


def normalize_state(state_str: str) -> str:
    if not state_str:
        return ""
    cleaned = re.sub(r"[^a-zA-Z\s]", "", state_str).strip()
    if "maharas" in cleaned.lower():
        return "Maharashtra"
    return cleaned.title()


def normalize_metal(metal_str: str) -> str:
    if not metal_str:
        return ""
    return re.sub(r"\s+", " ", metal_str.strip().upper())


def to_float(val: str) -> float:
    return float(str(val).replace(",", "").strip())


def extract_bill_number_int(bill_no: str) -> Optional[int]:
    digits = re.findall(r"\d+", str(bill_no))
    if digits:
        return int(digits[-1])
    return None


def get_financial_year(date_str: str) -> str:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(str(date_str).strip(), fmt)
            if dt.month >= 4:
                return f"{dt.year}-{str(dt.year + 1)[-2:]}"
            else:
                return f"{dt.year - 1}-{str(dt.year)[-2:]}"
        except ValueError:
            continue
    return ""


# ==========================================
# DIRECT PYTHON PDF PARSER (NO API)
# ==========================================
def parse_pdf_invoice(
    pdf_bytes: bytes,
    source_filename: str,
    default_company_info: Dict[str, str],
) -> Tuple[dict, List[dict]]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    if not reader.pages:
        raise ValueError("PDF me koi page nahi mila.")

    p1_text = reader.pages[0].extract_text(extraction_mode="layout") or ""
    if not p1_text.strip():
        raise ValueError("Is PDF me readable text nahi hai (ye scanned image PDF ho sakta hai).")

    p1_lines = p1_text.splitlines()

    # 1
