import io
import re
from datetime import datetime
from typing import List, Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
from pydantic import BaseModel, Field
import streamlit as st
from google import genai
from google.genai import types

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Invoice PDF to Excel Extractor",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Building Material Invoice PDF to Excel Extractor")
st.caption(
    "Multiple Invoice PDFs upload karein, purane Master Excel me naya data append karein, aur missing invoices ka alert payein."
)

# ==========================================
# PYDANTIC SCHEMA FOR STRUCTURED EXTRACTION
# ==========================================
class MetalRateItem(BaseModel):
    metal_type: str = Field(description="Metal/Product name e.g., 10 MM, 20 MM, C. SAND")
    total_ton: Optional[float] = Field(default=None, description="Total TON for this metal")
    rate: Optional[float] = Field(default=None, description="Rate per TON")
    amount: Optional[float] = Field(default=None, description="Amount for this metal")


class DispatchItem(BaseModel):
    sr_no: Optional[str] = Field(default="", description="Serial number of dispatch row")
    dispatch_date: Optional[str] = Field(default="", description="Date material was sent in DD/MM/YYYY format")
    challan_no: Optional[str] = Field(default="", description="Challan or slip number if available")
    vehicle_no: Optional[str] = Field(default="", description="Vehicle number e.g., MH04 MH 0777")
    metal_type: str = Field(description="Metal type e.g., 10 MM, 20 MM, C. SAND")
    ton: float = Field(description="Weight in TON")
    rate: Optional[float] = Field(default=None, description="Rate per TON if mentioned on this row")
    amount: Optional[float] = Field(default=None, description="Amount if mentioned on this row")


class InvoiceData(BaseModel):
    company_name: str = Field(description="Supplier company name from header")
    company_gstin: Optional[str] = Field(default="", description="Supplier GSTIN")
    company_pan: Optional[str] = Field(default="", description="Supplier PAN")
    company_details: Optional[str] = Field(default="", description="Supplier address, contact person, phone numbers")
    bill_no: str = Field(description="Invoice / Bill number e.g., 01, 04")
    invoice_date: str = Field(description="Invoice date formatted as DD/MM/YYYY")
    bill_to_name: str = Field(description="Customer name under Bill To")
    bill_to_address: Optional[str] = Field(default="", description="Customer billing address")
    shipping_address: Optional[str] = Field(default="", description="Shipping address if mentioned, else blank")
    customer_gst_pan: Optional[str] = Field(default="", description="Customer GST No or PAN No")
    customer_state: str = Field(description="Customer State name, corrected spelling e.g., Maharashtra")
    hsn_code: Optional[str] = Field(default="", description="HSN Code e.g., 2517")
    round_off: Optional[float] = Field(default=0.0, description="Round off amount if shown")
    summary_rates: List[MetalRateItem] = Field(
        default_factory=list,
        description="Summary table of Metal types and their Rates from Page 1",
    )
    dispatch_items: List[DispatchItem] = Field(
        default_factory=list,
        description=(
            "All individual vehicle/date/metal delivery rows. "
            "If invoice has a 2nd page challan table, extract every row from Page 2. "
            "If vehicle/date/metal rows are inside Page 1 invoice table directly, extract every row from Page 1."
        ),
    )


# ==========================================
# HELPER FUNCTIONS
# ==========================================
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


def parse_invoice_with_gemini(pdf_bytes: bytes, api_key: str, model_name: str) -> InvoiceData:
    client = genai.Client(api_key=api_key)
    prompt = """
    Extract all structured invoice data from this PDF accurately.
    Important rules:
    1. Convert all dates (Invoice Date and Dispatch Dates) strictly to DD/MM/YYYY format (e.g., 7/4/2026
