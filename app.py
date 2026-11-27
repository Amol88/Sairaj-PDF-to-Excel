import io
import json
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


def parse_invoice_with_gemini(pdf_bytes: bytes, api_key: str) -> InvoiceData:
    client = genai.Client(api_key=api_key)
    prompt = """
    Extract all structured invoice data from this PDF accurately.
    Important rules:
    1. Convert all dates (Invoice Date and Dispatch Dates) strictly to DD/MM/YYYY format (e.g., 7/4/2026 -> 07/04/2026, 27-Mar-26 -> 27/03/2026).
    2. Normalize state spelling (e.g., 'Maharastra.' -> 'Maharashtra').
    3. Look carefully at the layout:
       - Layout A (2-page invoice): Page 1 has summary rates by metal (10 MM, 20 MM, C. SAND), and Page 2 has the full Challan/Dispatch table (Sr.No., Date, Challan No, Vehicle No, Metal, TON). Extract Page 1 rates into `summary_rates` and ALL Page 2 rows into `dispatch_items`.
       - Layout B (1-page invoice): Page 1's main table directly contains Date, Vehicle No, Metal, TON, Rate, Amount for each delivery. Extract ALL rows into `dispatch_items` (and also capture unique metal rates in `summary_rates`).
    4. Watch for wrapped text in date/sr.no cells on Page 2 (e.g., '-Apr-26 3 24' means Sr No 24, Date 03/04/2026).
    5. Do not skip any dispatch row.
    """

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=InvoiceData,
            temperature=0.0,
        ),
    )
    return InvoiceData.model_validate_json(response.text)


def process_extracted_invoice(inv: InvoiceData, source_filename: str):
    """Converts extracted InvoiceData into Summary dict and Dispatch list with exact GST calculations."""
    state = normalize_state(inv.customer_state)
    is_maharashtra = state.lower() == "maharashtra"
    fy = get_financial_year(inv.invoice_date)
    bill_num_int = extract_bill_number_int(inv.bill_no)
    bill_no_clean = str(inv.bill_no).strip()

    # Build metal -> rate map from summary_rates
    rate_map = {}
    for item in inv.summary_rates:
        if item.rate is not None:
            rate_map[normalize_metal(item.metal_type)] = float(item.rate)

    dispatch_rows = []
    total_ton = 0.0
    gross_amount = 0.0

    # If dispatch_items exist, process each trip; otherwise fallback to summary_rates
    if inv.dispatch_items:
        for idx, d in enumerate(inv.dispatch_items, start=1):
            metal = normalize_metal(d.metal_type)
            ton = float(d.ton or 0.0)
            rate = float(d.rate) if d.rate is not None else rate_map.get(metal, 0.0)
            line_amount = round(ton * rate, 2)

            if is_maharashtra:
                cgst = round(line_amount * 0.025, 2)
                sgst = round(line_amount * 0.025, 2)
                igst = 0.0
            else:
                cgst = 0.0
                sgst = 0.0
                igst = round(line_amount * 0.05, 2)

            line_total = round(line_amount + cgst + sgst + igst, 2)
            total_ton += ton
            gross_amount += line_amount

            dispatch_rows.append(
                {
                    "Financial Year": fy,
                    "Bill No": bill_no_clean,
                    "Bill No (Num)": bill_num_int,
                    "Invoice Date": inv.invoice_date,
                    "Bill To Name": inv.bill_to_name,
                    "Customer State": state,
                    "Sr. No.": d.sr_no if d.sr_no else str(idx),
                    "Dispatch Date": d.dispatch_date,
                    "Challan No": d.challan_no,
                    "Vehicle No": d.vehicle_no,
                    "Metal Type": metal,
                    "TON": round(ton, 2),
                    "Rate": round(rate, 2),
                    "Amount (TON x Rate)": line_amount,
                    "CGST (2.5%)": cgst,
                    "SGST (2.5%)": sgst,
                    "IGST (5%)": igst,
                    "Line Total": line_total,
                    "Source File": source_filename,
                }
            )
    else:
        for idx, s in enumerate(inv.summary_rates, start=1):
            metal = normalize_metal(s.metal_type)
            ton = float(s.total_ton or 0.0)
            rate = float(s.rate or 0.0)
            line_amount = round(ton * rate, 2)

            if is_maharashtra:
                cgst = round(line_amount * 0.025, 2)
                sgst = round(line_amount * 0.025, 2)
                igst = 0.0
            else:
                cgst = 0.0
                sgst = 0.0
                igst = round(line_amount * 0.05, 2)

            line_total = round(line_amount + cgst + sgst + igst, 2)
            total_ton += ton
            gross_amount += line_amount

            dispatch_rows.append(
                {
                    "Financial Year": fy,
                    "Bill No": bill_no_clean,
                    "Bill No (Num)": bill_num_int,
                    "Invoice Date": inv.invoice_date,
                    "Bill To Name": inv.bill_to_name,
                    "Customer State": state,
                    "Sr. No.": str(idx),
                    "Dispatch Date": inv.invoice_date,
                    "Challan No": "",
                    "Vehicle No": "",
                    "Metal Type": metal,
                    "TON": round(ton, 2),
                    "Rate": round(rate, 2),
                    "Amount (TON x Rate)": line_amount,
                    "CGST (2.5%)": cgst,
                    "SGST (2.5%)": sgst,
                    "IGST (5%)": igst,
                    "Line Total": line_total,
                    "Source File": source_filename,
                }
            )

    gross_amount = round(gross_amount, 2)
    if is_maharashtra:
        inv_cgst = round(gross_amount * 0.025, 2)
        inv_sgst = round(gross_amount * 0.025, 2)
        inv_igst = 0.0
    else:
        inv_cgst = 0.0
        inv_sgst = 0.0
        inv_igst = round(gross_amount * 0.05, 2)

    raw_total = gross_amount + inv_cgst + inv_sgst + inv_igst
    grand_total = round(raw_total)
    round_off = round(grand_total - raw_total, 2)

    summary_row = {
        "Financial Year": fy,
        "Bill No": bill_no_clean,
        "Bill No (Num)": bill_num_int,
        "Invoice Date": inv.invoice_date,
        "Company Name": inv.company_name,
        "Company GSTIN": inv.company_gstin,
        "Company PAN": inv.company_pan,
        "Company Details": inv.company_details,
        "Bill To Name": inv.bill_to_name,
        "Bill To Address": inv.bill_to_address,
        "Shipping Address": inv.shipping_address,
        "Customer GST/PAN": inv.customer_gst_pan,
        "Customer State": state,
        "HSN Code": inv.hsn_code,
        "Total TON": round(total_ton, 2),
        "Gross Amount": gross_amount,
        "CGST (2.5%)": inv_cgst,
        "SGST (2.5%)": inv_sgst,
        "IGST (5%)": inv_igst,
        "Round Off": round_off,
        "Grand Total": float(grand_total),
        "Source File": source_filename,
    }

    return summary_row, dispatch_rows


def find_missing_invoices(summary_df: pd.DataFrame) -> pd.DataFrame:
    if summary_df.empty or "Bill No (Num)" not in summary_df.columns:
        return pd.DataFrame(columns=["Missing Bill No", "Status"])

    valid_nums = (
        summary_df["Bill No (Num)"].dropna().astype(int).unique().tolist()
    )
    if not valid_nums:
        return pd.DataFrame(columns=["Missing Bill No", "Status"])

    min_bill = min(valid_nums)
    max_bill = max(valid_nums)
    existing_set = set(valid_nums)

    # Check sequence from 1 (or min_bill) to max_bill
    start_num = 1 if min_bill > 1 else min_bill
    missing = [num for num in range(start_num, max_bill + 1) if num not in existing_set]

    return pd.DataFrame(
        {
            "Missing Bill No": [f"{m:02d}" for m in missing],
            "Status": ["Missing / Not Uploaded" for _ in missing],
        }
    )


def generate_excel_bytes(
    summary_df: pd.DataFrame,
    dispatch_df: pd.DataFrame,
    missing_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Invoice_Summary", index=False)
        dispatch_df.to_excel(writer, sheet_name="Challan_Dispatch_Details", index=False)
        missing_df.to_excel(writer, sheet_name="Missing_Invoices", index=False)

        wb = writer.book
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        alert_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws.freeze_panes = "A2"

            # Style header row
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Auto-fit columns
            for col_idx, col_cells in enumerate(ws.columns, start=1):
                max_len = 0
                for cell in col_cells:
                    val = str(cell.value) if cell.value is not None else ""
                    if len(val) > max_len:
                        max_len = len(val)
                    if sheet_name == "Missing_Invoices" and cell.row > 1:
                        cell.fill = alert_fill
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 42)

    return output.getvalue()


# ==========================================
# SESSION STATE INITIALIZATION
# ==========================================
if "summary_df" not in st.session_state:
    st.session_state.summary_df = pd.DataFrame()
if "dispatch_df" not in st.session_state:
    st.session_state.dispatch_df = pd.DataFrame()

# ==========================================
# SIDEBAR: API KEY & EXISTING MASTER EXCEL
# ==========================================
with st.sidebar:
    st.header("⚙️ Settings")

    # Auto-load from Streamlit Secrets if deployed on Streamlit Cloud
    default_api_key = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""
    api_key = st.text_input(
        "Gemini API Key",
        value=default_api_key,
        type="password",
        help="Streamlit Cloud Secrets me GEMINI_API_KEY save kar sakte hain.",
    )

    st.divider()
    st.subheader("1️⃣ Purana Master Excel Load Karein (Optional)")
    st.caption("Agar aapke paas pehle ka extracted Excel hai, to yahan upload karein taaki naye bills usi me add ho jayein.")
    master_excel = st.file_uploader(
        "Upload Existing Master Excel (.xlsx)",
        type=["xlsx"],
        key="master_excel_uploader",
    )

    if master_excel is not None:
        if st.button("📥 Load Master Excel Data", use_container_width=True):
            try:
                xls = pd.ExcelFile(master_excel)
                if "Invoice_Summary" in xls.sheet_names and "Challan_Dispatch_Details" in xls.sheet_names:
                    st.session_state.summary_df = pd.read_excel(xls, sheet_name="Invoice_Summary")
                    st.session_state.dispatch_df = pd.read_excel(xls, sheet_name="Challan_Dispatch_Details")
                    st.success(f"✅ {len(st.session_state.summary_df)} purane invoices load ho gaye!")
                else:
                    st.error("Invalid Master Excel format. 'Invoice_Summary' aur 'Challan_Dispatch_Details' sheets nahi mili.")
            except Exception as e:
                st.error(f"Excel load karne me error aaya: {e}")

    st.divider()
    overwrite_duplicates = st.checkbox(
        "Duplicate Bill No. aane par overwrite karein",
        value=False,
        help="Unchecked rakhne par pehle se extracted Bill No. skip ho jayega.",
    )

    if st.button("🗑️ Clear Current Session Data", use_container_width=True):
        st.session_state.summary_df = pd.DataFrame()
        st.session_state.dispatch_df = pd.DataFrame()
        st.rerun()

# ==========================================
# MAIN AREA: UPLOAD NEW INVOICE PDFs
# ==========================================
st.subheader("2️⃣ Naye Invoice PDFs Upload Karein (Single ya Multiple)")
uploaded_pdfs = st.file_uploader(
    "Invoice PDF files select karein",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_pdfs:
    if st.button("🚀 Extract & Update Excel Data", type="primary"):
        if not api_key:
            st.error("Kripya Sidebar me apna Gemini API Key daalein (ya Streamlit Secrets me set karein).")
        else:
            existing_bills = set()
            if not st.session_state.summary_df.empty and "Bill No" in st.session_state.summary_df.columns:
                existing_bills = set(st.session_state.summary_df["Bill No"].astype(str).str.strip())

            new_summaries = []
            new_dispatches = []
            skipped_files = []

            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, pdf_file in enumerate(uploaded_pdfs):
                status_text.text(f"Processing ({i + 1}/{len(uploaded_pdfs)}): {pdf_file.name}...")
                try:
                    pdf_bytes = pdf_file.read()
                    inv_data = parse_invoice_with_gemini(pdf_bytes, api_key)
                    bill_no_str = str(inv_data.bill_no).strip()

                    if bill_no_str in existing_bills and not overwrite_duplicates:
                        skipped_files.append(f"{pdf_file.name} (Bill No: {bill_no_str})")
                    else:
                        # Remove old entry if overwriting
                        if bill_no_str in existing_bills and overwrite_duplicates:
                            st.session_state.summary_df = st.session_state.summary_df[
                                st.session_state.summary_df["Bill No"].astype(str).str.strip() != bill_no_str
                            ]
                            st.session_state.dispatch_df = st.session_state.dispatch_df[
                                st.session_state.dispatch_df["Bill No"].astype(str).str.strip() != bill_no_str
                            ]

                        summary_row, dispatch_rows = process_extracted_invoice(inv_data, pdf_file.name)
                        new_summaries.append(summary_row)
                        new_dispatches.extend(dispatch_rows)
                        existing_bills.add(bill_no_str)

                except Exception as e:
                    st.error(f"❌ Error in {pdf_file.name}: {e}")

                progress_bar.progress((i + 1) / len(uploaded_pdfs))

            status_text.empty()
            progress_bar.empty()

            # Append new data to session state
            if new_summaries:
                new_sum_df = pd.DataFrame(new_summaries)
                new_disp_df = pd.DataFrame(new_dispatches)

                st.session_state.summary_df = pd.concat(
                    [st.session_state.summary_df, new_sum_df], ignore_index=True
                )
                st.session_state.dispatch_df = pd.concat(
                    [st.session_state.dispatch_df, new_disp_df], ignore_index=True
                )

                # Sort cleanly by numeric Bill No
                if "Bill No (Num)" in st.session_state.summary_df.columns:
                    st.session_state.summary_df = st.session_state.summary_df.sort_values(
                        by="Bill No (Num)", ascending=True
                    ).reset_index(drop=True)
                if "Bill No (Num)" in st.session_state.dispatch_df.columns:
                    st.session_state.dispatch_df = st.session_state.dispatch_df.sort_values(
                        by=["Bill No (Num)"], ascending=True
                    ).reset_index(drop=True)

                st.success(f"✅ {len(new_summaries)} naye invoice successfully extract aur add ho gaye!")

            if skipped_files:
                st.info("⏭️ Ye duplicate invoices skip kiye gaye: " + ", ".join(skipped_files))

# ==========================================
# DASHBOARD, MISSING INVOICE ALERT & EXPORT
# ==========================================
if not st.session_state.summary_df.empty:
    st.divider()

    # Check Missing Invoices
    missing_df = find_missing_invoices(st.session_state.summary_df)
    if not missing_df.empty:
        missing_list = ", ".join(missing_df["Missing Bill No"].tolist())
        st.warning(f"⚠️ **Missing Invoice Notification:** Ye Bill Numbers sequence me missing hain: **{missing_list}**")
    else:
        st.success("✅ Koi bhi invoice number sequence me missing nahi hai!")

    # Top KPI Metrics
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Invoices", len(st.session_state.summary_df))
    col2.metric("Total Trips / Entries", len(st.session_state.dispatch_df))
    col3.metric("Total TON", f"{st.session_state.summary_df['Total TON'].sum():,.2f}")
    col4.metric("Grand Total (₹)", f"₹ {st.session_state.summary_df['Grand Total'].sum():,.2f}")

    # Tabs for preview
    tab1, tab2, tab3 = st.tabs(["📋 Invoice Summary", "🚚 Challan / Dispatch Details", "⚠️ Missing Invoices"])

    with tab1:
        st.dataframe(st.session_state.summary_df, use_container_width=True, hide_index=True)

    with tab2:
        st.dataframe(st.session_state.dispatch_df, use_container_width=True, hide_index=True)

    with tab3:
        if missing_df.empty:
            st.write("No missing invoices found.")
        else:
            st.dataframe(missing_df, use_container_width=True, hide_index=True)

    # Download Master Excel Button
    excel_bytes = generate_excel_bytes(
        st.session_state.summary_df,
        st.session_state.dispatch_df,
        missing_df,
    )
    st.download_button(
        label="⬇️ Download Updated Master Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"Invoice_Master_Data_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
