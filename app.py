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
VEH_PAT = r"[A-Z]{2}\s*-?\s*\d{1,2}\s*-?\s*[A-Z]{1,3}\s*-?\s*\d{1,4}"


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

    # 1. Bill No & Invoice Date
    bill_no = ""
    invoice_date = ""
    for line in p1_lines:
        if not bill_no:
            m_bill = re.search(r"Bill\s*No\.?\s*[:\-]?\s*([A-Za-z0-9/\-]+)", line, re.IGNORECASE)
            if m_bill:
                bill_no = m_bill.group(1).strip()
        if not invoice_date:
            m_date = re.search(r"Date\s*[:\-]?\s*(" + DATE_PAT + r")", line, re.IGNORECASE)
            if m_date:
                invoice_date = parse_date_ddmmyyyy(m_date.group(1))

    # Fallback Bill No from filename if not found in text
    if not bill_no:
        m_fn = re.match(r"^\s*(\d+)", source_filename)
        if m_fn:
            bill_no = m_fn.group(1)

    # 2. Bill To & Shipping Address (column-based split)
    bill_to_idx = None
    gst_idx = None
    ship_col = None
    for idx, line in enumerate(p1_lines):
        if bill_to_idx is None and re.search(r"Bill\s*To", line, re.IGNORECASE):
            bill_to_idx = idx
        if gst_idx is None and re.search(r"GST\s*NO", line, re.IGNORECASE):
            gst_idx = idx
            break

    bill_to_lines = []
    shipping_lines = []
    if bill_to_idx is not None and gst_idx is not None:
        for line in p1_lines[bill_to_idx:gst_idx]:
            m_ship = re.search(r"Shipping\s*Ad+ress?\s*[:\-]?", line, re.IGNORECASE)
            if m_ship:
                ship_col = m_ship.start()
                after_label = line[m_ship.end() :].strip()
                if after_label:
                    shipping_lines.append(after_label)
                break

        for line in p1_lines[bill_to_idx + 1 : gst_idx]:
            if not line.strip():
                continue
            if ship_col is not None and len(line) > ship_col:
                left_part = line[:ship_col].strip()
                right_part = line[ship_col:].strip()
                right_part = re.sub(
                    r"^Shipping\s*Ad+ress?\s*[:\-]?\s*", "", right_part, flags=re.IGNORECASE
                ).strip()
                if left_part:
                    bill_to_lines.append(left_part)
                if right_part:
                    shipping_lines.append(right_part)
            else:
                left_part = re.sub(
                    r"Shipping\s*Ad+ress?\s*[:\-]?.*$", "", line, flags=re.IGNORECASE
                ).strip()
                if left_part:
                    bill_to_lines.append(left_part)

    bill_to_name = bill_to_lines[0] if bill_to_lines else ""
    bill_to_address = ", ".join(bill_to_lines[1:]) if len(bill_to_lines) > 1 else ""
    shipping_address = ", ".join(shipping_lines)

    # 3. Customer GST/PAN & State
    customer_gst_pan = ""
    customer_state = ""
    if gst_idx is not None:
        gst_line = p1_lines[gst_idx]
        parts = re.split(r"\s{4,}", gst_line.strip())
        if parts:
            left_gst = re.sub(
                r"^GST\s*NO\.?\s*/?\s*PAN\s*NO\.?\s*[:\-]*\s*", "", parts[0], flags=re.IGNORECASE
            ).strip()
            left_gst = re.sub(
                r"^(?:GST|PAN)\s*No\.?\s*[:\-]*\s*", "", left_gst, flags=re.IGNORECASE
            ).strip()
            customer_gst_pan = left_gst
        if len(parts) >= 2:
            customer_state = normalize_state(parts[-1])

    # Fallback: GSTIN starting with 27 belongs to Maharashtra
    if not customer_state and customer_gst_pan.startswith("27"):
        customer_state = "Maharashtra"

    # 4. HSN Code & Company Name (auto-read from Bank Details if available)
    hsn_code = ""
    detected_company_name = ""
    for line in p1_lines:
        if not hsn_code:
            m_hsn = re.search(r"HSN\s*CODE\.?\s*[:\-]?\s*([0-9]+)", line, re.IGNORECASE)
            if m_hsn:
                hsn_code = m_hsn.group(1).strip()
        if not detected_company_name:
            m_comp = re.search(r"^Name\s*[:\-]\s*([A-Za-z0-9\s&\.\-]+?)(?:\s{3,}|$)", line.strip())
            if m_comp:
                detected_company_name = m_comp.group(1).strip()

    company_name = detected_company_name or default_company_info["company_name"]

    # 5. Extract Page 1 Table Rows (between Product Descriptions and HSN CODE / Gross Total)
    tbl_start = None
    tbl_end = None
    for idx, line in enumerate(p1_lines):
        if tbl_start is None and re.search(r"Product\s*Descriptions?", line, re.IGNORECASE):
            tbl_start = idx + 1
        elif tbl_start is not None and (
            re.search(r"HSN\s*CODE", line, re.IGNORECASE)
            or re.search(r"Gross\s*Total", line, re.IGNORECASE)
            or re.search(r"Amount\s*Chag?eable", line, re.IGNORECASE)
        ):
            tbl_end = idx
            break

    summary_rates: Dict[str, float] = {}
    p1_dispatch_rows: List[dict] = []
    p1_summary_rows: List[dict] = []

    if tbl_start is not None:
        table_lines = p1_lines[tbl_start : (tbl_end if tbl_end else len(p1_lines))]
        for line in table_lines:
            s = line.strip()
            if not s:
                continue

            # Layout B1: Date + Vehicle + Metal + TON + Rate + Amount
            m_full = re.match(
                rf"^(?:\d+\s{{2,}})?({DATE_PAT})\s+({VEH_PAT})\s+(.+?)\s+({NUM_PAT})\s+({NUM_PAT})\s+({NUM_PAT})$",
                s,
                re.IGNORECASE,
            )
            if m_full:
                d_date, veh, metal, ton, rate, _ = m_full.groups()
                metal_norm = normalize_metal(metal)
                summary_rates[metal_norm] = to_float(rate)
                p1_dispatch_rows.append(
                    {
                        "sr_no": str(len(p1_dispatch_rows) + 1),
                        "dispatch_date": parse_date_ddmmyyyy(d_date),
                        "challan_no": "",
                        "vehicle_no": re.sub(r"\s+", " ", veh.strip().upper()),
                        "metal_type": metal_norm,
                        "ton": to_float(ton),
                        "rate": to_float(rate),
                    }
                )
                continue

            # Layout B2: Vehicle + Date + Metal + TON + Rate + Amount (if reversed)
            m_full_rev = re.match(
                rf"^(?:\d+\s{{2,}})?({VEH_PAT})\s+({DATE_PAT})\s+(.+?)\s+({NUM_PAT})\s+({NUM_PAT})\s+({NUM_PAT})$",
                s,
                re.IGNORECASE,
            )
            if m_full_rev:
                veh, d_date, metal, ton, rate, _ = m_full_rev.groups()
                metal_norm = normalize_metal(metal)
                summary_rates[metal_norm] = to_float(rate)
                p1_dispatch_rows.append(
                    {
                        "sr_no": str(len(p1_dispatch_rows) + 1),
                        "dispatch_date": parse_date_ddmmyyyy(d_date),
                        "challan_no": "",
                        "vehicle_no": re.sub(r"\s+", " ", veh.strip().upper()),
                        "metal_type": metal_norm,
                        "ton": to_float(ton),
                        "rate": to_float(rate),
                    }
                )
                continue

            # Layout A: Summary line -> Metal + TON + Rate + Amount
            m_sum = re.match(
                rf"^(.+?)\s+({NUM_PAT})\s+({NUM_PAT})\s+({NUM_PAT})$",
                s,
                re.IGNORECASE,
            )
            if m_sum:
                metal, ton, rate, _ = m_sum.groups()
                metal = re.sub(r"^\d+\s{2,}", "", metal).strip()
                metal_norm = normalize_metal(metal)
                summary_rates[metal_norm] = to_float(rate)
                p1_summary_rows.append(
                    {
                        "sr_no": str(len(p1_summary_rows) + 1),
                        "dispatch_date": invoice_date,
                        "challan_no": "",
                        "vehicle_no": "",
                        "metal_type": metal_norm,
                        "ton": to_float(ton),
                        "rate": to_float(rate),
                    }
                )

    # 6. Extract Page 2+ Challan / Dispatch Table (if present)
    p2_dispatch_rows: List[dict] = []
    for page in reader.pages[1:]:
        p_text = page.extract_text(extraction_mode="layout") or ""
        for line in p_text.splitlines():
            s = line.strip()
            if not s or re.search(r"Sr\.?\s*No|VEHIC", s, re.IGNORECASE):
                continue

            # Pattern: Sr.No. Date [ChallanNo] VehicleNo Metal TON
            m_p2 = re.match(
                rf"^(\d+)\s+({DATE_PAT})\s+(?:(\S+)\s+)?({VEH_PAT})\s+(.+?)\s+({NUM_PAT})$",
                s,
                re.IGNORECASE,
            )
            if m_p2:
                sr, d_date, challan, veh, metal, ton = m_p2.groups()
                metal_norm = normalize_metal(metal)
                p2_dispatch_rows.append(
                    {
                        "sr_no": sr.strip(),
                        "dispatch_date": parse_date_ddmmyyyy(d_date),
                        "challan_no": (challan or "").strip(),
                        "vehicle_no": re.sub(r"\s+", " ", veh.strip().upper()),
                        "metal_type": metal_norm,
                        "ton": to_float(ton),
                        "rate": summary_rates.get(metal_norm, 0.0),
                    }
                )

    # Choose best source of line items: Page 2 Challan > Page 1 Dispatch > Page 1 Summary
    raw_items = p2_dispatch_rows or p1_dispatch_rows or p1_summary_rows

    # 7. Calculate GST & Build Final Rows
    is_maharashtra = customer_state.lower() == "maharashtra"
    fy = get_financial_year(invoice_date)
    bill_num_int = extract_bill_number_int(bill_no)
    bill_no_clean = str(bill_no).strip()

    dispatch_rows = []
    total_ton = 0.0
    gross_amount = 0.0

    for item in raw_items:
        ton = float(item["ton"])
        rate = float(item["rate"])
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
                "Invoice Date": invoice_date,
                "Bill To Name": bill_to_name,
                "Customer State": customer_state,
                "Sr. No.": item["sr_no"],
                "Dispatch Date": item["dispatch_date"],
                "Challan No": item["challan_no"],
                "Vehicle No": item["vehicle_no"],
                "Metal Type": item["metal_type"],
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
        "Invoice Date": invoice_date,
        "Company Name": company_name,
        "Company GSTIN": default_company_info["company_gstin"],
        "Company PAN": default_company_info["company_pan"],
        "Company Details": default_company_info["company_details"],
        "Bill To Name": bill_to_name,
        "Bill To Address": bill_to_address,
        "Shipping Address": shipping_address,
        "Customer GST/PAN": customer_gst_pan,
        "Customer State": customer_state,
        "HSN Code": hsn_code,
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


def extract_pdfs_from_uploads(uploaded_files):
    pdf_list = []
    for f in uploaded_files:
        name = f.name
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(f.read())) as zf:
                for zname in zf.namelist():
                    if zname.lower().endswith(".pdf") and not zname.startswith("__MACOSX"):
                        clean_name = zname.split("/")[-1]
                        pdf_list.append((clean_name, zf.read(zname)))
        elif name.lower().endswith(".pdf"):
            pdf_list.append((name, f.read()))
    return pdf_list


def find_missing_invoices(summary_df: pd.DataFrame) -> pd.DataFrame:
    if summary_df.empty or "Bill No (Num)" not in summary_df.columns:
        return pd.DataFrame(columns=["Missing Bill No", "Status"])

    valid_nums = summary_df["Bill No (Num)"].dropna().astype(int).unique().tolist()
    if not valid_nums:
        return pd.DataFrame(columns=["Missing Bill No", "Status"])

    min_bill = min(valid_nums)
    max_bill = max(valid_nums)
    existing_set = set(valid_nums)

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

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

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
# SIDEBAR: COMPANY HEADER & MASTER EXCEL
# ==========================================
with st.sidebar:
    st.header("🏢 Company Header Details")
    st.caption("PDF ka top letterhead image hai, isliye ye details yahan se automatically fill hongi (zaroorat par edit kar sakte hain):")
    comp_name = st.text_input("Company Name", value="SAIKRUPA ENTERPRISES")
    comp_gstin = st.text_input("Company GSTIN", value="27AZRPP3943M1ZH")
    comp_pan = st.text_input("Company PAN", value="AZRPP3943M")
    comp_details = st.text_area(
        "Company Address & Contact",
        value="Water & Building Material Suppliers, At. Khandwal, Post. Amane, Tal. Bhiwandi, Dist. Thane. | Dinesh Patil: 9049855858 / 9881210811",
        height=90,
    )

    default_company_info = {
        "company_name": comp_name.strip(),
        "company_gstin": comp_gstin.strip(),
        "company_pan": comp_pan.strip(),
        "company_details": comp_details.strip(),
    }

    st.divider()
    st.subheader("1️⃣ Purana Master Excel Load Karein (Optional)")
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
                    st.error("Invalid Master Excel format.")
            except Exception as e:
                st.error(f"Excel load karne me error aaya: {e}")

    st.divider()
    overwrite_duplicates = st.checkbox(
        "Duplicate Bill No. aane par overwrite karein",
        value=False,
    )

    if st.button("🗑️ Clear Current Session Data", use_container_width=True):
        st.session_state.summary_df = pd.DataFrame()
        st.session_state.dispatch_df = pd.DataFrame()
        st.rerun()

# ==========================================
# MAIN AREA: UPLOAD PDFs OR ZIP
# ==========================================
st.subheader("2️⃣ Naye Invoice PDFs ya ZIP File Upload Karein")
uploaded_files = st.file_uploader(
    "Saari PDFs select karein ya unki 1 ZIP file upload karein (Fastest)",
    type=["pdf", "zip"],
    accept_multiple_files=True,
)

if uploaded_files:
    if st.button("⚡ Extract & Update Excel Data", type="primary"):
        pdf_items = extract_pdfs_from_uploads(uploaded_files)
        total_pdfs = len(pdf_items)

        if total_pdfs == 0:
            st.warning("Koi valid PDF file nahi mili.")
        else:
            existing_bills = set()
            if not st.session_state.summary_df.empty and "Bill No" in st.session_state.summary_df.columns:
                existing_bills = set(st.session_state.summary_df["Bill No"].astype(str).str.strip())

            new_summaries = []
            new_dispatches = []
            skipped_files = []
            error_files = []

            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, (fname, fbytes) in enumerate(pdf_items, start=1):
                try:
                    summary_row, dispatch_rows = parse_pdf_invoice(fbytes, fname, default_company_info)
                    bill_no_str = str(summary_row["Bill No"]).strip()

                    if bill_no_str in existing_bills and not overwrite_duplicates:
                        skipped_files.append(f"{fname} (Bill No: {bill_no_str})")
                    else:
                        if bill_no_str in existing_bills and overwrite_duplicates:
                            st.session_state.summary_df = st.session_state.summary_df[
                                st.session_state.summary_df["Bill No"].astype(str).str.strip() != bill_no_str
                            ]
                            st.session_state.dispatch_df = st.session_state.dispatch_df[
                                st.session_state.dispatch_df["Bill No"].astype(str).str.strip() != bill_no_str
                            ]

                        new_summaries.append(summary_row)
                        new_dispatches.extend(dispatch_rows)
                        existing_bills.add(bill_no_str)

                except Exception as e:
                    error_files.append(f"{fname}: {e}")

                progress_bar.progress(i / total_pdfs)
                status_text.text(f"Processed {i}/{total_pdfs}: {fname}")

            status_text.empty()
            progress_bar.empty()

            if new_summaries:
                new_sum_df = pd.DataFrame(new_summaries)
                new_disp_df = pd.DataFrame(new_dispatches)

                st.session_state.summary_df = pd.concat(
                    [st.session_state.summary_df, new_sum_df], ignore_index=True
                )
                st.session_state.dispatch_df = pd.concat(
                    [st.session_state.dispatch_df, new_disp_df], ignore_index=True
                )

                if "Bill No (Num)" in st.session_state.summary_df.columns:
                    st.session_state.summary_df = st.session_state.summary_df.sort_values(
                        by="Bill No (Num)", ascending=True
                    ).reset_index(drop=True)
                if "Bill No (Num)" in st.session_state.dispatch_df.columns:
                    st.session_state.dispatch_df = st.session_state.dispatch_df.sort_values(
                        by=["Bill No (Num)"], ascending=True
                    ).reset_index(drop=True)

                st.success(f"✅ {len(new_summaries)} naye invoice turant extract aur add ho gaye!")

            if skipped_files:
                st.info("⏭️ Ye duplicate invoices skip kiye gaye: " + ", ".join(skipped_files))

            if error_files:
                with st.expander(f"⚠️ {len(error_files)} files me issue aaya (click karke dekhein)"):
                    for err in error_files:
                        st.write(err)

# ==========================================
# DASHBOARD, MISSING INVOICE ALERT & EXPORT
# ==========================================
if not st.session_state.summary_df.empty:
    st.divider()

    missing_df = find_missing_invoices(st.session_state.summary_df)
    if not missing_df.empty:
        missing_list = ", ".join(missing_df["Missing Bill No"].tolist())
        st.warning(f"⚠️ **Missing Invoice Notification:** Ye Bill Numbers sequence me missing hain: **{missing_list}**")
    else:
        st.success("✅ Koi bhi invoice number sequence me missing nahi hai!")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Invoices", len(st.session_state.summary_df))
    col2.metric("Total Trips / Entries", len(st.session_state.dispatch_df))
    col3.metric("Total TON", f"{st.session_state.summary_df['Total TON'].sum():,.2f}")
    col4.metric("Grand Total (₹)", f"₹ {st.session_state.summary_df['Grand Total'].sum():,.2f}")

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
