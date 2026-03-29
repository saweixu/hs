import io
import re
import json
from datetime import date
from collections import defaultdict

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from requests import Session
from zeep import Client
from zeep.helpers import serialize_object
from zeep.transports import Transport

# =========================================================
# CONFIG
# =========================================================
WSDL_URL = "https://ec.europa.eu/taxation_customs/dds2/taric/services/goods?wsdl"
DEFAULT_COUNTRY = "CN"
TODAY_ISO = date.today().isoformat()

st.set_page_config(page_title="HS Code Analyzer", layout="wide")
st.title("HS Code Analyzer from Invoices")
st.caption(
    "Upload multiple invoice files, extract HS codes from INVOICE column C, "
    "remove duplicates, analyze them against TARIC, and export OUTPUT.xlsx"
)

# =========================================================
# HELPERS - EXCEL
# =========================================================
def find_sum_row(ws, search_col=2):
    """
    Find row where column B contains SUM.
    """
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, search_col).value
        if v is None:
            continue
        txt = str(v).strip().upper().replace(":", "").replace(" ", "")
        if txt == "SUM":
            return r
    return None


def get_merged_cell_value(ws, row, col):
    """
    Return cell value, including merged-cell master value.
    """
    cell = ws.cell(row=row, column=col)

    if not isinstance(cell, MergedCell):
        return cell.value

    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(merged_range.min_row, merged_range.min_col).value

    return None


def get_best_cell_value(ws_data, ws_formula, row, col):
    """
    Prefer calculated value, fallback to raw formula/text.
    """
    v_data = get_merged_cell_value(ws_data, row, col)
    if v_data not in (None, ""):
        return v_data

    v_formula = get_merged_cell_value(ws_formula, row, col)
    return v_formula


def normalize_hs_code(raw):
    """
    Extract possible HS/TARIC codes from a cell.
    Accepts formats like:
    - 9403700000
    - 9403.70.00
    - 9403 70 0000
    - HS: 9403700000
    """
    if raw is None:
        return []

    text = str(raw).strip()
    candidates = set()

    # direct 6-10 digit blocks
    for m in re.findall(r"(?<!\d)(\d{6,10})(?!\d)", text):
        candidates.add(m)

    # dotted / spaced / slashed groups
    for m in re.findall(r"(?<!\d)(\d{4}[.\s/-]?\d{2}(?:[.\s/-]?\d{2}){0,2})(?!\d)", text):
        digits = re.sub(r"\D", "", m)
        if 6 <= len(digits) <= 10:
            candidates.add(digits)

    return sorted(candidates)


def extract_hs_from_invoice_file(uploaded_file):
    """
    Extract HS codes from INVOICE!C20:C(before SUM).
    SUM is searched in column B.
    """
    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()

    wb_data = load_workbook(io.BytesIO(file_bytes), data_only=True)
    wb_formula = load_workbook(io.BytesIO(file_bytes), data_only=False)

    if "INVOICE" in wb_data.sheetnames:
        ws_data = wb_data["INVOICE"]
        ws_formula = wb_formula["INVOICE"]
        sheet_used = "INVOICE"
    else:
        sheet_used = wb_data.sheetnames[0]
        ws_data = wb_data[sheet_used]
        ws_formula = wb_formula[sheet_used]

    sum_row = find_sum_row(ws_data, search_col=2)
    if not sum_row:
        sum_row = find_sum_row(ws_formula, search_col=2)

    if not sum_row:
        return [], [], f"{uploaded_file.name}: SUM row not found in sheet '{sheet_used}'"

    results = []
    debug_rows = []

    for row in range(20, sum_row):
        # HS code in column C = 3
        cell_value = get_best_cell_value(ws_data, ws_formula, row, 3)

        debug_rows.append({
            "file_name": uploaded_file.name,
            "sheet_name": sheet_used,
            "row": row,
            "cell": f"C{row}",
            "raw_value": "" if cell_value is None else str(cell_value),
        })

        codes = normalize_hs_code(cell_value)
        for code in codes:
            results.append({
                "file_name": uploaded_file.name,
                "sheet_name": sheet_used,
                "row": row,
                "raw_cell_value": "" if cell_value is None else str(cell_value),
                "hs_code": code,
            })

    return results, debug_rows, None


# =========================================================
# HELPERS - TARIC
# =========================================================
def get_taric_client():
    """
    Create TARIC SOAP client safely.
    Raises readable exception instead of crashing the app.
    """
    session = Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept": "text/xml, application/xml, */*"
    })

    # Step 1: direct WSDL availability test
    try:
        response = session.get(WSDL_URL, timeout=30)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        raise Exception(f"Cannot access TARIC WSDL URL: {e}")

    # Step 2: SOAP client creation
    try:
        transport = Transport(session=session, timeout=30)
        client = Client(wsdl=WSDL_URL, transport=transport)
        return client
    except Exception as e:
        raise Exception(f"SOAP client creation failed: {e}")


def safe_serialize(obj):
    try:
        return serialize_object(obj)
    except Exception:
        try:
            return json.loads(json.dumps(obj, default=str))
        except Exception:
            return str(obj)


def flatten_strings(obj, found=None):
    if found is None:
        found = []

    if obj is None:
        return found

    if isinstance(obj, dict):
        for v in obj.values():
            flatten_strings(v, found)
    elif isinstance(obj, list):
        for v in obj:
            flatten_strings(v, found)
    else:
        found.append(str(obj))

    return found


def shorten_text(text, max_len=3000):
    if text is None:
        return ""
    text = str(text).strip()
    if len(text) <= max_len:
        return text
    return text[:max_len] + " ..."


def summarize_measures(serialized_response):
    leaves = flatten_strings(serialized_response, [])
    clean = []
    seen = set()

    keywords = (
        "duty", "third country duty", "erga omnes", "import",
        "measure", "certificate", "licence", "restriction",
        "prohibition", "anti-dumping", "tariff", "suspension",
        "quota", "additional code", "vat", "excise", "%"
    )

    for item in leaves:
        line = " ".join(item.split())
        low = line.lower()
        if len(line) < 2:
            continue
        if any(k in low for k in keywords):
            if line not in seen:
                seen.add(line)
                clean.append(line)

    if not clean:
        fallback = []
        for item in leaves:
            line = " ".join(item.split())
            if len(line) >= 3 and line not in fallback:
                fallback.append(line)
            if len(fallback) >= 15:
                break
        return " | ".join(fallback[:15])

    return " | ".join(clean[:20])


def taric_call_with_fallbacks(client, hs_code, country_code, reference_date):
    """
    Try several tradeMovement variants.
    """
    last_error = None
    variants = ["I", "IMPORT", "1", None]

    for tm in variants:
        try:
            kwargs = {
                "goodsCode": hs_code,
                "countryCode": country_code,
                "referenceDate": reference_date,
            }
            if tm is not None:
                kwargs["tradeMovement"] = tm

            resp = client.service.goodsMeasForWs(**kwargs)
            return resp, tm, None
        except Exception as e:
            last_error = str(e)

    return None, None, last_error


def analyze_hs_code(client, hs_code, country_code=DEFAULT_COUNTRY, reference_date=TODAY_ISO):
    description = ""
    measures_summary = ""
    raw_json = ""
    used_trade_movement = ""
    status = "OK"
    error = ""

    # Description
    try:
        d = client.service.goodsDescrForWs(
            goodsCode=hs_code,
            languageCode="EN",
            referenceDate=reference_date,
        )
        d_ser = safe_serialize(d)
        description = shorten_text(" | ".join(flatten_strings(d_ser, [])), 1000)
    except Exception as e:
        description = ""
        error = f"Description error: {e}"

    # Measures
    resp, used_tm, meas_error = taric_call_with_fallbacks(
        client=client,
        hs_code=hs_code,
        country_code=country_code,
        reference_date=reference_date
    )

    used_trade_movement = "" if used_tm is None else str(used_tm)

    if resp is not None:
        ser = safe_serialize(resp)
        measures_summary = shorten_text(summarize_measures(ser), 3000)
        raw_json = shorten_text(json.dumps(ser, ensure_ascii=False, indent=2, default=str), 12000)
    else:
        status = "ERROR"
        error = f"{error} | Measures error: {meas_error}".strip(" |")
        raw_json = ""

    if not error:
        error = ""

    return {
        "hs_code": hs_code,
        "country_code": country_code,
        "reference_date": reference_date,
        "trade_movement_used": used_trade_movement,
        "description_en": description,
        "measures_summary": measures_summary,
        "status": status,
        "error": error,
        "raw_response": raw_json,
    }


# =========================================================
# OUTPUT EXCEL
# =========================================================
def build_output_excel(df_hs_found, df_summary, df_debug):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_hs_found.to_excel(writer, index=False, sheet_name="HS_FOUND")
        df_summary.to_excel(writer, index=False, sheet_name="OUTPUT")
        df_debug.to_excel(writer, index=False, sheet_name="DEBUG_READ")

    output.seek(0)
    return output


# =========================================================
# UI
# =========================================================
uploaded_files = st.file_uploader(
    "Upload invoice files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"{len(uploaded_files)} file(s) uploaded.")

    all_rows = []
    all_debug_rows = []
    warnings = []

    for f in uploaded_files:
        rows, debug_rows, err = extract_hs_from_invoice_file(f)
        if err:
            warnings.append(err)
        all_rows.extend(rows)
        all_debug_rows.extend(debug_rows)

    if warnings:
        for w in warnings:
            st.warning(w)

    if not all_rows:
        st.error("No HS code found in uploaded files.")
        if all_debug_rows:
            with st.expander("Debug preview", expanded=True):
                st.dataframe(pd.DataFrame(all_debug_rows), use_container_width=True)
        st.stop()

    df_found = pd.DataFrame(all_rows)
    df_debug = pd.DataFrame(all_debug_rows)

    grouped_files = defaultdict(set)
    grouped_positions = defaultdict(list)

    for row in all_rows:
        grouped_files[row["hs_code"]].add(row["file_name"])
        grouped_positions[row["hs_code"]].append(f"{row['file_name']} [row {row['row']}]")

    unique_codes = sorted(df_found["hs_code"].dropna().astype(str).unique().tolist())

    st.success(f"{len(df_found)} HS code occurrence(s) found.")
    st.success(f"{len(unique_codes)} unique HS code(s) after duplicate removal.")

    with st.expander("Preview extracted HS codes", expanded=False):
        st.dataframe(df_found, use_container_width=True)

    with st.expander("Debug preview (read values from column C)", expanded=False):
        st.dataframe(df_debug, use_container_width=True)

    if st.button("Analyze HS Codes"):
        summary_rows = []

        with st.spinner("Connecting to TARIC and analyzing HS codes..."):
            try:
                client = get_taric_client()
                taric_available = True
                st.success("TARIC connection OK.")
            except Exception as e:
                taric_available = False
                st.warning(f"TARIC not available. Output will still be generated. Error: {e}")

            progress = st.progress(0)
            total = len(unique_codes)

            for i, hs in enumerate(unique_codes, start=1):
                if taric_available:
                    try:
                        result = analyze_hs_code(
                            client=client,
                            hs_code=hs,
                            country_code=DEFAULT_COUNTRY,
                            reference_date=TODAY_ISO
                        )
                    except Exception as e:
                        result = {
                            "hs_code": hs,
                            "country_code": DEFAULT_COUNTRY,
                            "reference_date": TODAY_ISO,
                            "trade_movement_used": "",
                            "description_en": "",
                            "measures_summary": "",
                            "status": "ERROR",
                            "error": str(e),
                            "raw_response": "",
                        }
                else:
                    result = {
                        "hs_code": hs,
                        "country_code": DEFAULT_COUNTRY,
                        "reference_date": TODAY_ISO,
                        "trade_movement_used": "",
                        "description_en": "",
                        "measures_summary": "",
                        "status": "TARIC_UNAVAILABLE",
                        "error": "TARIC connection unavailable in current environment.",
                        "raw_response": "",
                    }

                result["source_file_count"] = len(grouped_files[hs])
                result["source_files"] = " | ".join(sorted(grouped_files[hs]))
                result["source_positions"] = " | ".join(grouped_positions[hs])

                summary_rows.append(result)
                progress.progress(i / total)

        df_summary = pd.DataFrame(summary_rows)

        st.subheader("Analysis Result")
        st.dataframe(df_summary, use_container_width=True)

        xlsx_data = build_output_excel(df_found, df_summary, df_debug)

        st.download_button(
            label="Download OUTPUT.xlsx",
            data=xlsx_data,
            file_name="OUTPUT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Upload one or more invoice files to begin.")
