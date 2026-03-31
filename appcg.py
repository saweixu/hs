import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="Athina Logistics Tool",
    page_icon="logo.png",
    layout="wide"
)

st.sidebar.image("logo.png", width=200)
st.sidebar.markdown("### Athina Logistics")
st.sidebar.caption("Global Access")

st.set_page_config(page_title="HS Code Checker CG", layout="wide")
st.title("HS Code Checker CG")
st.caption(
    "Upload one or more invoice files. "
    "The app checks HS codes and flags:"
    " codes starting with 87, or codes found in the watch list."
)

WATCHLIST = {
    "9405119090",
    "8414591500",
    "9505900000",
    "9503003990",
    "9503009990",
    "9405499090",
    "8714999089",
    "9405219090",
    "4819400000",
    "8305900000",
    "6217100090",
    "4821109000",
    "4202929190",
    "9506919000",
    "8510200000",
    "8504409590",
    "9505109000",
    "9620009900",
    "8544492000",
    "3926909790",
    "8714950000",
}

ALLOWED_EXTENSIONS = [".xlsx", ".xlsm", ".xltx", ".xltm"]


# =========================================================
# HELPERS
# =========================================================
def normalize_hs(value):
    """
    Normalize HS code:
    - convert to string
    - remove spaces, dots, commas, dashes, slashes
    - keep only digits
    """
    if value is None:
        return ""

    s = str(value).strip()
    if not s:
        return ""

    s = s.replace("\n", " ").strip()
    s = re.sub(r"[^\d]", "", s)  # keep digits only
    return s


def get_merged_value(ws, row, col):
    """
    Return the visible value of a merged cell if needed.
    """
    cell = ws.cell(row=row, column=col)

    if not isinstance(cell, MergedCell):
        return cell.value

    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.min_row <= row <= merged_range.max_row
            and merged_range.min_col <= col <= merged_range.max_col
        ):
            return ws.cell(merged_range.min_row, merged_range.min_col).value

    return None


def find_sum_row(ws, search_col=2):
    """
    Find the row where column B contains SUM / SUM: / similar.
    """
    for r in range(1, ws.max_row + 1):
        val = get_merged_value(ws, r, search_col)
        if val is None:
            continue
        txt = str(val).strip().upper().replace(" ", "")
        if txt in {"SUM", "SUM:", "TOTAL", "TOTAL:"} or txt.startswith("SUM"):
            return r
    return None


def analyze_file(uploaded_file):
    """
    Analyze one uploaded Excel file.
    Returns:
      issues: list of dict
      summary: dict
    """
    issues = []

    filename = uploaded_file.name
    ext = Path(filename).suffix.lower()

    if ext not in ALLOWED_EXTENSIONS:
        return [], {
            "file": filename,
            "status": "ERROR",
            "message": f"Unsupported file type: {ext}",
            "checked_rows": 0,
            "issue_count": 0,
        }

    try:
        wb = load_workbook(filename=io.BytesIO(uploaded_file.getvalue()), data_only=True)
    except Exception as e:
        return [], {
            "file": filename,
            "status": "ERROR",
            "message": f"Cannot open file: {e}",
            "checked_rows": 0,
            "issue_count": 0,
        }

    if "INVOICE" not in wb.sheetnames:
        return [], {
            "file": filename,
            "status": "ERROR",
            "message": "Sheet 'INVOICE' not found",
            "checked_rows": 0,
            "issue_count": 0,
        }

    ws = wb["INVOICE"]
    sum_row = find_sum_row(ws, search_col=2)

    if not sum_row:
        return [], {
            "file": filename,
            "status": "ERROR",
            "message": "SUM row not found in column B",
            "checked_rows": 0,
            "issue_count": 0,
        }

    start_row = 20
    end_row = sum_row - 1

    if end_row < start_row:
        return [], {
            "file": filename,
            "status": "ERROR",
            "message": f"Invalid range: C{start_row}:C{end_row}",
            "checked_rows": 0,
            "issue_count": 0,
        }

    checked_rows = 0

    for row in range(start_row, end_row + 1):
        raw_hs = get_merged_value(ws, row, 3)  # column C
        hs = normalize_hs(raw_hs)

        if not hs:
            continue

        checked_rows += 1
        reasons = []

        if hs[:2] == "87":
            reasons.append("HS starts with 87")

        if hs in WATCHLIST:
            reasons.append("HS found in watch list")

        if reasons:
            issues.append(
                {
                    "File": filename,
                    "Sheet": "INVOICE",
                    "Row": row,
                    "Cell": f"C{row}",
                    "HS Code": hs,
                    "Reason": " | ".join(reasons),
                }
            )

    status = "OK" if not issues else "FLAGGED"
    message = "No issue found" if not issues else f"{len(issues)} issue(s) found"

    return issues, {
        "file": filename,
        "status": status,
        "message": message,
        "checked_rows": checked_rows,
        "issue_count": len(issues),
    }


def build_excel_report(summary_df, issues_df):
    """
    Create an Excel report in memory.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        issues_df.to_excel(writer, sheet_name="Issues", index=False)
    output.seek(0)
    return output


# =========================================================
# UI
# =========================================================
st.subheader("Upload invoice files")

uploaded_files = st.file_uploader(
    "Select Excel invoice files",
    type=["xlsx", "xlsm", "xltx", "xltm"],
    accept_multiple_files=True,
)

col1, col2 = st.columns([1, 1])
with col1:
    run_btn = st.button("Run HS check", type="primary")
with col2:
    show_watchlist = st.checkbox("Show watch list", value=False)

if show_watchlist:
    st.markdown("### Watch list")
    watch_df = pd.DataFrame({"HS Code": sorted(WATCHLIST)})
    st.dataframe(watch_df, use_container_width=True, hide_index=True)

if run_btn:
    if not uploaded_files:
        st.warning("Please upload at least one file.")
        st.stop()

    all_issues = []
    all_summaries = []

    progress = st.progress(0)
    status_box = st.empty()

    total_files = len(uploaded_files)

    for idx, f in enumerate(uploaded_files, start=1):
        status_box.info(f"Checking {idx}/{total_files}: {f.name}")
        issues, summary = analyze_file(f)
        all_issues.extend(issues)
        all_summaries.append(summary)
        progress.progress(idx / total_files)

    status_box.success("Check completed.")

    summary_df = pd.DataFrame(all_summaries)
    issues_df = pd.DataFrame(all_issues)

    if summary_df.empty:
        summary_df = pd.DataFrame(columns=["file", "status", "message", "checked_rows", "issue_count"])

    if issues_df.empty:
        issues_df = pd.DataFrame(columns=["File", "Sheet", "Row", "Cell", "HS Code", "Reason"])

    st.markdown("## Summary")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    total_flagged_files = int((summary_df["status"] == "FLAGGED").sum()) if "status" in summary_df.columns else 0
    total_error_files = int((summary_df["status"] == "ERROR").sum()) if "status" in summary_df.columns else 0
    total_issues = len(issues_df)

    c1, c2, c3 = st.columns(3)
    c1.metric("Files checked", len(summary_df))
    c2.metric("Flagged files", total_flagged_files)
    c3.metric("Total issues", total_issues)

    if total_error_files > 0:
        st.error(f"{total_error_files} file(s) could not be checked.")

    st.markdown("## Issues found")
    if issues_df.empty:
        st.success("No HS issue found.")
    else:
        st.dataframe(issues_df, use_container_width=True, hide_index=True)

    report_file = build_excel_report(summary_df, issues_df)

    st.download_button(
        label="Download Excel report",
        data=report_file,
        file_name="hs_code_check_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
