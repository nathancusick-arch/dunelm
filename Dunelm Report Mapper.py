import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.cell.cell import MergedCell
import io
from datetime import datetime

st.title("Dunelm Report Generator")

st.write("""
Upload:
1. Latest audits export
2. Last week's LIVE report

The tool will automatically generate:
- This week's LIVE report
- This week's completed report
""")

csv_file = st.file_uploader("Upload audits_basic_data_export.csv", type=["csv"])
live_file = st.file_uploader("Upload last week's LIVE report", type=["xlsx"])

# ============================================================
# COLUMN MAPPING
# ============================================================

COLUMN_MAP = {
    "Order": "order_internal_id",
    "Client": "client_name",
    "Visit": "internal_id",
    "Site": "site_internal_id",
    "Order Deadline": "responsibility",
    "Responsibility": "site_name",
    "Premises Name": "site_address_1",
    "Address1": "site_address_2",
    "Address2": "site_address_3",
    "Address3": None,
    "City": None,
    "Post Code": "site_post_code",
    "Submitted Date": "submitted_date",
    "Approved Date": "approval_date",
    "Item to order": "item_to_order",
    "Actual Visit Date": "date_of_visit",
    "Actual Visit Time": "time_of_visit",
    "AM / PM": None,
    "Pass-Fail": "primary_result",
    "Pass-Fail2": "secondary_result",
    "Abort Reason": "Please detail why you were unable to conduct this audit:",
    "Extra Site 1": "site_code",
    "Extra Site 2": None,
    "Extra Site 3": None,
    "Extra Site 4": None,
}

OUTPUT_COLUMNS = list(COLUMN_MAP.keys())

def map_data(df):
    def map_value(row, mapping):
        if mapping is None:
            return ""
        return str(row.get(mapping, "")).strip()

    final_df = pd.DataFrame(columns=OUTPUT_COLUMNS)

    for col in OUTPUT_COLUMNS:
        final_df[col] = df.apply(lambda r: map_value(r, COLUMN_MAP[col]), axis=1)

    av = final_df[~final_df["Item to order"].isin(["Allergens", "Restaurant", "On the Go"])].copy()
    narv = final_df[final_df["Item to order"].isin(["Allergens", "Restaurant", "On the Go"])].copy()

    return av, narv


# ============================================================
# MAIN PROCESS (AUTO RUN)
# ============================================================

if csv_file and live_file:

    st.info("Processing... please wait")

    # Load export
    df = pd.read_csv(csv_file, dtype=str).fillna("")
    av_df, narv_df = map_data(df)

    # Load workbook
    wb = load_workbook(live_file)

    weekly_ws = wb["Weekly"]
    narv_weekly_ws = wb["NARV Weekly"]
    ytd_ws = wb["YTD"]
    narv_ytd_ws = wb["NARV YTD"]

    # ============================================================
    # GET EXISTING VISITS
    # ============================================================

    def get_existing_visits(ws):
        visits = set()
        for row in ws.iter_rows(min_row=4, max_col=3):
            val = row[2].value
            if val:
                visits.add(str(val))
        return visits

    existing_visits = get_existing_visits(weekly_ws) | get_existing_visits(narv_weekly_ws)

    # Deduplicate
    av_df = av_df[~av_df["Visit"].isin(existing_visits)]
    narv_df = narv_df[~narv_df["Visit"].isin(existing_visits)]

    # ============================================================
    # YEAR CHECK
    # ============================================================

    def get_year_from_ws(ws):
        for row in ws.iter_rows(min_row=4, max_col=16):
            val = row[15].value
            if val:
                return pd.to_datetime(val).year
        return None

    new_year = None
    if not av_df.empty:
        new_year = pd.to_datetime(av_df["Actual Visit Date"]).min().year

    existing_year = get_year_from_ws(ytd_ws)

    reset_ytd = new_year and existing_year and new_year > existing_year

    # ============================================================
    # WRITE DATA
    # ============================================================

    def clear_data(ws):
        for row in ws.iter_rows(min_row=4, max_col=25):
            for cell in row:
                if not isinstance(cell, MergedCell):
                    cell.value = None

    def write_df(ws, df):
        clear_data(ws)
        for r_idx, row in enumerate(df.values, start=4):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=val)

    write_df(weekly_ws, av_df)
    write_df(narv_weekly_ws, narv_df)

    # ============================================================
    # YTD HANDLING
    # ============================================================

    def append_df(ws, df, reset):
        if reset:
            clear_data(ws)
            start_row = 4
        else:
            start_row = ws.max_row + 1

        for r_idx, row in enumerate(df.values, start=start_row):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=val)

    append_df(ytd_ws, av_df, reset_ytd)
    append_df(narv_ytd_ws, narv_df, reset_ytd)

    # ============================================================
    # EXTEND FORMULAS (AA:AB) - TRUE EXCEL LOGIC
    # ============================================================

    def extend_formulas(ws):
        max_row = ws.max_row

        for col in [27, 28]:  # AA, AB
            col_letter = ws.cell(row=4, column=col).column_letter
            base_cell = f"{col_letter}4"
            base_formula = ws[base_cell].value

            if not base_formula:
                continue

            for r in range(5, max_row + 1):
                target_cell = f"{col_letter}{r}"
                translated = Translator(base_formula, origin=base_cell).translate_formula(target_cell)
                ws[target_cell] = translated

    for ws in [weekly_ws, ytd_ws, narv_weekly_ws, narv_ytd_ws]:
        extend_formulas(ws)

    # ============================================================
    # SAVE LIVE REPORT
    # ============================================================

    today = datetime.today().strftime("%d.%m.%Y")

    live_buffer = io.BytesIO()
    wb.save(live_buffer)
    live_buffer.seek(0)

    # ============================================================
    # CREATE COMPLETED REPORT
    # ============================================================

    wb_completed = load_workbook(live_buffer)

    keep_tabs = [
        "Weekly Summary Table",
        "Performance by Area",
        "Allergens Weekly Summary Table",
        "Allergens Performance by Area",
    ]

    for sheet in wb_completed.sheetnames:
        if sheet not in keep_tabs:
            del wb_completed[sheet]

    # Convert to values (safe for merged cells)
    for ws in wb_completed.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell, MergedCell):
                    cell.value = cell.value

    completed_buffer = io.BytesIO()
    wb_completed.save(completed_buffer)
    completed_buffer.seek(0)

    # ============================================================
    # DOWNLOAD OUTPUTS
    # ============================================================

    st.success("Reports generated successfully!")

    st.download_button(
        "Download LIVE Report",
        live_buffer,
        file_name=f"Weekly Report format - {today} LIVE.xlsx"
    )

    st.download_button(
        "Download Completed Report",
        completed_buffer,
        file_name=f"Weekly Report format - {today}.xlsx"
    )
