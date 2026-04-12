import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.cell.cell import MergedCell
import io
from datetime import datetime

st.title("Dunelm Report Generator")

st.write("""
1. Upload the latest audits export file  
2. Upload last week's LIVE report  
3. Download this week's LIVE and completed reports  
""")

csv_file = st.file_uploader("Upload audits_basic_data_export.csv", type=["csv"])
live_file = st.file_uploader("Upload last week's LIVE report", type=["xlsx"])


# ============================================================
# COLUMN MAP
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

    for col in ["Actual Visit Date", "Submitted Date", "Approved Date"]:
        final_df[col] = pd.to_datetime(final_df[col], errors="coerce", dayfirst=True)

    final_df["Actual Visit Time"] = pd.to_datetime(
        final_df["Actual Visit Time"], errors="coerce"
    ).dt.time

    final_df["Extra Site 1"] = pd.to_numeric(final_df["Extra Site 1"], errors="coerce")

    av = final_df[~final_df["Item to order"].isin(["Allergens", "Restaurant", "On the Go"])].copy()
    narv = final_df[final_df["Item to order"].isin(["Allergens", "Restaurant", "On the Go"])].copy()

    return av, narv


# ============================================================
# HELPERS
# ============================================================

def get_last_data_row(ws):
    for r in range(ws.max_row, 3, -1):
        if ws.cell(r, 1).value:
            return r
    return 4


def write(ws, df):
    for row in ws.iter_rows(min_row=4, max_col=25):
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None

    for r, row in enumerate(df.itertuples(index=False), start=4):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(r, c, val)

            if c in [12, 13, 16] and pd.notna(val):
                cell.number_format = "DD/MM/YYYY"

            if c == 17 and pd.notna(val):
                cell.number_format = "HH:MM"


def fix_formulas(ws):
    last_row = get_last_data_row(ws)

    for col in [27, 28]:
        col_letter = ws.cell(row=4, column=col).column_letter
        base = ws[f"{col_letter}4"].value

        for r in range(5, 1000):
            ws[f"{col_letter}{r}"] = None

        for r in range(5, last_row + 1):
            ws[f"{col_letter}{r}"] = Translator(base, origin=f"{col_letter}4").translate_formula(f"{col_letter}{r}")


def append(ws, df):
    start = get_last_data_row(ws) + 1

    for r, row in enumerate(df.itertuples(index=False), start=start):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(r, c, val)

            if c in [12, 13, 16] and pd.notna(val):
                cell.number_format = "DD/MM/YYYY"

            if c == 17 and pd.notna(val):
                cell.number_format = "HH:MM"


def fix_summary(ws, start_row, data_ws):
    base_row = start_row + 1
    data_len = get_last_data_row(data_ws) - 3

    for r in range(base_row, base_row + data_len + 50):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    for i in range(data_len):
        for c in range(1, ws.max_column + 1):
            base = ws.cell(base_row, c).value
            if base:
                ws.cell(
                    base_row + i,
                    c,
                    Translator(base, origin=f"A{base_row}").translate_formula(f"A{base_row+i}")
                )


# ============================================================
# MAIN
# ============================================================

if csv_file and live_file:

    df = pd.read_csv(csv_file, dtype=str).fillna("")
    av_df, narv_df = map_data(df)

    wb = load_workbook(live_file)

    weekly = wb["Weekly"]
    narv_weekly = wb["NARV Weekly"]
    ytd = wb["YTD"]
    narv_ytd = wb["NARV YTD"]

    # DEDUPE
    existing = {
        str(r[2].value)
        for ws in [weekly, narv_weekly]
        for r in ws.iter_rows(min_row=4, max_col=3)
        if r[2].value
    }

    av_df = av_df[~av_df["Visit"].isin(existing)]
    narv_df = narv_df[~narv_df["Visit"].isin(existing)]

    # WRITE
    write(weekly, av_df)
    write(narv_weekly, narv_df)

    # YTD
    append(ytd, av_df)
    append(narv_ytd, narv_df)

    # FORMULAS
    for ws in [weekly, narv_weekly, ytd, narv_ytd]:
        fix_formulas(ws)

    # SUMMARY
    fix_summary(wb["Weekly Summary Table"], 21, weekly)
    fix_summary(wb["Allergens Weekly Summary Table"], 16, narv_weekly)

    # SAVE LIVE
    today = datetime.today().strftime("%d.%m.%Y")
    live_buffer = io.BytesIO()
    wb.save(live_buffer)
    live_buffer.seek(0)

    # COMPLETED (VALUES ONLY, PRESERVE FORMAT)
    wb_values = load_workbook(live_buffer, data_only=True)
    wb_completed = load_workbook(live_buffer)

    keep = [
        "Weekly Summary Table",
        " Performance by Area",
        "Allergens Weekly Summary Table",
        " Allergens Performance by Area"
    ]

    for sheet in list(wb_completed.sheetnames):
        if sheet not in keep:
            del wb_completed[sheet]

    for ws in wb_completed.worksheets:
        ws_val = wb_values[ws.title]
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if not isinstance(cell, MergedCell):
                    cell.value = ws_val.cell(r, c).value

    completed_buffer = io.BytesIO()
    wb_completed.save(completed_buffer)
    completed_buffer.seek(0)

    # DOWNLOAD
    st.download_button("Download LIVE Report", live_buffer,
                       file_name=f"Weekly Report format - {today} LIVE.xlsx")

    st.download_button("Download Completed Report", completed_buffer,
                       file_name=f"Weekly Report format - {today}.xlsx")
