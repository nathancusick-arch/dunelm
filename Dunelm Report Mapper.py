import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.formula.translate import Translator
from openpyxl.cell.cell import MergedCell
import io
from datetime import datetime

st.title("Dunelm Report Generator")

st.write("""
1. Upload the latest audits export file  
2. Upload last week's LIVE report  
3. Download this week's LIVE and completed reports  

The tool will automatically:
- Remove audits already reported
- Update Weekly and YTD tabs
- Rebuild summary tables
- Generate both LIVE and completed reports
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
        return row.get(mapping, "")

    final_df = pd.DataFrame(columns=OUTPUT_COLUMNS)

    for col in OUTPUT_COLUMNS:
        final_df[col] = df.apply(lambda r: map_value(r, COLUMN_MAP[col]), axis=1)

    # Type handling
    final_df["Actual Visit Date"] = pd.to_datetime(final_df["Actual Visit Date"], errors="coerce")
    final_df["Submitted Date"] = pd.to_datetime(final_df["Submitted Date"], errors="coerce")
    final_df["Approved Date"] = pd.to_datetime(final_df["Approved Date"], errors="coerce")
    final_df["Actual Visit Time"] = pd.to_datetime(final_df["Actual Visit Time"], errors="coerce").dt.time
    final_df["Extra Site 1"] = pd.to_numeric(final_df["Extra Site 1"], errors="coerce")

    av = final_df[~final_df["Item to order"].isin(["Allergens", "Restaurant", "On the Go"])].copy()
    narv = final_df[final_df["Item to order"].isin(["Allergens", "Restaurant", "On the Go"])].copy()

    return av, narv


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

    # ============================================================
    # DEDUPLICATION
    # ============================================================

    def get_ids(ws):
        return {
            str(r[2].value)
            for r in ws.iter_rows(min_row=4, max_col=3)
            if r[2].value
        }

    existing_ids = get_ids(weekly) | get_ids(narv_weekly)

    av_df = av_df[~av_df["Visit"].isin(existing_ids)]
    narv_df = narv_df[~narv_df["Visit"].isin(existing_ids)]

    # ============================================================
    # WRITE DATA WITH FORMATTING
    # ============================================================

    def write(ws, df):
        # Clear A:Y only
        for row in ws.iter_rows(min_row=4, max_col=25):
            for cell in row:
                if not isinstance(cell, MergedCell):
                    cell.value = None

        for r, row in enumerate(df.itertuples(index=False), start=4):
            for c, val in enumerate(row, start=1):
                cell = ws.cell(r, c, val)

                # Dates
                if c in [12, 13, 16] and pd.notna(val):
                    cell.number_format = "DD/MM/YYYY"

                # Time
                if c == 17 and pd.notna(val):
                    cell.number_format = "HH:MM"

    write(weekly, av_df)
    write(narv_weekly, narv_df)

    # ============================================================
    # FORMULA EXTENSION
    # ============================================================

    def fill_formulas(ws):
        max_row = ws.max_row
        for col in [27, 28]:  # AA, AB
            col_letter = ws.cell(row=4, column=col).column_letter
            base_formula = ws[f"{col_letter}4"].value

            if not base_formula:
                continue

            for r in range(5, max_row + 1):
                ws[f"{col_letter}{r}"] = Translator(
                    base_formula,
                    origin=f"{col_letter}4"
                ).translate_formula(f"{col_letter}{r}")

    for ws in [weekly, narv_weekly, ytd, narv_ytd]:
        fill_formulas(ws)

    # ============================================================
    # SUMMARY TABLES
    # ============================================================

    def rebuild_summary(ws, start_row, count):
        base_row = start_row + 1

        base_formulas = [
            ws.cell(base_row, c).value
            for c in range(1, ws.max_column + 1)
        ]

        # Clear existing
        for r in range(base_row, base_row + 1000):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = None

        # Rebuild
        for i in range(count):
            for c, formula in enumerate(base_formulas, start=1):
                if formula:
                    ws.cell(
                        base_row + i,
                        c,
                        Translator(formula, origin=f"A{base_row}").translate_formula(f"A{base_row+i}")
                    )

    rebuild_summary(wb["Weekly Summary Table"], 21, len(av_df))
    rebuild_summary(wb["Allergens Weekly Summary Table"], 16, len(narv_df))

    # ============================================================
    # SAVE LIVE REPORT
    # ============================================================

    today = datetime.today().strftime("%d.%m.%Y")

    live_buffer = io.BytesIO()
    wb.save(live_buffer)
    live_buffer.seek(0)

    # ============================================================
    # COMPLETED REPORT (EXACT SHEET NAMES)
    # ============================================================

    wb_values = load_workbook(live_buffer, data_only=True)
    wb_completed = Workbook()

    keep_sheets = [
        "Weekly Summary Table",
        " Performance by Area",
        "Allergens Weekly Summary Table",
        " Allergens Performance by Area"
    ]

    for name in keep_sheets:
        if name not in wb_values.sheetnames:
            continue

        src = wb_values[name]
        dest = wb_completed.create_sheet(title=name)

        for r in range(1, src.max_row + 1):
            for c in range(1, src.max_column + 1):
                dest.cell(r, c, src.cell(r, c).value)

    # Remove default sheet
    if "Sheet" in wb_completed.sheetnames:
        del wb_completed["Sheet"]

    completed_buffer = io.BytesIO()
    wb_completed.save(completed_buffer)
    completed_buffer.seek(0)

    # ============================================================
    # DOWNLOADS
    # ============================================================

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
