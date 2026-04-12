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
# COLUMN MAP (RESTORED LOGIC)
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

    # ✅ FIXED DATE HANDLING
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
    # DEDUPLICATION (FIXED)
    # ============================================================

    def get_ids(ws):
        return {
            str(r[2].value)
            for r in ws.iter_rows(min_row=4, max_col=3)
            if r[2].value
        }

    existing = get_ids(weekly) | get_ids(narv_weekly)

    av_df = av_df[~av_df["Visit"].isin(existing)]
    narv_df = narv_df[~narv_df["Visit"].isin(existing)]

    # ============================================================
    # WRITE FUNCTION
    # ============================================================

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

    write(weekly, av_df)
    write(narv_weekly, narv_df)

    # ============================================================
    # YTD (RESTORED)
    # ============================================================

    def append(ws, df):
        start = ws.max_row + 1
        for r, row in enumerate(df.itertuples(index=False), start=start):
            for c, val in enumerate(row, start=1):
                ws.cell(r, c, val)

    append(ytd, av_df)
    append(narv_ytd, narv_df)

    # ============================================================
    # FORMULAS (AA:AB FIXED)
    # ============================================================

    def fix_formulas(ws):
        last_row = ws.max_row

        for col in [27, 28]:
            col_letter = ws.cell(row=4, column=col).column_letter
            base = ws[f"{col_letter}4"].value

            # clear
            for r in range(5, ws.max_row + 200):
                ws[f"{col_letter}{r}"] = None

            # fill only needed rows
            for r in range(5, last_row + 1):
                ws[f"{col_letter}{r}"] = Translator(base, origin=f"{col_letter}4").translate_formula(f"{col_letter}{r}")

    for ws in [weekly, narv_weekly, ytd, narv_ytd]:
        fix_formulas(ws)

    # ============================================================
    # SUMMARY TABLES (FIXED LENGTH)
    # ============================================================

    def fix_summary(ws, start_row, data_ws):
        base_row = start_row + 1
        data_len = data_ws.max_row - 3

        base = [ws.cell(base_row, c).value for c in range(1, ws.max_column + 1)]

        # clear only required range
        for r in range(base_row, base_row + data_len + 5):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = None

        for i in range(data_len):
            for c, f in enumerate(base, start=1):
                if f:
                    ws.cell(
                        base_row + i,
                        c,
                        Translator(f, origin=f"A{base_row}").translate_formula(f"A{base_row+i}")
                    )

    fix_summary(wb["Weekly Summary Table"], 21, weekly)
    fix_summary(wb["Allergens Weekly Summary Table"], 16, narv_weekly)

    # ============================================================
    # SAVE LIVE
    # ============================================================

    today = datetime.today().strftime("%d.%m.%Y")

    live_buffer = io.BytesIO()
    wb.save(live_buffer)
    live_buffer.seek(0)

    # ============================================================
    # COMPLETED REPORT (NO DATA LOSS)
    # ============================================================

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

    completed_buffer = io.BytesIO()
    wb_completed.save(completed_buffer)
    completed_buffer.seek(0)

    # ============================================================
    # DOWNLOAD
    # ============================================================

    st.download_button("Download LIVE Report", live_buffer,
                       file_name=f"Weekly Report format - {today} LIVE.xlsx")

    st.download_button("Download Completed Report", completed_buffer,
                       file_name=f"Weekly Report format - {today}.xlsx")
