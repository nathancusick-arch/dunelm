import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.cell.cell import MergedCell
import io
from datetime import datetime

st.title("Dunelm Report Generator")

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

    # ================= TYPE FIXES =================
    final_df["Actual Visit Date"] = pd.to_datetime(final_df["Actual Visit Date"], errors="coerce")
    final_df["Actual Visit Time"] = pd.to_datetime(final_df["Actual Visit Time"], errors="coerce").dt.time
    final_df["Submitted Date"] = pd.to_datetime(final_df["Submitted Date"], errors="coerce")
    final_df["Approved Date"] = pd.to_datetime(final_df["Approved Date"], errors="coerce")
    final_df["Extra Site 1"] = pd.to_numeric(final_df["Extra Site 1"], errors="coerce")

    av = final_df[~final_df["Item to order"].isin(["Allergens", "Restaurant", "On the Go"])].copy()
    narv = final_df[final_df["Item to order"].isin(["Allergens", "Restaurant", "On the Go"])].copy()

    return av, narv


# ============================================================
# MAIN PROCESS
# ============================================================

if csv_file and live_file:

    st.info("Processing...")

    df = pd.read_csv(csv_file, dtype=str).fillna("")
    av_df, narv_df = map_data(df)

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
            if row[2].value:
                visits.add(str(row[2].value))
        return visits

    existing_visits = get_existing_visits(weekly_ws) | get_existing_visits(narv_weekly_ws)

    av_df = av_df[~av_df["Visit"].isin(existing_visits)]
    narv_df = narv_df[~narv_df["Visit"].isin(existing_visits)]

    # ============================================================
    # YEAR CHECK
    # ============================================================

    def get_year(ws):
        for row in ws.iter_rows(min_row=4, max_col=16):
            if row[15].value:
                return pd.to_datetime(row[15].value).year
        return None

    new_year = pd.to_datetime(av_df["Actual Visit Date"]).min().year if not av_df.empty else None
    existing_year = get_year(ytd_ws)
    reset_ytd = new_year and existing_year and new_year > existing_year

    # ============================================================
    # WRITE DATA
    # ============================================================

    def clear_data(ws):
        for row in ws.iter_rows(min_row=4, max_col=30):
            for cell in row:
                if not isinstance(cell, MergedCell):
                    cell.value = None

    def write_df(ws, df):
        clear_data(ws)
        for r, row in enumerate(df.values, start=4):
            for c, val in enumerate(row, start=1):
                ws.cell(row=r, column=c, value=val)

    write_df(weekly_ws, av_df)
    write_df(narv_weekly_ws, narv_df)

    def append_df(ws, df, reset):
        if reset:
            clear_data(ws)
            start = 4
        else:
            start = ws.max_row + 1

        for r, row in enumerate(df.values, start=start):
            for c, val in enumerate(row, start=1):
                ws.cell(row=r, column=c, value=val)

    append_df(ytd_ws, av_df, reset_ytd)
    append_df(narv_ytd_ws, narv_df, reset_ytd)

    # ============================================================
    # FORMULAS (AA:AB) WITH CLEARING
    # ============================================================

    def extend_formulas(ws):
        max_row = ws.max_row

        for col in [27, 28]:
            col_letter = ws.cell(row=4, column=col).column_letter
            base = ws[f"{col_letter}4"].value

            # clear column first
            for r in range(5, ws.max_row + 50):
                ws[f"{col_letter}{r}"] = None

            for r in range(5, max_row + 1):
                ws[f"{col_letter}{r}"] = Translator(base, origin=f"{col_letter}4").translate_formula(f"{col_letter}{r}")

    for ws in [weekly_ws, ytd_ws, narv_weekly_ws, narv_ytd_ws]:
        extend_formulas(ws)

    # ============================================================
    # SUMMARY TABLE RESIZE
    # ============================================================

    def resize_summary(ws, start_row, data_length):
        formula_row = start_row + 1

        for r in range(formula_row, formula_row + 1000):
            for c in range(1, 50):
                ws.cell(row=r, column=c).value = None

        for i in range(data_length):
            for c in range(1, 50):
                base = ws.cell(row=formula_row, column=c).value
                if base:
                    ws.cell(
                        row=formula_row + i,
                        column=c,
                        value=Translator(base, origin=f"A{formula_row}").translate_formula(f"A{formula_row+i}")
                    )

    resize_summary(wb["Weekly Summary Table"], 21, len(av_df))
    resize_summary(wb["Allergens Weekly Summary Table"], 16, len(narv_df))

    # ============================================================
    # SAVE LIVE
    # ============================================================

    today = datetime.today().strftime("%d.%m.%Y")

    live_buffer = io.BytesIO()
    wb.save(live_buffer)
    live_buffer.seek(0)

    # ============================================================
    # COMPLETED REPORT (TRUE VALUES)
    # ============================================================

    wb_values = load_workbook(live_buffer, data_only=True)
    wb_completed = load_workbook(live_buffer)

    keep = [
        "Weekly Summary Table",
        "Performance by Area",
        "Allergens Weekly Summary Table",
        "Allergens Performance by Area"
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

    # ============================================================
    # DOWNLOADS
    # ============================================================

    st.success("Done!")

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
