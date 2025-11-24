import streamlit as st
import pandas as pd
import io

st.title("Dunelm Weekly Report Mapper")

st.write("""
          1. Export the previous 2 weeks worth of data
          2. Drop the file in the below box, it should then give you the output file in your downloads
          3. Standard bits - Check data vs previous week, remove data already reported, paste over new data
          4. Copy and paste over values etc!!!
          5. Done.
          """)

# ============================================================
# FILE UPLOADER
# ============================================================

csv_file = st.file_uploader("Upload audits_basic_data_export.csv", type=["csv"])

# ============================================================
# PROCESS ONLY WHEN CSV IS UPLOADED
# ============================================================

if csv_file is not None:

    # Load export
    df = pd.read_csv(csv_file, dtype=str).fillna("")

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

    # Output columns derive directly from mapping keys
    OUTPUT_COLUMNS = list(COLUMN_MAP.keys())

    # ============================================================
    # APPLY MAPPING
    # ============================================================

    def map_value(row, mapping):
        if mapping is None:
            return ""
        return str(row.get(mapping, "")).strip()

    final_df = pd.DataFrame(columns=OUTPUT_COLUMNS)

    for col in OUTPUT_COLUMNS:
        final_df[col] = df.apply(lambda r: map_value(r, COLUMN_MAP[col]), axis=1)

    # ============================================================
    # PREVIEW & DOWNLOAD
    # ============================================================

    st.subheader("Preview of Output")
    st.write(final_df)

    output_buffer = io.BytesIO()
    final_df.to_csv(output_buffer, index=False, encoding="utf-8-sig")
    output_buffer.seek(0)

    st.download_button(
        label="Download Dunelm Weekly Report CSV",
        data=output_buffer,
        file_name="Dunelm Report Data.csv",
        mime="text/csv"
    )
