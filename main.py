import streamlit as st
import pandas as pd
import time
from io import BytesIO

st.set_page_config(layout="wide")

st.title("Billed Status")

# ---------------- FILE UPLOAD SECTION ----------------

st.subheader("Upload Files")

col1, col2 = st.columns(2)

with col1:
    uploaded_file = st.file_uploader(
        "Upload PH File (Excel)",
        type=["xlsx", "xls"]
    )

    st.caption(
        """
        Required Sheet Name: PH
        
        Required Columns:
        company_no, locn_no, phreq_no, so_no, To_Date, Totarrear
        
        Note:
        Headers can be located anywhere within the first 10 rows.
        """
    )

with col2:
    uploaded_file_2 = st.file_uploader(
        "Upload Billing File (Excel)",
        type=["xlsx", "xls"]
    )

    st.caption(
        """
        Upload a single Billing workbook containing BOTH required sheets:
        
        Required Sheet Names:
        - Billed
        - Not Billed
        
        Required Columns:
        so_number, dont_bill_reason, dont_bill_remarks
        
        Note:
        Headers can be located anywhere within the first 10 rows.
        """
    )
    

# ---------------- RUN BUTTON ----------------

run = st.button("Run")

# ---------------- PROCESSING CONTAINER ----------------

log_container = st.container()

# ---------------- HEADER FINDER FUNCTION ----------------

def find_header_row(df_raw, required_cols):
    """
    Searches first 10 rows for required headers.
    """
    for i in range(min(10, len(df_raw))):
        temp_cols = df_raw.iloc[i].astype(str).str.strip().tolist()

        if all(col in temp_cols for col in required_cols):
            df_raw.columns = temp_cols
            df_cleaned = df_raw.iloc[i + 1:].reset_index(drop=True)
            return df_cleaned

    return None

# ---------------- MAIN PROCESS ----------------

if run:

    with log_container:

        status_text = st.empty()

        try:

            # ---------------- FILE VALIDATION ----------------

            status_text.info("Validating uploaded files...")
            time.sleep(0.2)

            if uploaded_file is None or uploaded_file_2 is None:
                st.error("Please upload both files before running.")
                st.stop()

            # ---------------- READ PH FILE ----------------

            status_text.info("Reading PH file...")
            time.sleep(0.2)

            try:
                ph_raw = pd.read_excel(
                    uploaded_file,
                    sheet_name="PH",
                    header=None
                )
            except Exception as e:
                st.error(f"Error reading PH file or sheet 'PH' not found: {e}")
                st.stop()

            # ---------------- READ BILLING FILE ----------------

            status_text.info("Reading Billing file...")
            time.sleep(0.2)

            try:
                billed_raw = pd.read_excel(
                    uploaded_file_2,
                    sheet_name="Billed",
                    header=None
                )

                notbilled_raw = pd.read_excel(
                    uploaded_file_2,
                    sheet_name="Not Billed",
                    header=None
                )

            except Exception as e:
                st.error(
                    f"Error reading Billing file. Required sheets: "
                    f"'Billed' and 'Not Billed'. Error: {e}"
                )
                st.stop()

            # ---------------- REQUIRED COLUMNS ----------------

            required_ph_cols = [
                "company_no",
                "locn_no",
                "phreq_no",
                "so_no",
                "To_Date",
                "Totarrear"
            ]

            required_billed_cols = [
                "so_number"
            ]

            required_notbilled_cols = [
                "so_number",
                "dont_bill_reason",
                "dont_bill_remarks"
            ]

            # ---------------- HEADER DETECTION ----------------

            status_text.info("Detecting headers in PH sheet...")
            time.sleep(0.2)

            try:
                ph = find_header_row(ph_raw, required_ph_cols)

                if ph is None:
                    st.error(
                        "Could not detect required headers in PH sheet "
                        "within first 10 rows."
                    )
                    st.stop()

            except Exception as e:
                st.error(f"Error during PH header detection: {e}")
                st.stop()

            status_text.info("Detecting headers in Billed sheet...")
            time.sleep(0.2)

            try:
                Billed = find_header_row(
                    billed_raw,
                    required_billed_cols
                )

                if Billed is None:
                    st.error(
                        "Could not detect required headers in Billed sheet "
                        "within first 10 rows."
                    )
                    st.stop()

            except Exception as e:
                st.error(f"Error during Billed header detection: {e}")
                st.stop()

            status_text.info("Detecting headers in Not Billed sheet...")
            time.sleep(0.2)

            try:
                Notbilled = find_header_row(
                    notbilled_raw,
                    required_notbilled_cols
                )

                if Notbilled is None:
                    st.error(
                        "Could not detect required headers in "
                        "Not Billed sheet within first 10 rows."
                    )
                    st.stop()

            except Exception as e:
                st.error(f"Error during Not Billed header detection: {e}")
                st.stop()

            # ---------------- COLUMN VALIDATION ----------------

            status_text.info("Validating required columns...")
            time.sleep(0.2)

            try:

                missing_ph = [
                    col for col in required_ph_cols
                    if col not in ph.columns
                ]

                missing_billed = [
                    col for col in required_billed_cols
                    if col not in Billed.columns
                ]

                missing_notbilled = [
                    col for col in required_notbilled_cols
                    if col not in Notbilled.columns
                ]

                if missing_ph:
                    st.error(f"Missing columns in PH sheet: {missing_ph}")
                    st.stop()

                if missing_billed:
                    st.error(
                        f"Missing columns in Billed sheet: {missing_billed}"
                    )
                    st.stop()

                if missing_notbilled:
                    st.error(
                        f"Missing columns in Not Billed sheet: "
                        f"{missing_notbilled}"
                    )
                    st.stop()

            except Exception as e:
                st.error(f"Error during column validation: {e}")
                st.stop()

            # ---------------- DATA CLEANING ----------------

            status_text.info("Cleaning and standardizing data...")
            time.sleep(0.2)

            try:

                ph_key = "so_no"
                billed_key = "so_number"
                notbilled_key = "so_number"

                Billed = Billed[[billed_key]].copy()

                Notbilled = Notbilled[
                    [
                        notbilled_key,
                        "dont_bill_reason",
                        "dont_bill_remarks"
                    ]
                ].copy()

                Billed["Status"] = "Billed"

                Notbilled["Status"] = "Not Billed"

                ph[ph_key] = (
                    ph[ph_key]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                Billed[billed_key] = (
                    Billed[billed_key]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                Notbilled[notbilled_key] = (
                    Notbilled[notbilled_key]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

            except Exception as e:
                st.error(f"Error during data cleaning: {e}")
                st.stop()

            # ---------------- MERGE PROCESS ----------------

            status_text.info("Merging billed data...")
            time.sleep(0.2)

            try:

                result = pd.merge(
                    ph,
                    Billed,
                    left_on=ph_key,
                    right_on=billed_key,
                    how="left"
                )

            except Exception as e:
                st.error(f"Error during billed merge: {e}")
                st.stop()

            status_text.info("Merging not billed data...")
            time.sleep(0.2)

            try:

                result = pd.merge(
                    result,
                    Notbilled,
                    left_on=ph_key,
                    right_on=notbilled_key,
                    how="left",
                    suffixes=("_billed", "_notbilled")
                )

            except Exception as e:
                st.error(f"Error during not billed merge: {e}")
                st.stop()

            # ---------------- FINAL PROCESSING ----------------

            status_text.info("Preparing final output...")
            time.sleep(0.2)

            try:

                result["Status"] = result[
                    "Status_billed"
                ].combine_first(
                    result["Status_notbilled"]
                )

                result["Status"] = result["Status"].fillna("NA")

                result = result.drop_duplicates(
                    subset=ph_key,
                    keep="first"
                )

                result = result[
                    [
                        "company_no",
                        "locn_no",
                        "phreq_no",
                        "so_no",
                        "To_Date",
                        "Totarrear",
                        "Status",
                        "dont_bill_reason",
                        "dont_bill_remarks"
                    ]
                ]

            except Exception as e:
                st.error(f"Error during final processing: {e}")
                st.stop()

            # ---------------- OUTPUT FILE ----------------

            status_text.info("Generating output file...")
            time.sleep(0.2)

            try:

                output = BytesIO()

                with pd.ExcelWriter(
                    output,
                    engine="openpyxl"
                ) as writer:
                    result.to_excel(
                        writer,
                        index=False,
                        sheet_name="Output"
                    )

                output.seek(0)

            except Exception as e:
                st.error(f"Error while generating output file: {e}")
                st.stop()

            # ---------------- SUCCESS ----------------

            status_text.success("Processing completed successfully.")

            st.dataframe(
                result,
                use_container_width=True
            )

            st.download_button(
                label="Download Output Excel",
                data=output,
                file_name="billed_status_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Unexpected error occurred: {e}")
            st.stop()

# ---------------- DOCUMENTATION SECTION ----------------

with st.expander("What This Tool Does"):

    st.write(
        """
        This tool compares PH records against Billed and Not Billed records
        to determine billing status for each SO number.

        It identifies whether each SO is:
        - Billed
        - Not Billed
        - Not Available in either source

        The final report also includes billing remarks and reasons
        wherever applicable.
        """
    )

with st.expander("How to Use"):

    st.write(
        """
        1. Upload the PH Excel file.
        2. Upload the Billing Excel file.
        3. Click Run.
        4. Wait for processing to complete.
        5. Download the final output report.
        """
    )

with st.expander("Output Details"):

    st.write(
        """
        The final output contains:

        - Company Number
        - Location Number
        - PH Request Number
        - SO Number
        - To Date
        - Total Arrear
        - Billing Status
        - Non-Billing Reason
        - Non-Billing Remarks

        Billing status is derived by checking whether the SO exists
        in Billed or Not Billed sheets.
        """
    )

with st.expander("Financial Logic"):

    st.write(
        """
        - If an SO number exists in the Billed sheet,
          status is marked as Billed.

        - If an SO number exists in the Not Billed sheet,
          status is marked as Not Billed.

        - If an SO number is not found in either sheet,
          status is marked as NA.

        - Non-billing remarks and reasons are pulled
          directly from the Not Billed sheet.
        """
    )
