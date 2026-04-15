import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")

st.title("Billed Status")

# ---------------- FILE UPLOAD ----------------
col1, col2 = st.columns(2)

with col1:
    uploaded_file = st.file_uploader("Upload PH File (Excel)", type=["xlsx", "xls"])
    st.caption("Sheet name: PH | Columns: company_no, locn_no, phreq_no, so_no, To_Date, Totarrear")

with col2:
    uploaded_file_2 = st.file_uploader("Upload Billing File (Excel)", type=["xlsx", "xls"])
    st.caption("Sheets: Billed, Not Billed | Columns: so_number, dont_bill_reason, dont_bill_remarks")

# ---------------- RUN BUTTON ----------------
run = st.button("Run")

# ---------------- PROCESS ----------------
if run:

    try:
        # ---------------- FILE VALIDATION ----------------
        if uploaded_file is None or uploaded_file_2 is None:
            st.error("Please upload both files before running.")
            st.stop()

        # ---------------- READ FILES ----------------
        try:
            ph = pd.read_excel(uploaded_file, sheet_name="PH")
        except Exception as e:
            st.error("Error reading PH file or sheet 'PH' not found.")
            st.stop()

        try:
            Billed = pd.read_excel(uploaded_file_2, sheet_name="Billed")
            Notbilled = pd.read_excel(uploaded_file_2, sheet_name="Not Billed")
        except Exception as e:
            st.error("Error reading Billing file. Required sheets: 'Billed' and 'Not Billed'.")
            st.stop()

        # ---------------- COLUMN VALIDATION ----------------
        required_ph_cols = ["company_no", "locn_no", "phreq_no", "so_no", "To_Date", "Totarrear"]
        required_billed_cols = ["so_number"]
        required_notbilled_cols = ["so_number", "dont_bill_reason", "dont_bill_remarks"]

        missing_ph = [col for col in required_ph_cols if col not in ph.columns]
        missing_billed = [col for col in required_billed_cols if col not in Billed.columns]
        missing_notbilled = [col for col in required_notbilled_cols if col not in Notbilled.columns]

        if missing_ph:
            st.error(f"Missing columns in PH file: {missing_ph}")
            st.stop()

        if missing_billed:
            st.error(f"Missing columns in Billed sheet: {missing_billed}")
            st.stop()

        if missing_notbilled:
            st.error(f"Missing columns in Not Billed sheet: {missing_notbilled}")
            st.stop()

        st.write("Cleaning data...")

        try:
            ph_key = "so_no"
            billed_key = "so_number"
            notbilled_key = "so_number"

            Billed = Billed[[billed_key]].copy()
            Notbilled = Notbilled[[notbilled_key, "dont_bill_reason", "dont_bill_remarks"]].copy()

            Billed["Status"] = "Billed"
            Notbilled["Status"] = "Not Billed"

            # Standardize keys
            ph[ph_key] = ph[ph_key].astype(str).str.strip().str.upper()
            Billed[billed_key] = Billed[billed_key].astype(str).str.strip().str.upper()
            Notbilled[notbilled_key] = Notbilled[notbilled_key].astype(str).str.strip().str.upper()

        except Exception as e:
            st.error(f"Error during data cleaning: {str(e)}")
            st.stop()

        st.write("Merging data...")

        try:
            result = pd.merge(ph, Billed, left_on=ph_key, right_on=billed_key, how="left")

            result = pd.merge(
                result,
                Notbilled,
                left_on=ph_key,
                right_on=notbilled_key,
                how="left",
                suffixes=("_billed", "_notbilled")
            )
        except Exception as e:
            st.error(f"Error during merge: {str(e)}")
            st.stop()

        st.write("Final processing...")

        try:
            result["Status"] = result["Status_billed"].combine_first(result["Status_notbilled"])
            result["Status"] = result["Status"].fillna("NA")

            result = result.drop_duplicates(subset=ph_key, keep="first")

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
            st.error(f"Error during final processing: {str(e)}")
            st.stop()

        st.success("Done")

        st.dataframe(result, use_container_width=True)

        try:
            csv = result.to_csv(index=False).encode("utf-8")

            st.download_button(
                label="Download Output CSV",
                data=csv,
                file_name="output.csv",
                mime="text/csv"
            )
        except Exception as e:
            st.error(f"Error while generating download file: {str(e)}")

    except Exception as e:
        st.error(f"Unexpected error occurred: {str(e)}")
