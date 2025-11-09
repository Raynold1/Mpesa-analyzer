import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Mpesa Sheet Merger", layout="wide")
st.title("📥 Mpesa Sheets Merger")

st.markdown("Upload an Excel workbook, enter the required column names (comma-separated), then click **Merge Sheets** to combine sheets that contain all required columns.")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx, .xls)", type=["xlsx", "xls"])
required_cols_input = st.text_input("Required columns (comma-separated)", value="Paid In, Withdrawn, Balance")
case_insensitive = st.checkbox("Case-insensitive column matching", value=True)

merge_button = st.button("Merge Sheets")

if uploaded_file is None:
    st.info("Please upload an Excel file to begin.")
else:
    st.write(f"Uploaded file: {uploaded_file.name}")
    if merge_button:
        try:
            required_columns = [c.strip() for c in required_cols_input.split(",") if c.strip()]
            if not required_columns:
                st.error("Enter at least one required column.")
            else:
                # Read Excel from uploaded bytes
                excel_bytes = uploaded_file.read()
                xls = pd.ExcelFile(io.BytesIO(excel_bytes))
                st.write(f"Found sheets: {xls.sheet_names}")

                merged_dfs = []
                included_sheets = []
                skipped = {}

                for sheet_name in xls.sheet_names:
                    try:
                        df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet_name)
                        cols = df.columns.tolist()
                        if case_insensitive:
                            cols_lower = [c.lower() for c in cols]
                            req_lower = [c.lower() for c in required_columns]
                            has_all = all(r in cols_lower for r in req_lower)
                        else:
                            has_all = all(r in cols for r in required_columns)

                        if has_all:
                            merged_dfs.append(df)
                            included_sheets.append(sheet_name)
                        else:
                            # record which required cols are missing for this sheet
                            if case_insensitive:
                                missing = [r for r in req_lower if r not in cols_lower]
                            else:
                                missing = [r for r in required_columns if r not in cols]
                            skipped[sheet_name] = missing
                    except Exception as e:
                        skipped[sheet_name] = f"read error: {e}"

                if merged_dfs:
                    final_df = pd.concat(merged_dfs, ignore_index=True)

                    st.success(f"Merged {len(included_sheets)} sheet(s): {included_sheets}")
                    if skipped:
                        st.info(f"Skipped sheets and reasons: {skipped}")

                    st.subheader("Preview (first 100 rows)")
                    st.dataframe(final_df.head(100))

                    # Prepare Excel download
                    towrite = io.BytesIO()
                    try:
                        final_df.to_excel(towrite, index=False)
                        towrite.seek(0)
                        st.download_button(
                            label="Download merged Excel (.xlsx)",
                            data=towrite,
                            file_name="merged_mpesa_statement.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception as e:
                        st.error(f"Could not create Excel file: {e}")

                    # Also provide CSV download as fallback
                    csv_bytes = final_df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        label="Download merged CSV",
                        data=csv_bytes,
                        file_name="merged_mpesa_statement.csv",
                        mime="text/csv",
                    )
                else:
                    st.warning("No sheets contained all the specified required columns.")
                    if skipped:
                        st.write("Skipped sheets and reasons:", skipped)

        except Exception as exc:
            st.error(f"An error occurred: {exc}")
