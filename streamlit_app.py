import io
import os
import sys
import pandas as pd
import streamlit as st

# Prevent running with plain python; instruct to use `streamlit run`
if __name__ == "__main__" and "streamlit" not in " ".join(sys.argv):
    print("Please run this app with Streamlit:  streamlit run /workspaces/Mpesa-analyzer/streamlit_app.py")
    sys.exit(0)

st.set_page_config(page_title="Mpesa Analyzer", layout="wide")
st.title("📥 Mpesa Analyzer")

st.markdown(
    "Upload an Excel workbook, enter the required column names (comma-separated), "
    "then click **Analyze Sheets** to combine sheets that contain all required columns."
)

uploaded_file = st.file_uploader("Upload Excel file (.xlsx, .xls)", type=["xlsx", "xls"])
required_cols_input = st.text_input("Required columns (comma-separated)", value="Paid In, Withdrawn, Balance")
case_insensitive = st.checkbox("Case-insensitive column matching", value=True)

merge_button = st.button("Analyze Sheets")

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
                # Read Excel once into a dict of DataFrames
                excel_bytes = uploaded_file.read()
                sheets = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=None)

                # Removed: st.write(f"Found sheets: {list(sheets.keys())}")

                merged_dfs = []
                included_sheets = []
                skipped = {}

                req_lower = [c.lower() for c in required_columns]

                for sheet_name, df in sheets.items():
                    try:
                        cols = df.columns.tolist()
                        if case_insensitive:
                            cols_lower = [c.lower() for c in cols]
                            has_all = all(r in cols_lower for r in req_lower)
                        else:
                            has_all = all(r in cols for r in required_columns)

                        if has_all:
                            merged_dfs.append(df)
                            included_sheets.append(sheet_name)
                        else:
                            if case_insensitive:
                                missing = [r for r in req_lower if r not in cols_lower]
                            else:
                                missing = [r for r in required_columns if r not in cols]
                            skipped[sheet_name] = missing
                    except Exception as e:
                        skipped[sheet_name] = f"read error: {e}"

                if merged_dfs:
                    final_df = pd.concat(merged_dfs, ignore_index=True)

                    # Removed: st.success(f"Merged {len(included_sheets)} sheet(s): {included_sheets}")
                    # Removed: if skipped: st.info(f"Skipped sheets and reasons: {skipped}")

                    # Removed: preview heading and dataframe display
                    # st.subheader("Preview (first 100 rows)")
                    # st.dataframe(final_df.head(100))

                    # Normalize column name lookup helpers
                    def find_col(df, target):
                        for c in df.columns:
                            if c.lower() == target.lower():
                                return c
                        for c in df.columns:
                            if target.lower() in c.lower():
                                return c
                        return None

                    paid_col = find_col(final_df, "Paid In")
                    withdrawn_col = find_col(final_df, "Withdrawn")

                    # Prepare merged file downloads (Excel and CSV)
                    try:
                        towrite = io.BytesIO()
                        final_df.to_excel(towrite, index=False, engine="openpyxl")
                        towrite.seek(0)
                        st.download_button(
                            label="Download merged Excel (.xlsx)",
                            data=towrite,
                            file_name="merged_mpesa_statement.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception:
                        # fallback without specifying engine
                        towrite = io.BytesIO()
                        final_df.to_excel(towrite, index=False)
                        towrite.seek(0)
                        st.download_button(
                            label="Download merged Excel (.xlsx)",
                            data=towrite,
                            file_name="merged_mpesa_statement.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

                    csv_bytes = final_df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        label="Download merged CSV",
                        data=csv_bytes,
                        file_name="merged_mpesa_statement.csv",
                        mime="text/csv",
                    )

                    # --------------------------
                    # Pivot table by Month (sum of Paid In and Withdrawn)
                    # --------------------------
                    st.markdown("---")
                    st.subheader("Pivot: Sum of Paid In and Withdrawn by Month")

                    # Detect candidate date columns
                    candidate_date_cols = [
                        c for c in final_df.columns
                        if any(k in c.lower() for k in ("date", "time", "completion", "timestamp"))
                    ]
                    if candidate_date_cols:
                        date_col = st.selectbox("Select date/time column to use for Month grouping", options=candidate_date_cols, index=0)
                    else:
                        date_col = st.text_input("No obvious date column found. Enter the date column name manually", value="")

                    if paid_col is None or withdrawn_col is None:
                        st.warning("Could not locate 'Paid In' and/or 'Withdrawn' columns in the merged data.")
                        st.write("Detected columns:", list(final_df.columns))
                    else:
                        if not date_col:
                            st.warning("No date column selected. Pivot cannot be computed until a date column is provided.")
                        else:
                            # parse dates
                            final_df["_parsed_date"] = pd.to_datetime(final_df[date_col], errors="coerce")
                            if final_df["_parsed_date"].isna().all():
                                st.error("Could not parse any valid dates from the selected date column. Check format or select another column.")
                            else:
                                # coerce numeric on paid/withdrawn
                                final_df[paid_col] = pd.to_numeric(final_df[paid_col], errors="coerce")
                                final_df[withdrawn_col] = pd.to_numeric(final_df[withdrawn_col], errors="coerce")

                                # Create a human-friendly month label (e.g., "January 2025")
                                final_df["_MonthLabel"] = final_df["_parsed_date"].dt.strftime("%B %Y")

                                pivot_df = (
                                    final_df
                                    .groupby(["_MonthLabel"], dropna=True)[[paid_col, withdrawn_col]]
                                    .sum(min_count=1)
                                    .reset_index()
                                    .rename(columns={"_MonthLabel": "Month", paid_col: "Sum Paid In", withdrawn_col: "Sum Withdrawn"})
                                )

                                # Optional: sort months chronologically
                                try:
                                    # create a helper datetime for sorting, then drop it
                                    pivot_df["_sort_dt"] = pd.to_datetime(pivot_df["Month"], format="%B %Y", errors="coerce")
                                    pivot_df = pivot_df.sort_values("_sort_dt").drop(columns=["_sort_dt"]).reset_index(drop=True)
                                except Exception:
                                    pass

                                st.write("Pivot (aggregated by month):")
                                st.dataframe(pivot_df)

                                # Download pivot as CSV and Excel
                                pivot_csv = pivot_df.to_csv(index=False).encode("utf-8")
                                st.download_button("Download pivot CSV", data=pivot_csv, file_name="pivot_mpesa.csv", mime="text/csv")

                                try:
                                    piv_bytes = io.BytesIO()
                                    pivot_df.to_excel(piv_bytes, index=False, engine="openpyxl")
                                    piv_bytes.seek(0)
                                    st.download_button(
                                        label="Download pivot Excel (.xlsx)",
                                        data=piv_bytes,
                                        file_name="pivot_mpesa.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    )
                                except Exception:
                                    # fallback
                                    piv_bytes = io.BytesIO()
                                    pivot_df.to_excel(piv_bytes, index=False)
                                    piv_bytes.seek(0)
                                    st.download_button(
                                        label="Download pivot Excel (.xlsx)",
                                        data=piv_bytes,
                                        file_name="pivot_mpesa.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    )

                                # Printable HTML view
                                pivot_html_table = pivot_df.to_html(index=False, classes="pivot-table", border=0, float_format="%.2f")
                                printable_html = f"""
                                <html>
                                  <head>
                                    <meta charset="utf-8"/>
                                    <style>
                                      body{{font-family: Arial, sans-serif; padding: 16px;}}
                                      table.pivot-table{{border-collapse: collapse; width:100%;}}
                                      table.pivot-table th, table.pivot-table td{{border:1px solid #ccc; padding:6px; text-align:left;}}
                                      .print-btn{{display:inline-block; margin-bottom:12px; padding:8px 12px; background:#1976d2; color:white; border-radius:4px; cursor:pointer; text-decoration:none;}}
                                    </style>
                                  </head>
                                  <body>
                                    <a class="print-btn" onclick="window.print()">Print pivot table</a>
                                    {pivot_html_table}
                                  </body>
                                </html>
                                """
                                # Render printable HTML inside app
                                st.components.v1.html(printable_html, height=500, scrolling=True)

                else:
                    st.warning("No sheets contained all the specified required columns.")
                    # Removed: if skipped: st.write("Skipped sheets and reasons:", skipped)

        except Exception as exc:
            st.error(f"An error occurred: {exc}")
