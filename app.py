import io
from processor import process_and_transform_excel
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Excel Grouping & Aggregator App", page_icon="📈", layout="wide"
)

st.title("📊 Multi-Excel Column Aggregator & Grouping Tool")
st.write(
    "10-20 Excel files select karke upload karein. Column Z ke duplicates merge hoke single row banenge aur positive/negative values Sum hongi."
)

uploaded_files = st.file_uploader(
    "Excel Files Choose Karein (Multiple Upload)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.info(f"Total {len(uploaded_files)} files select ki hain.")

    if st.button("⚡ Process & Group Excel Data"):
        with st.spinner("Processing Excel files..."):
            final_data = process_and_transform_excel(uploaded_files)

        if not final_data.empty:
            st.success("✅ Excel Data Successfully Grouped & Merged!")

            # Display Preview
            st.write("### Converted Data Preview:")
            st.dataframe(final_data)

            # Convert to Excel in-memory
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                final_data.to_excel(
                    writer, index=False, sheet_name="Merged_Summary"
                )

            # Download Button
            st.download_button(
                label="📥 New Converted Excel File Download Karein",
                data=output_buffer.getvalue(),
                file_name="Converted_Aggregated_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("Error: Selected files me valid data nahi mila.")
