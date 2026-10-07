import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Converter App", layout="wide")

st.title("📊 Multi-Excel Converter App")
st.write(
    "Apni 10-20 Excel files select karke upload karein, data convert hoga aur aap single output Excel file download kar sakte hain."
)

# Multi-file uploader (Ek sath 10-20 Excel files browse/select kar sakte hain)
uploaded_files = st.file_uploader(
    "Excel files upload karein", type=["xlsx", "xls"], accept_multiple_files=True
)

if uploaded_files:
    st.success(f"Aapne total {len(uploaded_files)} files select ki hain.")

    # Processing Button
    if st.button("Data Convert Karein"):
        processed_data_list = []

        with st.spinner("Files process ho rahi hain..."):
            for uploaded_file in uploaded_files:
                try:
                    # Excel file read karein
                    df = pd.read_excel(uploaded_file)

                    # -----------------------------------------------------------
                    # AAPKA CUSTOM CONVERSION CODE YAHAN AAEGA
                    # Example conversion logic:
                    # 1. Column ke naam Clean / Trim karna
                    df.columns = df.columns.astype(str).str.strip()

                    # 2. File ka naam record ke liye ek naye column mein daalna
                    df["Source_File"] = uploaded_file.name

                    # 3. Aap chahein toh specific columns filter / transform kar sakte hain:
                    # df = df[['Column1', 'Column2', 'Source_File']]
                    # -----------------------------------------------------------

                    processed_data_list.append(df)

                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {e}")

        if processed_data_list:
            # Sabhi files ke data ko ek single DataFrame mein merge karna
            final_df = pd.concat(processed_data_list, ignore_index=True)

            st.write("### Converted Data ka Preview:")
            st.dataframe(final_df.head(10))

            # Merged DataFrame ko Excel format mein convert karna memory mein
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Converted_Data")

            excel_data = output_buffer.getvalue()

            # Final Download Button
            st.download_button(
                label="📥 Converted Excel File Download Karein",
                data=excel_data,
                file_name="Final_Converted_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
