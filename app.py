import streamlit as st
import pandas as pd

# Page Layout
st.set_page_config(page_title="Inward & Outward Dispatch System", layout="wide")

st.title("📊 Inward & Outward Dispatch Tracking System")
st.write("Jindal Saw Logistics & Dispatch Management Report")

# Load Excel Data
@st.cache_data
def load_data():
    file_path = "Inward_Outward_FIXED_and_FAST.xlsx"
    daily_dispatch = pd.read_excel(file_path, sheet_name="Daily Dispatch Report (FIXED)", skiprows=4)
    plant_summary = pd.read_excel(file_path, sheet_name="Plant-wise Summary", skiprows=2)
    product_summary = pd.read_excel(file_path, sheet_name="Product-wise Summary", skiprows=2)
    inward_reg = pd.read_excel(file_path, sheet_name="Inward Register", skiprows=2)
    outward_reg = pd.read_excel(file_path, sheet_name="Outward Register", skiprows=2)
    return daily_dispatch, plant_summary, product_summary, inward_reg, outward_reg

try:
    daily_dispatch, plant_summary, product_summary, inward_reg, outward_reg = load_data()

    # Sidebar Navigation
    st.sidebar.header("📌 Navigation")
    report_type = st.sidebar.radio(
        "Report Select Karein:",
        ["Daily Dispatch Report", "Plant-wise Summary", "Product-wise Summary", "Inward Register", "Outward Register"]
    )

    if report_type == "Daily Dispatch Report":
        st.subheader("📋 Daily Dispatch Report (FIXED)")
        st.write(f"Total Rows: {len(daily_dispatch)}")
        st.dataframe(daily_dispatch, use_container_width=True)

    elif report_type == "Plant-wise Summary":
        st.subheader("🏭 Plant-wise Summary")
        st.dataframe(plant_summary, use_container_width=True)

    elif report_type == "Product-wise Summary":
        st.subheader("📦 Product-wise Summary")
        st.dataframe(product_summary, use_container_width=True)

    elif report_type == "Inward Register":
        st.subheader("📥 Inward Register")
        st.dataframe(inward_reg, use_container_width=True)

    elif report_type == "Outward Register":
        st.subheader("📤 Outward Register")
        st.dataframe(outward_reg, use_container_width=True)

except Exception as e:
    st.error(f"Error: {e}")
    st.info("Kripya check karein ki Excel file ka naam exact `Inward_Outward_FIXED_and_FAST.xlsx` hona chahiye.")
