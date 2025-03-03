import streamlit as st
import pandas as pd
import io
import msoffcrypto

# Set page configuration
st.set_page_config(page_title="Stock Portfolio Dashboard", layout="wide")

st.title("📈 Stock Portfolio Dashboard")

# File Upload
uploaded_file = st.file_uploader("Upload your stock file (CSV or XLSX)", type=["csv", "xlsx"])

# Ask for password only if an Excel file is uploaded
password = None
if uploaded_file is not None and uploaded_file.name.endswith(".xlsx"):
    password = st.text_input("This file may be password-protected. Enter password if required:", type="password")

@st.cache_data
def load_data(file, password):
    if file.name.endswith(".csv"):
        return pd.read_csv(file, header=7)

    elif file.name.endswith(".xlsx"):
        decrypted = io.BytesIO()
        file.seek(0)  # Reset file pointer

        try:
            # Try opening without a password
            df = pd.read_excel(file, sheet_name="Equity", header=7, engine="openpyxl")
            return df
        except Exception as e:
            if "File is not a zip file" in str(e) or "Workbook is encrypted" in str(e):
                if not password:
                    st.error("This file is encrypted. Please enter a password.")
                    return None
                
                try:
                    file.seek(0)
                    office_file = msoffcrypto.OfficeFile(file)
                    office_file.load_key(password=password)
                    office_file.decrypt(decrypted)

                    df = pd.read_excel(decrypted, sheet_name="Equity", header=7, engine="openpyxl")
                    return df
                except Exception as e:
                    st.error("Incorrect password or file could not be decrypted.")
                    return None
            else:
                st.error(f"Error reading Excel file: {e}")
                return None

# Load data if a file is uploaded
if uploaded_file is not None:
    df = load_data(uploaded_file, password)

    if df is not None:
        st.success("✅ File loaded successfully!")
        st.dataframe(df)
