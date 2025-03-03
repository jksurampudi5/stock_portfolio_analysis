import streamlit as st
import pandas as pd
import io
import msoffcrypto

st.set_page_config(page_title="Stock Portfolio Dashboard", layout="wide")
st.title("📈 Stock Portfolio Dashboard")
uploaded_file = st.file_uploader("Upload your stock file (CSV or XLSX)", type=["csv", "xlsx"])
if uploaded_file is not None:
    @st.cache_data
    def load_data(file):
        if file.name.endswith(".csv"):
            return pd.read_csv(file, header=7)
        elif file.name.endswith(".xlsx"):
            decrypted = io.BytesIO()
            file.seek(0)  # Reset file pointer

            # Try opening the Excel file without a password
            try:
                df = pd.read_excel(file, sheet_name="Equity", header=7, engine="openpyxl")
                return df  # If successful, return data

            except Exception as e:
                if "File is not a zip file" in str(e) or "Workbook is encrypted" in str(e):
                    st.warning("The file is password-protected. Please enter the password below.")
                    password = st.text_input("Password:", type="password")

                    if password:
                        try:
                            # Use msoffcrypto-tool to decrypt
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

    df = load_data(uploaded_file)

    if df is not None:
        st.write("✅ File loaded successfully!")
        st.dataframe(df)
