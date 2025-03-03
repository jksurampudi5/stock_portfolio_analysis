import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import io
import msoffcrypto

# Set page configuration (Must be the first Streamlit command)
st.set_page_config(page_title="Stock Portfolio Dashboard", layout="wide")

# Streamlit App Title
st.title("\U0001F4C8 Stock Portfolio Dashboard")

# File Upload
uploaded_file = st.file_uploader("Upload your stock file (CSV or XLSX)", type=["csv", "xlsx"])

def decrypt_excel(file, password):
    decrypted = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(file)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted)
    return decrypted

if uploaded_file is not None:
    @st.cache_data
    def load_data(file, password=None):
        if file.name.endswith(".csv"):
            df = pd.read_csv(file, header=7)
        elif file.name.endswith(".xlsx"):
            try:
                xls = pd.ExcelFile(file, engine="openpyxl")
            except Exception as e:
                if "Workbook is encrypted" in str(e):
                    password = st.text_input("This file is password-protected. Enter the password:", type="password")
                    if password:
                        try:
                            decrypted_file = decrypt_excel(file, password)
                            xls = pd.ExcelFile(decrypted_file, engine="openpyxl")
                        except Exception:
                            st.error("Incorrect password or file could not be decrypted.")
                            return None
                    else:
                        return None
                else:
                    st.error(f"Error reading Excel file: {e}")
                    return None
            
            if "Equity" in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name="Equity", header=7)
            else:
                st.error("The uploaded Excel file does not contain a sheet named 'Equity'.")
                return None
        else:
            st.error("Unsupported file format. Please upload a CSV or XLSX file.")
            return None
        
        df.columns = df.columns.str.strip()  # Clean column names
        return df

    df = load_data(uploaded_file)

    if df is not None:
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        st.download_button(
            label="\U0001F4E5 Download as CSV",
            data=csv_buffer.getvalue(),
            file_name="converted_equity.csv",
            mime="text/csv",
        )

        required_columns = ["Company Name", "Total Quantity", "Avg Trading Price", "LTP"]
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns. Please ensure your file contains: {', '.join(required_columns)}")
        else:
            df[["Total Quantity", "Avg Trading Price", "LTP"]] = df[["Total Quantity", "Avg Trading Price", "LTP"]].apply(pd.to_numeric, errors='coerce')
            df = df[df["LTP"] > 0]
            df["Profit/Loss"] = (df["LTP"] - df["Avg Trading Price"]) * df["Total Quantity"]

            st.subheader("\U0001F4C9 Profit & Loss Breakdown")
            fig = px.bar(
                df.sort_values("Profit/Loss", ascending=True),
                x="Profit/Loss",
                y="Company Name",
                color="Profit/Loss",
                color_continuous_scale=["red", "yellow", "green"],
                title="Stock Portfolio - Profit/Loss Breakdown",
                labels={"Profit/Loss": "Total Profit/Loss", "Company Name": "Stock"},
                orientation="h",
                text="Profit/Loss"
            )
            fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            fig.update_layout(xaxis_title="Profit/Loss Amount", yaxis_title="Stock", height=600)
            st.plotly_chart(fig, use_container_width=True)
