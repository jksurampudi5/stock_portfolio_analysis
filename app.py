import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import io

# Set page configuration (Must be the first Streamlit command)
st.set_page_config(page_title="Stock Portfolio Dashboard", layout="wide")



# Streamlit App Title
st.title("📈 Stock Portfolio Dashboard")

# Load custom CSS
def load_css(file_name):
    with open(file_name) as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)


# Apply custom CSS
load_css("style.css")


# File Upload
uploaded_file = st.file_uploader("Upload your stock file (CSV or XLSX)", type=["csv", "xlsx"])

if uploaded_file is not None:
    @st.cache_data
    def load_data(file, password=None):
        if file.name.endswith(".csv"):
            df = pd.read_csv(file, header=7)
        elif file.name.endswith(".xlsx"):
            try:
                # Try opening the file without a password
                xls = pd.ExcelFile(file, engine="openpyxl")
                if "Equity" in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name="Equity", header=7)
                else:
                    st.error("The uploaded Excel file does not contain a sheet named 'Equity'.")
                    return None
            except Exception as e:
                if "Workbook is encrypted" in str(e):
                    password = st.text_input("This file is password-protected. Enter the password:", type="password")
                    if password:
                        try:
                            xls = pd.ExcelFile(file, engine="openpyxl", password=password)
                            df = pd.read_excel(xls, sheet_name="Equity", header=7)
                        except Exception as e:
                            st.error("Incorrect password or file could not be decrypted.")
                            return None
                    else:
                        return None
                else:
                    st.error(f"Error reading Excel file: {e}")
                    return None
        else:
            st.error("Unsupported file format. Please upload a CSV or XLSX file.")
            return None
        
        df.columns = df.columns.str.strip()  # Clean column names
        return df

    df = load_data(uploaded_file)

    if df is not None:
        # Provide option to download XLSX as CSV
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()

        st.download_button(
            label="📥 Download as CSV",
            data=csv_data,
            file_name="converted_equity.csv",
            mime="text/csv",
        )

        # Identify relevant columns
        required_columns = ["Company Name", "Total Quantity", "Avg Trading Price", "LTP"]

        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns. Please ensure your file contains: {', '.join(required_columns)}")
        else:
            stock_column = "Company Name"
            quantity_column = "Total Quantity"
            purchase_price_column = "Avg Trading Price"
            ltp_column = "LTP"  # Last Traded Price (Current Market Price)

            # Convert necessary columns to numeric
            df[quantity_column] = pd.to_numeric(df[quantity_column], errors='coerce')
            df[purchase_price_column] = pd.to_numeric(df[purchase_price_column], errors='coerce')
            df[ltp_column] = pd.to_numeric(df[ltp_column], errors='coerce')

            # Remove invalid stocks where LTP is zero
            df = df[df[ltp_column] > 0]

            # Calculate Profit/Loss
            df["Profit/Loss"] = (df[ltp_column] - df[purchase_price_column]) * df[quantity_column]

            # Sidebar: Select a Stock
            st.sidebar.title("📌 Stock Selector")
            selected_stock = st.sidebar.selectbox("Select a Stock:", df[stock_column].unique())

            # Filter data for the selected stock
            stock_data = df[df[stock_column] == selected_stock]

            # Display Stock Details
            st.subheader(f"📊 {selected_stock} Stock Details")

            # Show key financial details with conditional formatting
            def highlight_profit_loss(val):
                color = 'red' if val < 0 else 'green' if val > 0 else 'black'
                return f'color: {color}'

            styled_df = stock_data[[stock_column, quantity_column, purchase_price_column, ltp_column, "Profit/Loss"]].style.applymap(
                highlight_profit_loss, subset=["Profit/Loss"]
            )
            st.dataframe(styled_df)

            # Display Current Market Price
            st.subheader("📌 Current Market Price")
            if not stock_data.empty:
                st.write(f"The latest market price for **{selected_stock}** is **{stock_data[ltp_column].iloc[-1]:,.2f}**")
            else:
                st.write("No data available for the selected stock.")

            # Separate profit and loss stocks
            profit_stocks = df[df["Profit/Loss"] > 0]
            loss_stocks = df[df["Profit/Loss"] < 0]

            # Layout for Profit & Loss Summary
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📈 Profit Stocks")
                st.dataframe(profit_stocks[[stock_column, "Profit/Loss", quantity_column]])
                st.write(f"**Total Profit: {profit_stocks['Profit/Loss'].sum():,.2f}**")

            with col2:
                st.subheader("📉 Loss Stocks")
                st.dataframe(loss_stocks[[stock_column, "Profit/Loss", quantity_column]])
                st.write(f"**Total Loss: {loss_stocks['Profit/Loss'].sum():,.2f}**")

            # Improved Interactive Chart for Profit/Loss Breakdown
            st.subheader("📊 Profit & Loss Breakdown")

            fig = px.bar(
                df.sort_values("Profit/Loss", ascending=True), 
                x="Profit/Loss", 
                y=stock_column, 
                color="Profit/Loss",
                color_continuous_scale=["red", "yellow", "green"],
                title="Stock Portfolio - Profit/Loss Breakdown",
                labels={"Profit/Loss": "Total Profit/Loss", stock_column: "Stock"},
                orientation="h",  # Horizontal bar chart
                text="Profit/Loss"
            )

            fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            fig.update_layout(xaxis_title="Profit/Loss Amount", yaxis_title="Stock", height=600)

            st.plotly_chart(fig, use_container_width=True)
