import os
import requests
import pandas as pd
import win32com.client as win32
from dotenv import load_dotenv
import numpy as np
from datetime import datetime

# Load environment variables from .env file
load_dotenv()

# Alpha Vantage API Key
API_KEY = os.getenv('ALPHA_VANTAGE_API_KEY')
if not API_KEY:
    raise ValueError("API Key not found. Please set ALPHA_VANTAGE_API_KEY in the .env file.")

# Get today's date and format it for the filename
today_date = datetime.today().strftime('%m-%d-%y')

# Set the path to save the file (use an absolute path)
save_dir = r'C:\Users\big_j\PycharmProjects\FX-Trading-and-Risk-Management-Tools'
excel_file_path = os.path.join(save_dir, f'FX_Trading_Report_{today_date}.xlsx')
macro_enabled_excel_file_path = os.path.join(save_dir, f'FX_Trading_Report_macro_{today_date}.xlsm')

# Function to fetch FX data from Alpha Vantage
def fetch_fx_data(from_currency, to_currency):
    url = f'https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency={from_currency}&to_currency={to_currency}&apikey={API_KEY}'
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        exchange_rate = float(data['Realtime Currency Exchange Rate']['5. Exchange Rate'])
        return exchange_rate
    except (requests.RequestException, KeyError, ValueError) as e:
        print(f"Error fetching data for {from_currency}/{to_currency}: {e}")
        return None

# P&L Calculation function
def calculate_pnl(trades):
    market_rate_fluctuation = np.random.uniform(-0.02, 0.02, len(trades))
    trades['Market_Rate'] = trades['FX_Rate'] * (1 + market_rate_fluctuation)
    trades['PnL'] = (trades['Market_Rate'] - trades['FX_Rate']) * trades['Notional']
    return trades

# Generate FX trade report with real-time data
def generate_fx_report(trades):
    trades.to_excel(excel_file_path, index=False)
    print(f"FX Trading Report with real-time data generated successfully at {excel_file_path}.")

# Automate Excel to insert VBA and run the macro
def automate_excel():
    vba_code = """
    Sub CalculateFXPricing()
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row

        ' Add a new column header for FX Pricing
        Cells(1, 8).Value = "FX_Pricing"

        Dim i As Integer
        For i = 2 To lastRow
            ' Basic pricing model: Notional * FX Rate
            Cells(i, 8).Value = Cells(i, 3).Value * Cells(i, 4).Value
        Next i

        MsgBox "FX Pricing calculated!"
    End Sub
    """
    try:
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Open(excel_file_path)
        vb_module = workbook.VBProject.VBComponents.Add(1)
        vb_module.CodeModule.AddFromString(vba_code)
        workbook.SaveAs(macro_enabled_excel_file_path, FileFormat=52)
        excel_app.Application.Run("CalculateFXPricing")
        workbook.Save()
        print(f"VBA macro added and executed successfully at {macro_enabled_excel_file_path}!")
    except Exception as e:
        print(f"Error automating Excel: {e}")
    finally:
        workbook.Close(SaveChanges=True)
        excel_app.Quit()

# Create a dataframe for FX trades and pull real-time exchange rates
fx_trades = pd.DataFrame({
    'Trade_ID': ['T001', 'T002', 'T003'],
    'Currency_Pair': ['EUR/USD', 'GBP/USD', 'USD/JPY'],
    'Notional': [1000000, 1500000, 1200000],
    'FX_Rate': [fetch_fx_data('EUR', 'USD'),
                fetch_fx_data('GBP', 'USD'),
                fetch_fx_data('USD', 'JPY')],
    'Trade_Type': ['Forward', 'Spot', 'NDF']
})

# Filter out trades with None FX_Rate
fx_trades = fx_trades.dropna(subset=['FX_Rate'])

# Calculate P&L for each trade
fx_trades = calculate_pnl(fx_trades)

# Generate the FX report
generate_fx_report(fx_trades)

# Automate Excel to insert VBA and run the macro
automate_excel()