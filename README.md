# FX Trading and Risk Management Tools

This project is a Python-based automation tool designed to generate real-time **FX trade reports**, calculate **P&L (Profit & Loss)**, and automate pricing calculations using **Excel VBA macros**. The tool integrates with the **Alpha Vantage API** to fetch real-time foreign exchange (FX) rates and leverages Excel for automating pricing computations.

## Features
- **Real-time FX rate fetching**: Utilizes the Alpha Vantage API to pull real-time exchange rates for major currency pairs.
- **P&L Calculation**: Simulates market fluctuations to calculate the P&L for each trade.
- **Excel Report Generation**: Exports FX trade data, including calculated P&L, to an Excel file with the current date in the filename.
- **Excel VBA Automation**: Inserts a VBA macro into the generated Excel file to perform pricing calculations automatically.
- **Macro-Enabled Workbook**: Saves the Excel workbook as a macro-enabled `.xlsm` file for further use.

## Technologies and Libraries
- **Python** (Core script)
- **Alpha Vantage API** (Real-time FX data)
- **Pandas** (Data manipulation)
- **NumPy** (For P&L calculation)
- **win32com.client** (For Excel automation)
- **dotenv** (To securely store API keys)
- **Excel VBA** (Automated pricing calculations)

## Requirements
- Python 3.x
- Required Python libraries:
  - `pandas`
  - `numpy`
  - `requests`
  - `python-dotenv`
  - `pywin32`
- An API Key from [Alpha Vantage](https://www.alphavantage.co/support/#api-key) to fetch FX rates.
- Microsoft Excel installed (for automation via `win32com.client` and VBA).

## Setup

### 1. Clone the Repository:
```bash
git clone https://github.com/junelus/fx_trading_tool.git
cd fx_trading_tool
```

### 2. Install the Required Libraries from `requirements.txt`:
```bash
pip install -r requirements.txt
```

### 3. Set Up Your API Key:
Create a `.env` file in the project root directory with the following content:
```bash
ALPHA_VANTAGE_API_KEY=your_alpha_vantage_api_key
```
Replace `your_alpha_vantage_api_key` with your actual API key from Alpha Vantage.

### 4. Run the Script:
Simply run the script to generate the FX trade report, including P&L calculations and Excel pricing automation:
```bash
python fx_trading_tool.py
```

## How It Works

1. **Fetch Real-Time FX Data**: The script fetches real-time FX exchange rates for EUR/USD, GBP/USD, and USD/JPY using the Alpha Vantage API.
   
2. **P&L Calculation**: Based on the fetched FX rates and simulated market fluctuations, the script calculates the P&L for each trade using the formula:
   ```
   PnL = (Market Rate - FX Rate) * Notional
   ```
   
3. **Generate Excel Report**: The trade data, along with P&L and market rates, is exported to an Excel file named `FX_Trading_Report_MM-DD-YY.xlsx`.

4. **VBA Macro Insertion**: The script inserts a VBA macro into the Excel file to perform FX pricing calculations. This macro calculates:
   ```
   Pricing = Notional * FX Rate
   ```
   The macro runs automatically when the script completes execution.

5. **Save as Macro-Enabled Workbook**: The Excel report is saved as a macro-enabled workbook (`.xlsm`).

## Output
- **FX_Trading_Report_MM-DD-YY.xlsx**: The main report containing the fetched FX rates, calculated P&L, and market rates.
- **FX_Trading_Report_macro_MM-DD-YY.xlsm**: A macro-enabled workbook with embedded VBA code for further pricing automation in Excel.

## Example Workflow

1. The script starts by pulling real-time FX rates.
2. It calculates market fluctuations and P&L for each trade.
3. It generates an Excel report and injects a VBA macro to perform pricing calculations automatically.
4. The output files are saved in your specified directory.

### Sample Excel Output:
| Trade_ID | Currency_Pair | Notional  | FX_Rate | Trade_Type | Market_Rate | PnL         |
|----------|---------------|-----------|---------|------------|-------------|-------------|
| T001     | EUR/USD       | 1,000,000 | 1.12    | Forward    | 1.13        | 10,000.00   |
| T002     | GBP/USD       | 1,500,000 | 1.35    | Spot       | 1.36        | 15,000.00   |
| T003     | USD/JPY       | 1,200,000 | 109.5   | NDF        | 110.1       | 72,000.00   |

## Notes
- Ensure you have Microsoft Excel installed on your machine since the script relies on Excel for report generation and VBA macro execution.
- The script can be extended to include more currency pairs by updating the `fx_trades` DataFrame and API calls accordingly.

## Future Improvements
- **Additional Currencies**: Include more currency pairs for a wider range of FX trading analysis.
- **Improved Market Data**: Integrate with a more comprehensive market data provider to simulate realistic market fluctuations for P&L calculations.
- **Enhanced Reporting**: Add more detailed reporting features, such as visualizations and trend analysis.

## Use Case:
- Designed for front-office developers working with FX trading desks, particularly for enhancing risk management, P&L reporting, and pricing tools.