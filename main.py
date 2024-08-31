import yfinance as yf
import pandas as pd

# Define the ticker symbol for Microsoft
ticker = "ESEA"

# Download the company's data using yfinance
msft = yf.Ticker(ticker)

# Fetch the financial statements
income_statement = msft.financials.T  # Transposed for better readability
balance_sheet = msft.balance_sheet.T  # Transposed for better readability
cash_flow = msft.cashflow.T  # Transposed for better readability

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter('financials.xlsx', engine='openpyxl') as writer:
    # Write each DataFrame to a specific sheet
    income_statement.to_excel(writer, sheet_name='Income Statement')
    balance_sheet.to_excel(writer, sheet_name='Balance Sheet')
    cash_flow.to_excel(writer, sheet_name='Cash Flow Statement')

print("Financial statements have been exported to financials.xlsx")
