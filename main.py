import yfinance as yf
import pandas as pd
import sys

# Check if the ticker symbol is provided via command line arguments
if len(sys.argv) != 2:
    print("Usage: python main.py <ticker>")
    sys.exit(1)

# Fetch the ticker symbol from command line arguments
ticker = sys.argv[1]

# Download the company's data using yfinance
company = yf.Ticker(ticker)

# Fetch the financial statements and information
info = pd.DataFrame(company.info.items())
income_statement = company.financials.T  # Transposed for better readability
balance_sheet = company.balance_sheet.T  # Transposed for better readability
cash_flow = company.cashflow.T  # Transposed for better readability

# Create a Pandas Excel writer using openpyxl as the engine
with pd.ExcelWriter(f'{ticker}_financials.xlsx', engine='openpyxl') as writer:
    # Write each DataFrame to a specific sheet
    info.to_excel(writer, sheet_name='Info', index=False)
    income_statement.to_excel(writer, sheet_name='Income Statement')
    balance_sheet.to_excel(writer, sheet_name='Balance Sheet')
    cash_flow.to_excel(writer, sheet_name='Cash Flow Statement')

# Print summary of key statistics available in the "info" DataFrame to stdout
print("Key Information Summary:")
for index, row in info.iterrows():
    print(f"{row[0]}: {row[1]}")

print(f"\nFinancial statements and company info have been exported to {ticker}_financials.xlsx.")
