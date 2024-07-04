import pandas as pd
import xlsxwriter
import numpy as np
import yfinance as yf
import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

df = pd.read_excel("/Users/rohanraval/Desktop/PlayingAround/Book.xlsx")
df.rename(columns = {"Unnamed: 0" : "Stock"}, inplace=True)
df = df.get(["Stock", "Symbols", "Purchase Price"])

def current_stock_price(symbol):
    try:
        company = yf.Ticker(f"{symbol}.NS")
        closing_price = company.history(period = "1d")["Close"][0]
    except:
        company = yf.Ticker(f"{symbol.upper()}.BO")
        closing_price = company.history(period = "1d")["Close"][0]
    return round(closing_price, 2)

def stop_loss(purchase_price):
    return round(purchase_price - (purchase_price*0.1), 2)

df = df.assign(Stock = df.get("Stock").apply(lambda stock: stock.title()))
df = df.assign(Symbols = df.get("Symbols").apply(lambda symbol: symbol.upper()))
df = df.assign(LivePrice = df.get("Symbols").apply(current_stock_price))
df = df.assign(StopLoss = df.get("Purchase Price").apply(stop_loss))

change=[]
for i in range(df.shape[0]):
    chng = round(((df.iloc[i].get("LivePrice") - df.iloc[i].get("Purchase Price")) / df.iloc[i].get("Purchase Price")) * 100, 2)
    change.append(f"{chng}%")

df = df.assign(Change = change)
df = df.reindex(columns=["Stock", "Symbols", "Purchase Price", "LivePrice", "Change", "StopLoss"])

file_path = "/Users/rohanraval/Desktop/PlayingAround/Stocks.xlsx"

df.to_excel(file_path, index=False)

df = pd.read_excel(file_path)

output_file = file_path
workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet()

rupee_format = workbook.add_format({'num_format': '₹#,##0.00'})
green_font = workbook.add_format({'font_color': '#355E3B'})  # Hex code for green
red_font = workbook.add_format({'font_color': '#FF0000'})
bold_format = workbook.add_format({'bold': True})

# Write headers to the new workbook
for c, header in enumerate(df.columns):
    worksheet.write(0, c, header, bold_format)

for r, row in df.iterrows():
    for c, value in enumerate(row):
        if df.columns[c] == 'Purchase Price' or df.columns[c] == 'LivePrice':
            worksheet.write(r + 1, c, value, rupee_format)  # Apply Rupee format to specific column
        elif df.columns[c] == 'Change':
            if value[0] != "-":
                worksheet.write(r + 1, c, value, green_font)    # Apply green font color to specific column
            else:
                worksheet.write(r + 1, c, value, red_font)    # Apply green font color to specific column
        else:
            worksheet.write(r + 1, c, value)

workbook.close()
