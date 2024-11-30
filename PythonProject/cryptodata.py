import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import time


def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching data: {response.status_code}")
        return []
def extract_crypto_data(data):
    crypto_list = []
    for coin in data:
        crypto_list.append({
            "Cryptocurrency Name": coin["name"],
            "Symbol": coin["symbol"],
            "Current Price (USD)": coin["current_price"],
            "Market Capitalization": coin["market_cap"],
            "24-hour Trading Volume": coin["total_volume"],
            "Price Change (24h %)": coin["price_change_percentage_24h"]
        })
    return crypto_list
def analyze_crypto_data(df):
    top_5_by_market_cap = df.nlargest(5, "Market Capitalization")
    average_price = df["Current Price (USD)"].mean()

    highest_change = df["Price Change (24h %)"].max()
    lowest_change = df["Price Change (24h %)"].min()
    
    return {
        "Top 5 by Market Cap": top_5_by_market_cap,
        "Average Price": average_price,
        "Highest 24h Change": highest_change,
        "Lowest 24h Change": lowest_change
    }

def save_to_excel(data, file_name):
    df = pd.DataFrame(data)
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Cryptocurrency Data")
        analysis = analyze_crypto_data(df)
        analysis_sheet = writer.book.create_sheet("Analysis")
        analysis_sheet.append(["Metric", "Value"])
        analysis_sheet.append(["Average Price", analysis["Average Price"]])
        analysis_sheet.append(["Highest 24h Change", analysis["Highest 24h Change"]])
        analysis_sheet.append(["Lowest 24h Change", analysis["Lowest 24h Change"]])
        
        top_5_sheet = writer.book.create_sheet("Top 5 by Market Cap")
        for r in dataframe_to_rows(analysis["Top 5 by Market Cap"], index=False, header=True):
            top_5_sheet.append(r)

def run_live_updates():
    while True:
        file_name = "crypto_live_data.xlsx"
        
        data = fetch_crypto_data()
        if data:  
            extracted_data = extract_crypto_data(data)
            save_to_excel(extracted_data, file_name)
            print(f"Data saved to {file_name} at {time.strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            print("Failed to fetch data, retrying in 5 minutes...")
        time.sleep(300)

if __name__ == "__main__":
    run_live_updates()