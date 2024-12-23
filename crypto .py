import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import time

# Function to fetch live cryptocurrency data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")
        return []
    except ValueError:
        print("Error: Response is not in JSON format")
        return []

# Function to analyze data
def analyze_data(data):
    df = pd.DataFrame(data)
    df = df[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]

    # Analysis
    top_5_by_market_cap = df.nlargest(5, "market_cap")[["name", "market_cap"]]
    avg_price = df["current_price"].mean()
    highest_change = df.nlargest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]
    lowest_change = df.nsmallest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]

    analysis = {
        "Top 5 by Market Cap": top_5_by_market_cap,
        "Average Price": avg_price,
        "Highest Change (24h)": highest_change,
        "Lowest Change (24h)": lowest_change
    }
    return df, analysis

# Function to update Excel sheet
def update_excel(df):
    file_name = "Live_Crypto_Data.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Crypto Data"

    # Write data to sheet
    for row in dataframe_to_rows(df, index=False, header=True):
        sheet.append(row)

    # Save the workbook
    workbook.save(file_name)
    print(f"Excel sheet updated: {file_name}")

# Main loop to fetch, analyze, and update data
def main():
    while True:
        print("Fetching live data...")
        data = fetch_crypto_data()

        if data:
            df, analysis = analyze_data(data)

            print("\nAnalysis:")
            for key, value in analysis.items():
                print(f"{key}:")
                print(value)

            update_excel(df)
            print("Excel updated. Waiting for 5 minutes...")
        else:
            print("Failed to fetch data. Retrying in 5 minutes...")

        time.sleep(600)

if __name__ == "__main__":
    main()