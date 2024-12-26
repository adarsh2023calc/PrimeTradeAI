import requests
import numpy as np
import pandas as pd
from time import sleep
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from googleapiclient.discovery import build

def extract_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 250,  
        "page": 1,       
        "sparkline": False
    }

    crypto_list = []  

    try:
        print("** Fetching data from CoinGeckoAPI**")
        for page in range(1, 4):  # Pages 1 to 3
            params["page"] = page
            response = requests.get(url, params=params)
            response.raise_for_status()  # Raise an error for bad HTTP responses
            data = response.json()
            # Extract required fields
            for coin in data:
                crypto_list.append({
                    "Name": coin["name"],
                    "Symbol": coin["symbol"].upper(),
                    "Price (USD)": coin["current_price"],
                    "Market Cap": coin["market_cap"],
                    "Trading Volume (24h)": coin["total_volume"],
                    "Price Change (24h %)": coin["price_change_percentage_24h"],
                    "ATH (All-Time High)": coin["ath"],
                    "ATH Change (%)": coin["ath_change_percentage"],
                    "Low (24h)": coin["low_24h"],
                    "High (24h)": coin["high_24h"],
                    "Market Cap Rank": coin["market_cap_rank"],
                    "Total Supply": coin.get("total_supply", "N/A"),
                    "Circulating Supply": coin.get("circulating_supply", "N/A"),
                    "Max Supply": coin.get("max_supply", "N/A"),
                    "Last Updated": coin["last_updated"],
                })

        df = pd.DataFrame(crypto_list)
        
        print(f"Successfully fetched {len(df)} records.")
        print(df.head())  

        # Save the data to a JSON file
        with open("cryptocurrency_data.json", "w") as json_file:
            json.dump(crypto_list, json_file, indent=4)

        print("Data has been saved as cryptocurrency_data.json")

    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")
        df = pd.DataFrame(crypto_list)
        print(f"Successfully fetched {len(df)} records.")
        print(df.head())  
        with open("cryptocurrency_data.json", "w") as json_file:
            json.dump(crypto_list, json_file, indent=4)

def update_df_to_excel():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)

    # Create a new Google Sheet using the Google Drive API
    
    wb= client.open("Cryptocurrency")

    sheet = wb.get_worksheet(0)

    # Load data
    df = pd.read_json("cryptocurrency_data.json")

    # Clear and update data in the first worksheet
    sheet.clear()
    # Replace `inf` and `-inf` with NaN
    df.replace([np.inf, -np.inf], 0, inplace=True)

    # Fill NaN values with a placeholder or default value
    df.fillna(0, inplace=True)
    sheet.update([df.columns.values.tolist()] + df.values.tolist())

    # Create a second worksheet for analysis (if not already present)
    if len(wb.worksheets()) == 1:
        wb.add_worksheet(title="Analysis", rows="100", cols="20")
    sheet2 = wb.get_worksheet(1)

    # Perform analysis
    top_5 = df.nlargest(5, "Market Cap")
    average_price = df['Price (USD)'].mean()
    max_change = df.loc[df['Price Change (24h %)'].idxmax()]
    min_change = df.loc[df['Price Change (24h %)'].idxmin()]

    # Update top 5 cryptocurrencies in cells
    sheet2.update("A1", [["Top 5 Cryptocurrencies by Market Cap"]])  # Corrected
    sheet2.update("A2:O2", [top_5.columns.tolist()])  # Headers
    sheet2.update("A3:O7", top_5.values.tolist())  # Top 5 data

    # Update average price in a specific cell
    sheet2.update("A12", [["Average Price"]])
    sheet2.update("A13", [[f"${average_price:.2f}"]])

    # Update max change information
    sheet2.update("B12", [["Max 24H Price Change"]])
    sheet2.update("B12:C12", [["Name", "Price Change (%)"]])
    sheet2.update("B13:C13", [[max_change['Name'], max_change['Price Change (24h %)']]])

    # Update min change information
    sheet2.update("D12", [["Min 24H Price Change"]])
    sheet2.update("D12:E12", [["Name", "Price Change (%)"]])
    sheet2.update("D13:E13", [[min_change['Name'], min_change['Price Change (24h %)']]])

    print("Spreadsheet updated and saved.")
    print("Spreadsheet URL:", wb.url)  # This prints the URL of the spreadsheet

if __name__ == "__main__":
    while True:
        extract_data()
        update_df_to_excel()
        print("Sleeping for 5 minutes before the next update...")
        sleep(300)  # Sleep for 5 minutes
