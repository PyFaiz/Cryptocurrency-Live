import requests
import time
import pandas as pd
from datetime import datetime
import os

def fetch_top_cryptocurrencies(api_key, limit=50):
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        'vs_currency': 'usd',
        'order': 'market_cap_desc',
        'per_page': limit,
        'page': 1,
        'sparkline': False,
        'price_change_percentage': '24h',
        'x_cg_demo_api_key': api_key
    }
    
    try: 
        response = requests.get(url, params=params)
        if response.status_code == 200:
            data = response.json()
            crypto_data = []
            for crypto in data:
                crypto_info = {
                    'name': crypto['name'],
                    'symbol': crypto['symbol'].upper(),
                    'current_price': crypto['current_price'],
                    'market_cap': crypto['market_cap'],
                    'trading_volume_24h': crypto['total_volume'],
                    'price_change_24h': crypto['price_change_percentage_24h'],
                    'last_updated': pd.to_datetime(crypto['last_updated']).strftime('%Y-%m-%d %H:%M:%S'),
                    'data_refresh_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                crypto_data.append(crypto_info)
            return pd.DataFrame(crypto_data)
        
        elif response.status_code == 429:
            print("Rate limit exceeded. Waiting before retrying...")
            time.sleep(120)
            return fetch_top_cryptocurrencies(api_key, limit)
        
        else:
            print(f"API Error: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None
def analyze_crypto_data(df):

    if df is None or df.empty:
        print("No data to analyze.")
        return None
    
    analysis_results = {}

    top_5_by_market_cap = df.head(5)[['name', 'symbol', 'market_cap', 'current_price']]
    analysis_results['top_5_by_market_cap'] = top_5_by_market_cap
    

    avg_price = df['current_price'].mean()
    analysis_results['average_price'] = avg_price
    
    price_change_df = df.dropna(subset=['price_change_24h'])
    
    if not price_change_df.empty:
        max_price_change = price_change_df.loc[price_change_df['price_change_24h'].idxmax()]
        min_price_change = price_change_df.loc[price_change_df['price_change_24h'].idxmin()]
        
        analysis_results['highest_price_change'] = {
            'name': max_price_change['name'],
            'symbol': max_price_change['symbol'],
            'price_change_24h': max_price_change['price_change_24h']
        }
        
        analysis_results['lowest_price_change'] = {
            'name': min_price_change['name'],
            'symbol': min_price_change['symbol'],
            'price_change_24h': min_price_change['price_change_24h']
        }
    
    return analysis_results

def display_analysis_results(analysis_results):

    if analysis_results is None:
        print("No analysis results to display.")
        return
    
    print("\n" + "="*80)
    print("CRYPTOCURRENCY MARKET ANALYSIS")
    print("="*80)
    
    print("\nTOP 5 CRYPTOCURRENCIES BY MARKET CAP:")
    top_5_df = analysis_results['top_5_by_market_cap'].copy()
    for i, row in top_5_df.iterrows():
        print(f"{i+1}. {row['name']} ({row['symbol']}): ${row['market_cap']:,.0f}")
    
    print(f"\nAVERAGE PRICE OF TOP 50 CRYPTOCURRENCIES: ${analysis_results['average_price']:,.2f}")

    if 'highest_price_change' in analysis_results:
        highest = analysis_results['highest_price_change']
        lowest = analysis_results['lowest_price_change']
        
        print("\nPRICE CHANGE EXTREMES (24H):")
        print(f"Highest Increase: {highest['name']} ({highest['symbol']}): {highest['price_change_24h']:.2f}%")
        print(f"Highest Decrease: {lowest['name']} ({lowest['symbol']}): {lowest['price_change_24h']:.2f}%")

def display_crypto_data(df):

    if df is None or df.empty:
        print("No data to display.")
        return
    
    print(f"\n{'=' * 100}")
    print(f"TOP {len(df)} CRYPTOCURRENCIES BY MARKET CAPITALIZATION")
    print(f"Data fetched at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'=' * 100}")

    display_df = df.copy()

    display_df['current_price'] = display_df['current_price'].apply(lambda x: f"${x:,.2f}" if x >= 1 else f"${x:.8f}")
    display_df['market_cap'] = display_df['market_cap'].apply(lambda x: f"${x:,.0f}")
    display_df['trading_volume_24h'] = display_df['trading_volume_24h'].apply(lambda x: f"${x:,.0f}")
    display_df['price_change_24h'] = display_df['price_change_24h'].apply(lambda x: f"{x:.2f}%" if x else "N/A")

    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', 1000)
    print(display_df.to_string(index=False))

def update_live_excel(new_df, filename='live_crypto_data.xlsx'):
    try:
        # Always overwrite the existing file with fresh data
        new_df.to_excel(filename, index=False)
        print(f"Successfully updated {filename}")
    except Exception as e:
        print(f"Error updating Excel file: {str(e)}")

def main_loop(api_key, interval=300):
    while True:
        print(f"\n{'=' * 40} New Update {'=' * 40}")
        start_time = time.time()
        
        crypto_df = fetch_top_cryptocurrencies(api_key)
        if crypto_df is not None:
            update_live_excel(crypto_df)
            display_crypto_data(crypto_df)
            analysis = analyze_crypto_data(crypto_df)
            display_analysis_results(analysis)
        
        elapsed = time.time() - start_time
        sleep_time = max(interval - elapsed, 0)
        print(f"\nNext update in {sleep_time/60:.1f} minutes...")
        time.sleep(sleep_time)

if __name__ == "__main__":
    API_KEY = "CG-AoDLSTnUqmSFJTV1hp8MFTYZ"
    
    print("Starting real-time cryptocurrency monitor...")
    print("Live data will be OVERWRITTEN in 'live_crypto_data.xlsx' every 5 minutes")
    time.sleep(2)
    
    try:
        main_loop(API_KEY, interval=300)  
    except KeyboardInterrupt:
        print("\nMonitoring stopped by user. Final snapshot saved in 'live_crypto_data.xlsx'")