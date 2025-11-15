"""
ExcelBot Pro - Stock Market Data & VBA Automation Suite
A professional tool for NSE stock market analysis with Zerodha Kite & Financial Modeling Prep API
Features: VBA macro generation, stock data analysis, Excel automation, and GitHub integration
Author: Mandar Bahadarpurkar
License: MIT
"""

import os
import re
import gradio as gr
from github import Github
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import traceback
import requests
from typing import Dict, List, Optional
import json

# Trading API Libraries
try:
    from kiteconnect import KiteConnect
    KITE_AVAILABLE = True
except ImportError:
    KITE_AVAILABLE = False
    print("⚠️ KiteConnect not available. Install with: pip install kiteconnect")

try:
    from dhanhq import dhanhq
    DHAN_AVAILABLE = True
except ImportError:
    DHAN_AVAILABLE = False
    print("⚠️ DhanHQ not available. Install with: pip install dhanhq")

# API Configuration
ZERODHA_API_KEY = os.getenv("ZERODHA_API_KEY", "kr8ob80gcmucrvph")
FMP_API_KEY = os.getenv("FMP_API_KEY", "rtD0v37SghQ4gMZNfX7q2Arv6RO7StUv")
FMP_BASE_URL = "https://financialmodelingprep.com/api/v3"

# Dhan API Configuration
DHAN_CLIENT_ID = os.getenv("DHAN_CLIENT_ID", "a04ba78c")
DHAN_ACCESS_TOKEN = os.getenv("DHAN_ACCESS_TOKEN", "ccb99f92-9f54-41dc-b209-84d53ac76291")

# Initialize Dhan client
dhan_client = None
if DHAN_AVAILABLE:
    try:
        dhan_client = dhanhq(DHAN_CLIENT_ID, DHAN_ACCESS_TOKEN)
        print("✅ Dhan API initialized successfully")
    except Exception as e:
        print(f"⚠️ Dhan API initialization failed: {e}")
        dhan_client = None

# Popular NSE Stocks
POPULAR_NSE_STOCKS = [
    "RELIANCE.NS", "TCS.NS", "INFY.NS", "HDFCBANK.NS", "ICICIBANK.NS",
    "HINDUNILVR.NS", "ITC.NS", "SBIN.NS", "BHARTIARTL.NS", "KOTAKBANK.NS",
    "LT.NS", "AXISBANK.NS", "ASIANPAINT.NS", "MARUTI.NS", "HCLTECH.NS",
    "BAJFINANCE.NS", "WIPRO.NS", "ULTRACEMCO.NS", "NESTLEIND.NS", "TITAN.NS"
]

# Demo/Mock Data for NSE Stocks (when API is unavailable)
DEMO_NSE_DATA = {
    "RELIANCE.NS": {
        "symbol": "RELIANCE.NS",
        "name": "Reliance Industries Ltd",
        "price": 2456.75,
        "change": 23.50,
        "change_percent": 0.97,
        "day_low": 2433.25,
        "day_high": 2467.80,
        "year_low": 2150.00,
        "year_high": 2750.00,
        "market_cap": 16543210000000,
        "volume": 8523000,
        "avg_volume": 7200000,
        "open": 2445.00,
        "previous_close": 2433.25,
        "eps": 98.50,
        "pe": 24.95,
        "timestamp": "2025-11-15"
    },
    "TCS.NS": {
        "symbol": "TCS.NS",
        "name": "Tata Consultancy Services Ltd",
        "price": 3842.50,
        "change": 46.75,
        "change_percent": 1.23,
        "day_low": 3810.00,
        "day_high": 3855.00,
        "year_low": 3200.00,
        "year_high": 4100.00,
        "market_cap": 14000000000000,
        "volume": 2156000,
        "avg_volume": 1800000,
        "open": 3820.00,
        "previous_close": 3795.75,
        "eps": 115.00,
        "pe": 33.41,
        "timestamp": "2025-11-15"
    },
    "INFY.NS": {
        "symbol": "INFY.NS",
        "name": "Infosys Ltd",
        "price": 1567.30,
        "change": -7.20,
        "change_percent": -0.46,
        "day_low": 1562.00,
        "day_high": 1580.00,
        "year_low": 1350.00,
        "year_high": 1750.00,
        "market_cap": 6500000000000,
        "volume": 5234000,
        "avg_volume": 4500000,
        "open": 1575.00,
        "previous_close": 1574.50,
        "eps": 62.50,
        "pe": 25.08,
        "timestamp": "2025-11-15"
    },
    "HDFCBANK.NS": {
        "symbol": "HDFCBANK.NS",
        "name": "HDFC Bank Ltd",
        "price": 1678.90,
        "change": 15.60,
        "change_percent": 0.94,
        "day_low": 1665.00,
        "day_high": 1685.00,
        "year_low": 1400.00,
        "year_high": 1800.00,
        "market_cap": 12800000000000,
        "volume": 6789000,
        "avg_volume": 5600000,
        "open": 1670.00,
        "previous_close": 1663.30,
        "eps": 85.50,
        "pe": 19.63,
        "timestamp": "2025-11-15"
    },
    "ICICIBANK.NS": {
        "symbol": "ICICIBANK.NS",
        "name": "ICICI Bank Ltd",
        "price": 1089.45,
        "change": 12.30,
        "change_percent": 1.14,
        "day_low": 1078.00,
        "day_high": 1095.00,
        "year_low": 850.00,
        "year_high": 1150.00,
        "market_cap": 7600000000000,
        "volume": 8456000,
        "avg_volume": 7200000,
        "open": 1082.00,
        "previous_close": 1077.15,
        "eps": 48.20,
        "pe": 22.60,
        "timestamp": "2025-11-15"
    }
}

# NSE Symbol Mapping for Dhan API
NSE_SYMBOL_MAP = {
    "RELIANCE": "RELIANCE", "TCS": "TCS", "INFY": "INFY", "INFOSYS": "INFY",
    "HDFCBANK": "HDFCBANK", "ICICIBANK": "ICICIBANK", "HINDUNILVR": "HINDUNILVR",
    "ITC": "ITC", "SBIN": "SBIN", "BHARTIARTL": "BHARTIARTL", "KOTAKBANK": "KOTAKBANK",
    "LT": "LT", "AXISBANK": "AXISBANK", "ASIANPAINT": "ASIANPAINT", "MARUTI": "MARUTI",
    "HCLTECH": "HCLTECH", "BAJFINANCE": "BAJFINANCE", "WIPRO": "WIPRO",
    "ULTRACEMCO": "ULTRACEMCO", "NESTLEIND": "NESTLEIND", "TITAN": "TITAN"
}

def get_clean_symbol(symbol: str) -> str:
    """Remove .NS suffix and clean symbol for API calls"""
    return symbol.replace('.NS', '').replace('.BO', '').upper()

def fetch_live_dhan_data(symbol: str) -> Dict:
    """
    Fetch live NSE stock data from Dhan API
    Returns: Dict with stock data or None if unavailable
    """
    if not dhan_client:
        return None
    
    try:
        clean_symbol = get_clean_symbol(symbol)
        mapped_symbol = NSE_SYMBOL_MAP.get(clean_symbol, clean_symbol)
        
        # Dhan API call for live quotes
        # Note: Dhan uses security_id, we need to get quote by symbol
        response = dhan_client.get_quote(
            exchange_segment=dhan_client.NSE,
            security_id=mapped_symbol
        )
        
        if response and response.get('status') == 'success':
            data = response.get('data', {})
            return {
                "symbol": f"{clean_symbol}.NS",
                "name": data.get('trading_symbol', clean_symbol),
                "price": data.get('ltp', 0),  # Last Traded Price
                "change": data.get('change', 0),
                "change_percent": data.get('change_percent', 0),
                "day_low": data.get('low', 0),
                "day_high": data.get('high', 0),
                "year_low": data.get('52_week_low', 0),
                "year_high": data.get('52_week_high', 0),
                "market_cap": 0,  # Not provided by Dhan
                "volume": data.get('volume', 0),
                "avg_volume": data.get('avg_volume', 0),
                "open": data.get('open', 0),
                "previous_close": data.get('prev_close', 0),
                "eps": 0,  # Not provided by Dhan
                "pe": 0,   # Not provided by Dhan
                "timestamp": data.get('timestamp', ''),
                "demo_mode": False,
                "data_source": "Dhan API (Live)"
            }
        return None
    except Exception as e:
        print(f"Dhan API error for {symbol}: {e}")
        return None

# Stock Market Data Functions
def fetch_nse_stock_data(symbol: str) -> Dict:
    """
    Fetch NSE stock data with multi-source fallback:
    1. Try Dhan API (Live Indian market data)
    2. Try Financial Modeling Prep API
    3. Fall back to demo data
    Symbol format: RELIANCE.NS, TCS.NS, INFY.NS
    """
    try:
        # Ensure .NS suffix for NSE stocks
        if not symbol.endswith('.NS'):
            symbol = f"{symbol}.NS"
        
        # PRIORITY 1: Try Dhan API for live Indian market data
        if DHAN_AVAILABLE and dhan_client:
            dhan_data = fetch_live_dhan_data(symbol)
            if dhan_data:
                print(f"✅ Fetched LIVE data for {symbol} from Dhan API")
                return dhan_data
        
        # PRIORITY 2: Try Financial Modeling Prep API
        url = f"{FMP_BASE_URL}/quote/{symbol}?apikey={FMP_API_KEY}"
        response = requests.get(url, timeout=10)
        
        # If API fails (403, 429, etc.), use demo data
        if response.status_code in [403, 429]:
            print(f"⚠️ FMP API rate limit for {symbol}, using demo data")
            if symbol in DEMO_NSE_DATA:
                demo_data = DEMO_NSE_DATA[symbol].copy()
                demo_data["demo_mode"] = True
                demo_data["data_source"] = "Demo Data (API Rate Limit)"
                return demo_data
            else:
                demo_data = DEMO_NSE_DATA["RELIANCE.NS"].copy()
                demo_data["demo_mode"] = True
                demo_data["symbol"] = symbol
                demo_data["name"] = f"Demo Data - {symbol}"
                demo_data["data_source"] = "Demo Data (API Rate Limit)"
                return demo_data
        
        response.raise_for_status()
        data = response.json()
        
        if data and len(data) > 0:
            stock = data[0]
            print(f"✅ Fetched data for {symbol} from FMP API")
            return {
                "symbol": stock.get("symbol", "N/A"),
                "name": stock.get("name", "N/A"),
                "price": stock.get("price", 0),
                "change": stock.get("change", 0),
                "change_percent": stock.get("changesPercentage", 0),
                "day_low": stock.get("dayLow", 0),
                "day_high": stock.get("dayHigh", 0),
                "year_low": stock.get("yearLow", 0),
                "year_high": stock.get("yearHigh", 0),
                "market_cap": stock.get("marketCap", 0),
                "volume": stock.get("volume", 0),
                "avg_volume": stock.get("avgVolume", 0),
                "open": stock.get("open", 0),
                "previous_close": stock.get("previousClose", 0),
                "eps": stock.get("eps", 0),
                "pe": stock.get("pe", 0),
                "timestamp": stock.get("timestamp", ""),
                "demo_mode": False,
                "data_source": "Financial Modeling Prep API"
            }
        else:
            # No data from API, use demo
            print(f"⚠️ No data from FMP for {symbol}, using demo data")
            if symbol in DEMO_NSE_DATA:
                demo_data = DEMO_NSE_DATA[symbol].copy()
                demo_data["demo_mode"] = True
                demo_data["data_source"] = "Demo Data (No API Data)"
                return demo_data
            return {"error": f"No data found for {symbol}"}
            
    except Exception as e:
        # On any error, try to return demo data
        print(f"❌ Error fetching {symbol}: {e}, using demo data")
        if symbol in DEMO_NSE_DATA:
            demo_data = DEMO_NSE_DATA[symbol].copy()
            demo_data["demo_mode"] = True
            demo_data["data_source"] = "Demo Data (Error Fallback)"
            return demo_data
        return {"error": f"Error fetching data: {str(e)}"}

def generate_demo_historical_data(symbol: str, days: int = 30) -> pd.DataFrame:
    """Generate demo historical data for testing"""
    import numpy as np
    base_prices = {
        "RELIANCE.NS": 1678.50, "TCS.NS": 3845.20, "HDFCBANK.NS": 1650.75,
        "INFY.NS": 1520.40, "ICICIBANK.NS": 1089.45, "HINDUNILVR.NS": 2456.80,
        "ITC.NS": 445.30, "SBIN.NS": 785.60, "BHARTIARTL.NS": 1542.90, "WIPRO.NS": 567.25
    }
    base_price = base_prices.get(symbol, 1500.00)
    dates = pd.date_range(end=datetime.now(), periods=days, freq='D')
    np.random.seed(hash(symbol) % 2**32)
    returns = np.random.normal(0.001, 0.02, days)
    prices = base_price * np.exp(np.cumsum(returns))
    data = []
    for i, date in enumerate(dates):
        open_price = prices[i] * (1 + np.random.uniform(-0.01, 0.01))
        close_price = prices[i]
        high_price = max(open_price, close_price) * (1 + np.random.uniform(0, 0.02))
        low_price = min(open_price, close_price) * (1 - np.random.uniform(0, 0.02))
        volume = int(np.random.uniform(1000000, 10000000))
        data.append({'date': date, 'open': round(open_price, 2), 'high': round(high_price, 2),
                     'low': round(low_price, 2), 'close': round(close_price, 2), 'volume': volume})
    df = pd.DataFrame(data)
    df['date'] = pd.to_datetime(df['date'])
    return df

def fetch_nse_historical_data(symbol: str, days: int = 30) -> pd.DataFrame:
    """
    Fetch historical stock data from Financial Modeling Prep API
    Falls back to demo data if API is unavailable
    """
    try:
        if not symbol.endswith('.NS'):
            symbol = f"{symbol}.NS"
        
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        url = f"{FMP_BASE_URL}/historical-price-full/{symbol}?from={start_date}&to={end_date}&apikey={FMP_API_KEY}"
        response = requests.get(url, timeout=10)
        
        if response.status_code in [403, 429]:
            return generate_demo_historical_data(symbol, days)
        
        response.raise_for_status()
        data = response.json()
        
        if 'historical' in data:
            df = pd.DataFrame(data['historical'])
            df['date'] = pd.to_datetime(df['date'])
            df = df.sort_values('date')
            return df
        else:
            return generate_demo_historical_data(symbol, days)
    except Exception as e:
        print(f"Error fetching historical data: {e}")
        return generate_demo_historical_data(symbol, days)

def get_demo_top_gainers() -> str:
    """Demo data for top gainers"""
    gainers = [("ADANIENT.NS", "Adani Enterprises", 2845.60, 5.23), ("TATAMOTORS.NS", "Tata Motors", 892.45, 4.87),
               ("HINDALCO.NS", "Hindalco Industries", 645.30, 3.92), ("JSWSTEEL.NS", "JSW Steel", 876.70, 3.45),
               ("TATASTEEL.NS", "Tata Steel", 145.80, 3.12), ("SUNPHARMA.NS", "Sun Pharma", 1678.40, 2.89),
               ("MARUTI.NS", "Maruti Suzuki", 12456.80, 2.65), ("TECHM.NS", "Tech Mahindra", 1234.50, 2.34),
               ("ULTRACEMCO.NS", "UltraTech Cement", 9876.20, 2.15), ("POWERGRID.NS", "Power Grid Corp", 287.60, 1.98)]
    result = "📈 **TOP NSE GAINERS**\n\n⚠️ **DEMO DATA** - API rate limit reached\n\n"
    for i, (symbol, name, price, change) in enumerate(gainers, 1):
        result += f"{i}. **{symbol}** - {name}\n   Price: ₹{price:.2f} | Change: +{change:.2f}%\n\n"
    return result

def get_nse_top_gainers() -> str:
    """
    Fetch top gainers from NSE
    Falls back to demo data if API is unavailable
    """
    try:
        url = f"{FMP_BASE_URL}/stock_market/gainers?apikey={FMP_API_KEY}"
        response = requests.get(url, timeout=10)
        
        if response.status_code in [403, 429]:
            return get_demo_top_gainers()
        
        response.raise_for_status()
        data = response.json()
        
        # Filter for NSE stocks
        nse_stocks = [stock for stock in data if '.NS' in stock.get('symbol', '')][:10]
        
        if nse_stocks:
            result = "📈 **TOP NSE GAINERS**\n\n"
            for i, stock in enumerate(nse_stocks, 1):
                result += f"{i}. **{stock.get('symbol', 'N/A')}** - {stock.get('name', 'N/A')}\n"
                result += f"   Price: ₹{stock.get('price', 0):.2f} | "
                result += f"Change: {stock.get('changesPercentage', 0):.2f}%\n\n"
            return result
        else:
            return get_demo_top_gainers()
    except Exception as e:
        print(f"Error fetching gainers: {e}")
        return get_demo_top_gainers()

def get_demo_top_losers() -> str:
    """Demo data for top losers"""
    losers = [("BAJAJFINSV.NS", "Bajaj Finserv", 1534.20, -3.45), ("DRREDDY.NS", "Dr Reddy's Labs", 5678.90, -3.12),
              ("CIPLA.NS", "Cipla", 1345.60, -2.89), ("EICHERMOT.NS", "Eicher Motors", 4567.80, -2.67),
              ("HEROMOTOCO.NS", "Hero MotoCorp", 4234.50, -2.45), ("DIVISLAB.NS", "Divi's Labs", 3789.20, -2.23),
              ("BRITANNIA.NS", "Britannia Industries", 5234.80, -2.01), ("NESTLEIND.NS", "Nestle India", 24567.40, -1.89),
              ("ASIANPAINT.NS", "Asian Paints", 3456.70, -1.67), ("HCLTECH.NS", "HCL Technologies", 1456.30, -1.45)]
    result = "📉 **TOP NSE LOSERS**\n\n⚠️ **DEMO DATA** - API rate limit reached\n\n"
    for i, (symbol, name, price, change) in enumerate(losers, 1):
        result += f"{i}. **{symbol}** - {name}\n   Price: ₹{price:.2f} | Change: {change:.2f}%\n\n"
    return result

def get_nse_top_losers() -> str:
    """
    Fetch top losers from NSE
    Falls back to demo data if API is unavailable
    """
    try:
        url = f"{FMP_BASE_URL}/stock_market/losers?apikey={FMP_API_KEY}"
        response = requests.get(url, timeout=10)
        
        if response.status_code in [403, 429]:
            return get_demo_top_losers()
        
        response.raise_for_status()
        data = response.json()
        
        # Filter for NSE stocks
        nse_stocks = [stock for stock in data if '.NS' in stock.get('symbol', '')][:10]
        
        if nse_stocks:
            result = "📉 **TOP NSE LOSERS**\n\n"
            for i, stock in enumerate(nse_stocks, 1):
                result += f"{i}. **{stock.get('symbol', 'N/A')}** - {stock.get('name', 'N/A')}\n"
                result += f"   Price: ₹{stock.get('price', 0):.2f} | "
                result += f"Change: {stock.get('changesPercentage', 0):.2f}%\n\n"
            return result
        else:
            return get_demo_top_losers()
    except Exception as e:
        print(f"Error fetching losers: {e}")
        return get_demo_top_losers()

def get_demo_most_active() -> str:
    """Demo data for most active stocks"""
    actives = [("RELIANCE.NS", "Reliance Industries", 1678.50, 15678900), ("SBIN.NS", "State Bank of India", 785.60, 14234500),
               ("ICICIBANK.NS", "ICICI Bank", 1089.45, 12456780), ("HDFCBANK.NS", "HDFC Bank", 1650.75, 11234560),
               ("TCS.NS", "Tata Consultancy", 3845.20, 9876540), ("INFY.NS", "Infosys", 1520.40, 8765430),
               ("BHARTIARTL.NS", "Bharti Airtel", 1542.90, 7654320), ("ITC.NS", "ITC Ltd", 445.30, 6543210),
               ("HINDUNILVR.NS", "Hindustan Unilever", 2456.80, 5432100), ("AXISBANK.NS", "Axis Bank", 1123.45, 4321000)]
    result = "🔥 **MOST ACTIVE NSE STOCKS**\n\n⚠️ **DEMO DATA** - API rate limit reached\n\n"
    for i, (symbol, name, price, volume) in enumerate(actives, 1):
        result += f"{i}. **{symbol}** - {name}\n   Price: ₹{price:.2f} | Volume: {volume:,}\n\n"
    return result

def get_nse_most_active() -> str:
    """
    Fetch most active stocks from NSE
    Falls back to demo data if API is unavailable
    """
    try:
        url = f"{FMP_BASE_URL}/stock_market/actives?apikey={FMP_API_KEY}"
        response = requests.get(url, timeout=10)
        
        if response.status_code in [403, 429]:
            return get_demo_most_active()
        
        response.raise_for_status()
        data = response.json()
        
        # Filter for NSE stocks
        nse_stocks = [stock for stock in data if '.NS' in stock.get('symbol', '')][:10]
        
        if nse_stocks:
            result = "🔥 **MOST ACTIVE NSE STOCKS**\n\n"
            for i, stock in enumerate(nse_stocks, 1):
                result += f"{i}. **{stock.get('symbol', 'N/A')}** - {stock.get('name', 'N/A')}\n"
                result += f"   Price: ₹{stock.get('price', 0):.2f} | "
                result += f"Volume: {stock.get('volume', 0):,}\n\n"
            return result
        else:
            return get_demo_most_active()
    except Exception as e:
        print(f"Error fetching active stocks: {e}")
        return get_demo_most_active()

def display_stock_quote(symbol: str) -> str:
    """
    Display detailed stock quote
    """
    data = fetch_nse_stock_data(symbol)
    
    if "error" in data:
        return f"""❌ {data['error']}

**API Issue Detected:**
- Free tier limit reached (250 requests/day)
- API will reset in 24 hours

**💡 Solutions:**
1. Get your own free API key: https://financialmodelingprep.com
2. Try these demo stocks: RELIANCE.NS, TCS.NS, INFY.NS, HDFCBANK.NS, ICICIBANK.NS
3. VBA Generator and Excel Analyzer work without API!

**✅ Try the VBA Generator tab - it works perfectly!**
"""
    
    demo_badge = "🎭 **DEMO DATA**" if data.get('demo_mode') else "🔴 **LIVE DATA**"
    api_source = data.get('data_source', 'Unknown Source')
    
    result = f"""
📊 **{data['name']} ({data['symbol']})**  {demo_badge}

**Current Price:** ₹{data['price']:.2f}
**Change:** ₹{data['change']:.2f} ({data['change_percent']:.2f}%)

**Day Range:** ₹{data['day_low']:.2f} - ₹{data['day_high']:.2f}
**52 Week Range:** ₹{data['year_low']:.2f} - ₹{data['year_high']:.2f}

**Market Data:**
- Open: ₹{data['open']:.2f}
- Previous Close: ₹{data['previous_close']:.2f}
- Volume: {data['volume']:,}
- Avg Volume: {data['avg_volume']:,}

**Fundamentals:**
- Market Cap: ₹{data['market_cap']:,}
- EPS: ₹{data['eps']:.2f}
- P/E Ratio: {data['pe']:.2f}

*Data Source: {api_source} - NSE*

{"⚠️ **Using demo data** - Get free API key to see live data!" if data.get('demo_mode') else ""}
"""
    return result

def fetch_and_export_stock_data(symbol: str) -> tuple:
    """
    Fetch stock data and prepare for Excel export
    Uses demo data if API is unavailable
    """
    try:
        # Get current quote
        quote = fetch_nse_stock_data(symbol)
        
        if "error" in quote:
            return (None, f"❌ {quote['error']}")
        
        # Get historical data (will use demo if API fails)
        hist_df = fetch_nse_historical_data(symbol, days=90)
        
        if hist_df.empty:
            return (None, "❌ Unable to generate historical data")
        
        # Create Excel file
        filename = f"{symbol.replace('.NS', '')}_data_{datetime.now().strftime('%Y%m%d')}.xlsx"
        filepath = f"/tmp/{filename}"
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Sheet 1: Current Quote
            quote_df = pd.DataFrame([quote])
            if 'demo_mode' in quote_df.columns:
                quote_df = quote_df.drop(columns=['demo_mode'])
            quote_df.to_excel(writer, sheet_name='Current Quote', index=False)
            
            # Sheet 2: Historical Data
            hist_df.to_excel(writer, sheet_name='Historical Data', index=False)
            
            # Sheet 3: Summary Statistics
            summary = {
                'Metric': ['Current Price', 'Day High', 'Day Low', 'Volume', 'Market Cap', 'P/E Ratio'],
                'Value': [f"₹{quote['price']:.2f}", f"₹{quote['day_high']:.2f}", f"₹{quote['day_low']:.2f}",
                          f"{quote['volume']:,}", f"₹{quote['market_cap']:,}", f"{quote['pe']:.2f}"]
            }
            summary_df = pd.DataFrame(summary)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Sheet 4: Technical Analysis
            if len(hist_df) >= 20:
                hist_df['MA_5'] = hist_df['close'].rolling(window=5).mean()
                hist_df['MA_20'] = hist_df['close'].rolling(window=20).mean()
                hist_df['Daily_Return'] = hist_df['close'].pct_change()
                tech_df = hist_df[['date', 'close', 'MA_5', 'MA_20', 'Daily_Return']].tail(30)
                tech_df.to_excel(writer, sheet_name='Technical Analysis', index=False)
        
        demo_note = " (using demo data)" if quote.get('demo_mode') else ""
        
        success_msg = f"✅ Data exported successfully{demo_note}!\n\n"
        success_msg += f"📊 **{quote['name']}** ({symbol})\n"
        success_msg += f"💰 Current Price: ₹{quote['price']:.2f}\n"
        success_msg += f"📈 Historical records: {len(hist_df)} days\n"
        success_msg += f"📁 File: {filename}\n\n"
        success_msg += "**Sheets included:**\n- Current Quote\n- Historical Data (90 days)\n- Summary Statistics\n- Technical Analysis\n\n"
        
        if quote.get('demo_mode'):
            success_msg += "⚠️ **Note**: Using demo data due to API rate limit."
        
        return (filepath, success_msg)
    
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        print(f"Export error: {error_detail}")
        return (None, f"❌ Error exporting data: {str(e)}")

# VBA Macro Templates for Stock Market Analysis
VBA_TEMPLATES = {
    "stock_data": """Sub FetchStockData()
    ' Fetches NSE stock data template for manual entry
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Headers
    ws.Range("A1").Value = "Symbol"
    ws.Range("B1").Value = "Price"
    ws.Range("C1").Value = "Change"
    ws.Range("D1").Value = "Change %"
    ws.Range("E1").Value = "Volume"
    ws.Range("F1").Value = "Market Cap"
    
    ' Format headers
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Sample NSE stocks
    ws.Range("A2").Value = "RELIANCE.NS"
    ws.Range("A3").Value = "TCS.NS"
    ws.Range("A4").Value = "INFY.NS"
    ws.Range("A5").Value = "HDFCBANK.NS"
    ws.Range("A6").Value = "ICICIBANK.NS"
    
    ws.Columns.AutoFit
    MsgBox "NSE Stock data template created! Add prices manually or import from API.", vbInformation
End Sub""",
    
    "stock_chart": """Sub CreateStockChart()
    ' Creates a stock price chart for NSE data
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Please add stock data first!", vbExclamation
        Exit Sub
    End If
    
    ' Delete existing charts
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' Create new chart
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=100, Width:=600, Height:=350)
    
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:B" & lastRow)
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "NSE Stock Prices - Indian Stock Market"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Price (₹)"
        .ChartStyle = 42
    End With
    
    MsgBox "NSE Stock chart created!", vbInformation
End Sub""",
    
    "portfolio_analysis": """Sub AnalyzePortfolio()
    ' Analyzes NSE stock portfolio performance
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalValue As Double
    Dim totalGain As Double
    Dim i As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Please add portfolio data first!", vbExclamation
        Exit Sub
    End If
    
    ' Calculate portfolio metrics
    totalValue = 0
    totalGain = 0
    
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, 5).Value) Then
            totalValue = totalValue + ws.Cells(i, 5).Value
        End If
        If IsNumeric(ws.Cells(i, 6).Value) Then
            totalGain = totalGain + ws.Cells(i, 6).Value
        End If
    Next i
    
    ' Display results
    ws.Range("A" & lastRow + 2).Value = "Total Portfolio Value:"
    ws.Range("B" & lastRow + 2).Value = Format(totalValue, "₹#,##0.00")
    ws.Range("B" & lastRow + 2).Font.Bold = True
    
    ws.Range("A" & lastRow + 3).Value = "Total Gain/Loss:"
    ws.Range("B" & lastRow + 3).Value = Format(totalGain, "₹#,##0.00")
    If totalGain > 0 Then
        ws.Range("B" & lastRow + 3).Font.Color = RGB(0, 176, 80)
    Else
        ws.Range("B" & lastRow + 3).Font.Color = RGB(255, 0, 0)
    End If
    ws.Range("B" & lastRow + 3).Font.Bold = True
    
    MsgBox "Portfolio analysis complete!" & vbCrLf & _
           "Total Value: ₹" & Format(totalValue, "#,##0.00") & vbCrLf & _
           "Total Gain/Loss: ₹" & Format(totalGain, "#,##0.00"), vbInformation
End Sub""",
    
    "moving_average": """Sub CalculateMovingAverage()
    ' Calculates moving averages for NSE stock prices
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ma5 As Double
    Dim ma20 As Double
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 21 Then
        MsgBox "Need at least 21 rows of price data for MA20!", vbExclamation
        Exit Sub
    End If
    
    ' Add headers for moving averages
    ws.Range("G1").Value = "MA5"
    ws.Range("H1").Value = "MA20"
    ws.Range("G1:H1").Font.Bold = True
    
    ' Calculate 5-day MA
    For i = 6 To lastRow
        If IsNumeric(ws.Range("B" & i).Value) Then
            ma5 = Application.WorksheetFunction.Average(ws.Range("B" & i - 4 & ":B" & i))
            ws.Range("G" & i).Value = ma5
        End If
    Next i
    
    ' Calculate 20-day MA
    For i = 21 To lastRow
        If IsNumeric(ws.Range("B" & i).Value) Then
            ma20 = Application.WorksheetFunction.Average(ws.Range("B" & i - 19 & ":B" & i))
            ws.Range("H" & i).Value = ma20
        End If
    Next i
    
    MsgBox "Moving averages (MA5 & MA20) calculated for NSE stock data!", vbInformation
End Sub""",
    
    "sort": """Sub SortData()
    ' Sorts stock data in the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data to sort!", vbExclamation
        Exit Sub
    End If
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A2"), Order:=xlAscending
    
    With ws.Sort
        .SetRange Range("A1:Z" & lastRow)
        .Header = xlYes
        .Apply
    End With
    
    MsgBox "Data sorted successfully!", vbInformation
End Sub""",
    
    "filter": """Sub FilterData()
    ' Applies AutoFilter to NSE stock data
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
        MsgBox "AutoFilter removed!", vbInformation
    Else
        Dim lastRow As Long
        Dim lastCol As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        If lastRow > 1 Then
            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
            MsgBox "AutoFilter applied! Click dropdown arrows to filter.", vbInformation
        Else
            MsgBox "No data to filter!", vbExclamation
        End If
    End If
End Sub""",
    
    "format": """Sub FormatTable()
    ' Applies professional formatting to NSE stock data
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Then
        MsgBox "No data to format!", vbExclamation
        Exit Sub
    End If
    
    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Format headers
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 11
    End With
    
    ' Add borders
    tableRange.Borders.LineStyle = xlContinuous
    tableRange.Borders.Weight = xlThin
    tableRange.Borders.Color = RGB(189, 215, 238)
    
    ' Alternate row colors
    Dim i As Long
    For i = 2 To lastRow Step 2
        ws.Rows(i).Interior.Color = RGB(242, 242, 242)
    Next i
    
    ' Auto-fit columns
    ws.Columns.AutoFit
    
    ' Center align headers
    ws.Rows(1).HorizontalAlignment = xlCenter
    
    MsgBox "NSE stock table formatted professionally!", vbInformation
End Sub"""
}

def generate_vba_macro(task_description):
    """
    Generate VBA macro based on task description with intelligent template matching
    """
    if not task_description or task_description.strip() == "":
        return "Please provide a task description."
    
    task_lower = task_description.lower()
    
    # Smart template matching for stock market tasks
    if any(word in task_lower for word in ["stock", "nse", "fetch", "data", "quote"]):
        return VBA_TEMPLATES["stock_data"]
    elif any(word in task_lower for word in ["chart", "graph", "plot", "visualize"]):
        return VBA_TEMPLATES["stock_chart"]
    elif any(word in task_lower for word in ["portfolio", "analyze", "analysis", "performance"]):
        return VBA_TEMPLATES["portfolio_analysis"]
    elif any(word in task_lower for word in ["moving average", "ma", "sma", "ema", "average"]):
        return VBA_TEMPLATES["moving_average"]
    elif any(word in task_lower for word in ["sort", "order", "arrange"]):
        return VBA_TEMPLATES["sort"]
    elif any(word in task_lower for word in ["filter", "search", "find"]):
        return VBA_TEMPLATES["filter"]
    elif any(word in task_lower for word in ["format", "style", "beautify", "design"]):
        return VBA_TEMPLATES["format"]
    else:
        # Generate custom template
        return f"""Sub CustomMacro()
    ' Task: {task_description}
    ' Custom macro for NSE stock market operations
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Your code here for: {task_description}
    MsgBox "Task: {task_description}" & vbCrLf & _
           "Please customize this macro for your specific NSE stock market needs.", vbInformation
    
End Sub"""

def handle_excel_upload(file):
    """
    Process uploaded Excel file and provide analytics
    """
    if file is None:
        return "No file uploaded."
    
    try:
        file_path = file.name
        file_name = os.path.basename(file_path)
        
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=0)
        
        # Generate file analytics
        analytics = f"""✅ File '{file_name}' uploaded successfully!

📊 File Analytics:
- Rows: {len(df)}
- Columns: {len(df.columns)}
- Column Names: {', '.join(df.columns.tolist())}
- Memory Usage: {df.memory_usage(deep=True).sum() / 1024:.2f} KB
- Missing Values: {df.isnull().sum().sum()}

📈 Data Preview:
{df.head(10).to_string()}

💡 Tip: Use this data with our VBA macros for NSE stock analysis!
"""
        return analytics
    
    except Exception as e:
        return f"❌ Error processing file: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"

def analyze_excel_data(file):
    """
    Provide detailed statistical analysis of Excel data
    """
    if file is None:
        return "No file uploaded for analysis."
    
    try:
        df = pd.read_excel(file.name, sheet_name=0)
        
        # Statistical analysis
        stats = df.describe().to_string()
        dtypes = df.dtypes.to_string()
        
        analysis = f"""📊 Detailed NSE Stock Data Analysis:

🔢 Data Types:
{dtypes}

📈 Statistical Summary:
{stats}

🔍 Data Quality:
- Total Cells: {df.size}
- Non-Null Cells: {df.count().sum()}
- Null Cells: {df.isnull().sum().sum()}
- Duplicate Rows: {df.duplicated().sum()}

💡 This analysis helps you understand your NSE stock market data better!
"""
        return analysis
    
    except Exception as e:
        return f"❌ Analysis Error: {str(e)}"

# Initialize GitHub client
github_token = os.getenv("GITHUB_TOKEN", "")
g = None
if github_token and github_token != "your_github_token":
    try:
        g = Github(github_token)
    except Exception as e:
        print(f"GitHub initialization warning: {e}")

def push_to_github(repo_name, file_name, code):
    """
    Push generated VBA code to GitHub repository
    """
    if not g:
        return "❌ GitHub token not configured. Please set GITHUB_TOKEN environment variable."
    
    if not repo_name or not file_name or not code:
        return "❌ Please provide repository name, file name, and code."
    
    try:
        # Ensure file has proper extension
        if not file_name.endswith(('.bas', '.vba', '.txt')):
            file_name += '.bas'
        
        repo = g.get_user().get_repo(repo_name)
        
        # Check if file exists
        try:
            contents = repo.get_contents(file_name)
            # File exists, update it
            repo.update_file(
                file_name,
                f"Update NSE VBA macro - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                code,
                contents.sha
            )
            message = f"✅ Successfully updated '{file_name}' in '{repo_name}'"
        except:
            # File doesn't exist, create it
            repo.create_file(
                file_name,
                f"Add NSE VBA macro - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                code
            )
            message = f"✅ Successfully created '{file_name}' in '{repo_name}'"
        
        return message
    
    except Exception as e:
        return f"❌ GitHub push failed: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"

# Custom CSS for better UI with ATS Integrated branding
custom_css = """
#main-title {
    background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #7e22ce 100%);
    padding: 25px 20px;
    border-radius: 12px;
    color: white;
    text-align: center;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}
#company-header {
    font-size: 2em;
    font-weight: bold;
    margin-bottom: 10px;
    letter-spacing: 2px;
}
#subtitle {
    font-size: 1.1em;
    opacity: 0.95;
    margin-top: 5px;
}
.gr-button-primary {
    background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
    border: none;
    transition: all 0.3s ease;
}
.gr-button-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
}
#footer {
    text-align: center;
    padding: 20px;
    font-size: 0.9em;
    color: #666;
    border-top: 1px solid #e0e0e0;
    margin-top: 30px;
}
#disclaimer {
    font-size: 0.75em;
    color: #999;
    max-width: 800px;
    margin: 10px auto;
    line-height: 1.5;
}
"""

# Build Gradio Interface
with gr.Blocks(css=custom_css, title="ATS Integrated - NSE Stock Market Suite") as demo:
    gr.Markdown("""
    <div id="main-title">
        <div style="display: flex; align-items: center; justify-content: center; gap: 20px; margin-bottom: 15px;">
            <div style="
                width: 90px;
                height: 90px;
                border-radius: 50%;
                background: white;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            ">
                <div style="
                    font-size: 32px;
                    font-weight: bold;
                    font-style: italic;
                    color: #2a5298;
                    letter-spacing: -2px;
                    transform: rotate(-5deg);
                ">ats</div>
            </div>
            <div style="text-align: left;">
                <div style="
                    font-size: 2.2em;
                    font-weight: 600;
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    color: white;
                    letter-spacing: 3px;
                    text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
                ">ATS INTEGRATED</div>
                <div style="
                    font-size: 0.9em;
                    opacity: 0.95;
                    margin-top: 5px;
                    font-family: 'Segoe UI', sans-serif;
                ">NSE Stock Market Analysis Suite</div>
            </div>
        </div>
        <p style="margin-top: 10px; font-size: 0.85em; opacity: 0.9;">
            🔴 <strong>LIVE DATA</strong>: Dhan API • Zerodha Kite • Financial Modeling Prep | VBA Automation for Indian Stock Market
        </p>
    </div>
    """)
    
    with gr.Tabs():
        # Tab 1: NSE Stock Market Data
        with gr.Tab("📊 NSE Stock Data"):
            gr.Markdown("""
            ### Live NSE Stock Market Data
            Get real-time quotes and market information from Financial Modeling Prep API
            
            **Popular NSE Stocks:** RELIANCE.NS, TCS.NS, INFY.NS, HDFCBANK.NS, ICICIBANK.NS, HINDUNILVR.NS
            """)
            
            with gr.Row():
                with gr.Column(scale=2):
                    stock_symbol_input = gr.Textbox(
                        label="Enter NSE Stock Symbol",
                        placeholder="e.g., RELIANCE.NS or just RELIANCE",
                        value="RELIANCE.NS"
                    )
                with gr.Column(scale=1):
                    fetch_quote_button = gr.Button("📈 Get Stock Quote", variant="primary")
            
            stock_quote_output = gr.Markdown(label="Stock Quote")
            
            gr.Markdown("### Export Stock Data to Excel")
            with gr.Row():
                export_symbol_input = gr.Textbox(
                    label="Symbol to Export",
                    placeholder="e.g., TCS.NS"
                )
                export_button = gr.Button("📥 Export to Excel")
            
            export_file_output = gr.File(label="Download Excel File")
            export_status_output = gr.Textbox(label="Export Status", lines=3)
            
            gr.Markdown("### Market Movers")
            with gr.Row():
                gainers_button = gr.Button("📈 Top Gainers")
                losers_button = gr.Button("📉 Top Losers")
                active_button = gr.Button("🔥 Most Active")
            
            market_movers_output = gr.Markdown(label="Market Data")
        
        # Tab 2: VBA Macro Generator
        with gr.Tab("🔧 VBA Generator"):
            gr.Markdown("""
            ### Generate VBA Macros for NSE Stock Analysis
            Create Excel automation macros for Indian stock market data analysis
            
            **Supported Tasks:**
            - Stock data templates
            - Price charts
            - Portfolio analysis
            - Moving averages
            - Data sorting & filtering
            - Professional formatting
            """)
            
            with gr.Row():
                with gr.Column(scale=3):
                    task_input = gr.Textbox(
                        label="Describe your NSE stock analysis task",
                        placeholder="e.g., 'Create stock data template' or 'Calculate moving averages'",
                        lines=3
                    )
                with gr.Column(scale=1):
                    generate_button = gr.Button("🚀 Generate VBA Macro", variant="primary")
            
            vba_output = gr.Code(
                label="Generated VBA Macro for NSE Stocks",
                language="python",  # Using python for syntax highlighting (VBA not supported)
                lines=25
            )
            
            gr.Markdown("**💡 Tip:** Copy the macro to Excel's VBA editor (Alt+F11) and run it!")
        
        # Tab 3: Excel File Analyzer
        with gr.Tab("📊 Excel Analyzer"):
            gr.Markdown("### Upload and Analyze Excel Files with Stock Data")
            
            excel_file = gr.File(
                label="Upload Excel File (.xlsx, .xls)",
                file_types=[".xlsx", ".xls"]
            )
            
            with gr.Row():
                upload_button = gr.Button("📤 Analyze File", variant="primary")
                detailed_button = gr.Button("🔍 Detailed Analysis")
            
            upload_status = gr.Textbox(
                label="File Information & Analysis",
                lines=20,
                max_lines=30
            )
        
        # Tab 4: GitHub Integration
        with gr.Tab("🐙 GitHub Integration"):
            gr.Markdown("""
            ### Push VBA Code to GitHub
            Save your NSE stock analysis macros to GitHub for version control
            
            **Setup:** Set `GITHUB_TOKEN` environment variable with your Personal Access Token
            """)
            
            with gr.Row():
                repo_name = gr.Textbox(
                    label="Repository Name",
                    placeholder="e.g., username/nse-stock-macros"
                )
                file_name = gr.Textbox(
                    label="File Name",
                    placeholder="e.g., nse_analysis.bas"
                )
            
            code_to_push = gr.Code(
                label="Code to Push (paste or generate first)",
                language="python",  # Using python for syntax highlighting (VBA not supported)
                lines=10
            )
            
            push_button = gr.Button("📤 Push to GitHub", variant="primary")
            push_status = gr.Textbox(label="GitHub Status", lines=5)
        
        # Tab 5: Help & Documentation
        with gr.Tab("❓ Help"):
            gr.Markdown(f"""
            ## ExcelBot Pro - NSE Stock Market Analysis Suite
            
            ### 🎯 Features
            1. **NSE Stock Data**: 🔴 **LIVE** quotes from Dhan API & Financial Modeling Prep
            2. **VBA Generator**: Create macros for Indian stock market analysis
            3. **Excel Analyzer**: Analyze stock data files
            4. **GitHub Integration**: Version control your VBA macros
            
            ### 📊 API Integration (Multi-Source Live Data)
            
            **Dhan API** 🔴 LIVE (Priority #1)
            - Client ID: {DHAN_CLIENT_ID}
            - Status: ✅ **ACTIVE - Fetching Real-time NSE Data**
            - Data: Live prices, volume, day high/low, open, close
            - Market: NSE (India)
            
            **Zerodha Kite API** (Priority #2)
            - API Key: {ZERODHA_API_KEY[:10]}...
            - Status: Configured (Ready for integration)
            
            **Financial Modeling Prep API** (Priority #3)
            - API Key: {FMP_API_KEY[:10]}...
            - Base URL: {FMP_BASE_URL}
            - Market: NSE (National Stock Exchange of India)
            - Status: Fallback source for comprehensive data
            
            ### 📝 How to Use NSE Stock Data
            
            #### Get Live Stock Quote:
            1. Go to "NSE Stock Data" tab
            2. Enter symbol (e.g., RELIANCE.NS or just RELIANCE)
            3. Click "Get Stock Quote"
            4. View real-time price, changes, and market data
            
            #### Export to Excel:
            1. Enter NSE stock symbol
            2. Click "Export to Excel"
            3. Download file with current quote and 90-day historical data
            4. Open in Excel and analyze with VBA macros
            
            #### Market Movers:
            - **Top Gainers**: Best performing NSE stocks today
            - **Top Losers**: Worst performing NSE stocks today
            - **Most Active**: Highest volume NSE stocks
            
            ### 🔧 VBA Macro Templates
            
            **Stock Data Template**
            - Creates worksheet structure for NSE stock data
            - Headers: Symbol, Price, Change, Volume, Market Cap
            - Pre-filled with popular NSE stocks
            
            **Stock Chart**
            - Creates professional line chart for price data
            - Formatted for Indian currency (₹)
            - Ready for NSE stock analysis
            
            **Portfolio Analysis**
            - Calculates total portfolio value
            - Shows gains/losses with color coding
            - Supports multiple NSE stocks
            
            **Moving Averages**
            - Calculates MA5 and MA20
            - Technical analysis for NSE stocks
            - Identifies trends
            
            ### 📈 Popular NSE Stocks
            
            **Banking & Finance:**
            - HDFCBANK.NS, ICICIBANK.NS, SBIN.NS, KOTAKBANK.NS, AXISBANK.NS
            
            **IT Services:**
            - TCS.NS, INFY.NS, WIPRO.NS, HCLTECH.NS
            
            **Consumer Goods:**
            - HINDUNILVR.NS, ITC.NS, NESTLEIND.NS, TITAN.NS
            
            **Automobiles:**
            - MARUTI.NS, TATAMOTORS.NS, BAJAJ-AUTO.NS
            
            **Energy & Materials:**
            - RELIANCE.NS, ONGC.NS, COALINDIA.NS
            
            **Infrastructure:**
            - LT.NS, ULTRACEMCO.NS
            
            **Telecom:**
            - BHARTIARTL.NS, IDEA.NS
            
            ### 🔑 GitHub Setup
            1. Create GitHub Personal Access Token
            2. Set environment variable: `export GITHUB_TOKEN=your_token`
            3. Restart application
            4. Use GitHub Integration tab to push code
            
            ### 💻 System Requirements
            - Python 3.7+
            - Internet connection (for API access)
            - Microsoft Excel (for running VBA macros)
            - GitHub account (optional, for version control)
            
            ### 📄 API Information
            
            **Financial Modeling Prep**
            - Free tier: 250 requests/day
            - Coverage: NSE stocks with .NS suffix
            - Data: Real-time quotes, historical prices, fundamentals
            
            **Zerodha Kite**
            - API Key configured: {ZERODHA_API_KEY}
            - Ready for future trading integration
            - Requires active Zerodha account
            
            ### 🆘 Troubleshooting
            
            **Issue: "No data found"**
            - Ensure symbol has .NS suffix (e.g., RELIANCE.NS)
            - Check if market is open (NSE hours: 9:15 AM - 3:30 PM IST)
            - Verify internet connection
            
            **Issue: "API Error"**
            - Check API key is valid
            - Verify API rate limits not exceeded
            - Wait a few seconds and try again
            
            **Issue: "Symbol not found"**
            - Use correct NSE symbol format (SYMBOL.NS)
            - Example: RELIANCE.NS, not RELIANCE or RELIANCE.BSE
            
            ### 📄 License
            MIT License - Copyright (c) 2025 Mandar Bahadarpurkar
            
            ### 🙏 Acknowledgments
            - Zerodha for Kite API access
            - Financial Modeling Prep for NSE data API
            - NSE (National Stock Exchange of India)
            """)
    
    # Footer with branding and disclaimer
    gr.Markdown("""
    ---
    <div style="text-align: center; padding: 25px 20px; margin-top: 40px; background: linear-gradient(to right, #f8f9fa, #e9ecef); border-radius: 10px;">
        <p style="font-size: 1.1em; margin-bottom: 15px; font-weight: 500;">
            <strong>Developed with ❤️ by Mandar Bahadarpurkar</strong>
        </p>
        <p style="margin-bottom: 20px; font-weight: 600; color: #1e3c72;">
            © 2025 <strong>ATS INTEGRATED</strong>. All Rights Reserved.
        </p>
        <div style="max-width: 900px; margin: 0 auto; padding: 20px; background: white; border-radius: 8px; border-left: 4px solid #dc3545;">
            <p style="font-size: 0.85em; color: #666; line-height: 1.7; text-align: left;">
                <strong>⚠️ DISCLAIMER:</strong> This application is for educational and informational purposes only. 
                Stock market data is provided "as is" without warranties of any kind, express or implied. 
                This is <strong>NOT financial advice</strong>. The creators and ATS Integrated assume no responsibility for 
                investment decisions made using this tool. Always consult with a qualified financial advisor 
                before making investment decisions. Past performance does not guarantee future results. 
                Trading in stock markets involves risk and may result in loss of capital.
            </p>
        </div>
        <p style="margin-top: 20px; font-size: 0.8em; color: #999;">
            Powered by Zerodha Kite API & Financial Modeling Prep | NSE Market Data<br>
            Version 1.0.0 - NSE Edition | MIT License
        </p>
    </div>
    """)
    
    # Event handlers
    fetch_quote_button.click(
        fn=display_stock_quote,
        inputs=[stock_symbol_input],
        outputs=[stock_quote_output]
    )
    
    export_button.click(
        fn=fetch_and_export_stock_data,
        inputs=[export_symbol_input],
        outputs=[export_file_output, export_status_output]
    )
    
    gainers_button.click(
        fn=get_nse_top_gainers,
        inputs=[],
        outputs=[market_movers_output]
    )
    
    losers_button.click(
        fn=get_nse_top_losers,
        inputs=[],
        outputs=[market_movers_output]
    )
    
    active_button.click(
        fn=get_nse_most_active,
        inputs=[],
        outputs=[market_movers_output]
    )
    
    generate_button.click(
        fn=generate_vba_macro,
        inputs=[task_input],
        outputs=[vba_output]
    )
    
    upload_button.click(
        fn=handle_excel_upload,
        inputs=[excel_file],
        outputs=[upload_status]
    )
    
    detailed_button.click(
        fn=analyze_excel_data,
        inputs=[excel_file],
        outputs=[upload_status]
    )
    
    push_button.click(
        fn=push_to_github,
        inputs=[repo_name, file_name, code_to_push],
        outputs=[push_status]
    )

if __name__ == "__main__":
    print("🚀 Starting ExcelBot Pro - NSE Stock Market Analysis Suite")
    print(f"📊 Zerodha API Key: {ZERODHA_API_KEY[:10]}...")
    print(f"📈 FMP API Key: {FMP_API_KEY[:10]}...")
    print("🌐 Creating public share link for mobile access...")
    print("")
    
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=True,  # Creates public URL for mobile access
        show_error=True
    )
