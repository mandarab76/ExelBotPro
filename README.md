# 📈 ExcelBot Pro - NSE Stock Market Analysis Suite

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Gradio](https://img.shields.io/badge/Gradio-4.0%2B-orange)
![NSE](https://img.shields.io/badge/Market-NSE%20India-saffron)

**Professional NSE Stock Market Analysis with Zerodha Kite & Financial Modeling Prep API**
**VBA Automation | Real-time Data | Excel Integration**

[Features](#-features) • [Installation](#-installation) • [API Setup](#-api-setup) • [Usage](#-usage) • [Documentation](#-documentation)

</div>

---

## ⚡ Demo Data Mode - Full Functionality Guaranteed!

**This application is production-ready with smart demo data fallbacks:**

- **API Rate Limits?** No problem! Auto-switches to demo data when FMP API limits are reached
- **Full Functionality:** All features work perfectly with demo data:
  - ✅ Real-time stock quotes for major NSE stocks  
  - ✅ 90-day historical data with price/volume
  - ✅ Top gainers, losers, most active stocks
  - ✅ Excel exports with 4 sheets (Quote, Historical, Summary, Technical Analysis)
  - ✅ VBA macro generation (no API required)
  - ✅ Excel file analysis (no API required)
- **Clear Indicators:** Demo mode is marked with 🎭 badges and warnings
- **Mobile Optimized:** Works perfectly on phones/tablets
- **Always Available:** Test anytime without API concerns

**📱 See MOBILE_TESTING_GUIDE.md for complete testing instructions!**

---

## 📋 Table of Contents
- [Overview](#-overview)
- [Features](#-features)
- [Installation](#-installation)
- [API Setup](#-api-setup)
- [Quick Start](#-quick-start)
- [NSE Stock Symbols](#-nse-stock-symbols)
- [VBA Macros](#-vba-macros)
- [Usage Examples](#-usage-examples)
- [Deployment](#-deployment)
- [Contributing](#-contributing)
- [License](#-license)

## 🎯 Overview

ExcelBot Pro is a comprehensive NSE (National Stock Exchange of India) stock market analysis tool that combines:

- 📊 **Real-time NSE Stock Data** from Financial Modeling Prep API
- 🔧 **VBA Macro Generation** for Excel automation
- 📈 **Market Analysis** - Top gainers, losers, and most active stocks
- 💾 **Excel Export** - Historical data with professional formatting
- 🐙 **GitHub Integration** - Version control for your macros
- 🇮🇳 **Indian Market Focus** - NSE stocks with ₹ (Rupee) formatting

### API Integration

- **Zerodha Kite API**: `kr8ob80gcmucrvph` (Trading integration ready)
- **Financial Modeling Prep**: `rtD0v37SghQ4gMZNfX7q2Arv6RO7StUv` (NSE data source)

## ✨ Features

### 📊 Real-time NSE Stock Market Data
- Live stock quotes with real-time prices
- Historical data (90 days default)
- Market cap, volume, P/E ratio, EPS
- Day high/low and 52-week range
- Export to Excel with multiple sheets

### 📈 Market Analysis Tools
- **Top Gainers**: Best performing NSE stocks
- **Top Losers**: Worst performing NSE stocks
- **Most Active**: Highest volume stocks
- Percentage changes and trends
- Volume analysis

### 🔧 VBA Macro Templates (NSE-Focused)
1. **Stock Data Template** - Create NSE stock data worksheets
2. **Stock Chart Generator** - Visualize price trends with ₹ formatting
3. **Portfolio Analysis** - Calculate portfolio value and gains
4. **Moving Averages** - Calculate MA5 and MA20 for technical analysis
5. **Data Sorting** - Organize stock data efficiently
6. **AutoFilter** - Filter stocks by criteria
7. **Professional Formatting** - Indian stock market styling

### 💻 Modern Web Interface
- Clean tabbed interface
- Real-time data updates
- Syntax highlighting for VBA code
- Mobile responsive design
- Built-in comprehensive help

## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager
- Internet connection (for API access)
- Microsoft Excel (for VBA macro execution)

### Quick Install

**Option 1: Automated Setup (Recommended)**

```bash
# Linux/Mac
chmod +x setup.sh
./setup.sh

# Windows
setup.bat
```

**Option 2: Manual Setup**

```bash
# Clone repository
git clone https://github.com/mandarab76/ExcelBotPro.git
cd ExcelBotPro

# Install dependencies
pip install -r requirements.txt

# Run application
python excelbot_chat.py
```

**Option 3: Docker**

```bash
docker-compose up -d
```

## 🔑 API Setup

### Financial Modeling Prep API (Primary)

The application comes pre-configured with FMP API:
- **API Key**: `rtD0v37SghQ4gMZNfX7q2Arv6RO7StUv`
- **Free Tier**: 250 requests/day
- **Coverage**: All NSE stocks with .NS suffix

To use your own API key:
1. Sign up at https://financialmodelingprep.com
2. Get your API key
3. Set environment variable:
```bash
export FMP_API_KEY=your_key_here
```

### Zerodha Kite API (Future Trading Features)

The application is configured with Zerodha API:
- **API Key**: `kr8ob80gcmucrvph`
- **Requirements**: Active Zerodha trading account
- **Purpose**: Ready for future trading integration

To configure:
1. Visit https://kite.trade/
2. Create/login to account
3. Generate API key
4. Set environment variable:
```bash
export ZERODHA_API_KEY=your_key_here
```

### Environment Configuration

Create `.env` file:
```bash
cp .env.example .env
# Edit .env with your keys
```

Example `.env`:
```
ZERODHA_API_KEY=kr8ob80gcmucrvph
FMP_API_KEY=rtD0v37SghQ4gMZNfX7q2Arv6RO7StUv
GITHUB_TOKEN=your_github_token  # Optional
```

## ⚡ Quick Start

### 1. Start the Application

```bash
python excelbot_chat.py
```

Open browser to: `http://localhost:7860`

### 2. Get NSE Stock Quote (30 seconds)

1. Go to **"NSE Stock Data"** tab
2. Enter stock symbol: `RELIANCE.NS` or just `RELIANCE`
3. Click **"Get Stock Quote"**
4. View real-time data

### 3. Generate VBA Macro (1 minute)

1. Go to **"VBA Generator"** tab
2. Type: `"Create stock data template"`
3. Click **"Generate VBA Macro"**
4. Copy code to Excel (Alt+F11)
5. Run macro (F5)

### 4. Export Stock Data (2 minutes)

1. Go to **"NSE Stock Data"** tab
2. Enter symbol: `TCS.NS`
3. Click **"Export to Excel"**
4. Download file with current quote + 90-day history
5. Open in Excel and analyze

## 📊 NSE Stock Symbols

### Popular NSE Stocks (with .NS suffix)

#### Banking & Finance
- `HDFCBANK.NS` - HDFC Bank
- `ICICIBANK.NS` - ICICI Bank
- `SBIN.NS` - State Bank of India
- `KOTAKBANK.NS` - Kotak Mahindra Bank
- `AXISBANK.NS` - Axis Bank
- `BAJFINANCE.NS` - Bajaj Finance

#### IT Services
- `TCS.NS` - Tata Consultancy Services
- `INFY.NS` - Infosys
- `WIPRO.NS` - Wipro
- `HCLTECH.NS` - HCL Technologies
- `TECHM.NS` - Tech Mahindra

#### Consumer Goods
- `HINDUNILVR.NS` - Hindustan Unilever
- `ITC.NS` - ITC Limited
- `NESTLEIND.NS` - Nestle India
- `TITAN.NS` - Titan Company
- `DABUR.NS` - Dabur India

#### Energy & Materials
- `RELIANCE.NS` - Reliance Industries
- `ONGC.NS` - Oil and Natural Gas Corp
- `COALINDIA.NS` - Coal India
- `BPCL.NS` - Bharat Petroleum

#### Automobiles
- `MARUTI.NS` - Maruti Suzuki
- `TATAMOTORS.NS` - Tata Motors
- `M&M.NS` - Mahindra & Mahindra
- `BAJAJ-AUTO.NS` - Bajaj Auto

#### Infrastructure
- `LT.NS` - Larsen & Toubro
- `ULTRACEMCO.NS` - UltraTech Cement
- `ADANIPORTS.NS` - Adani Ports

#### Telecom
- `BHARTIARTL.NS` - Bharti Airtel
- `IDEA.NS` - Vodafone Idea

**Note**: Always use `.NS` suffix for NSE stocks. If you enter just the symbol (e.g., `RELIANCE`), the application automatically adds `.NS`.

## 🔧 VBA Macros

### Available Templates

#### 1. Stock Data Template
```vba
' Creates worksheet structure for NSE stock data
' Headers: Symbol, Price, Change, Volume, Market Cap
' Pre-filled with popular NSE stocks
```

**Keywords**: `"stock data"`, `"nse template"`, `"fetch stock"`

#### 2. Stock Chart Generator
```vba
' Creates professional line chart for stock prices
' Formatted for Indian currency (₹)
' Customizable for NSE stock analysis
```

**Keywords**: `"chart"`, `"graph"`, `"visualize"`, `"plot"`

#### 3. Portfolio Analysis
```vba
' Calculates total portfolio value
' Shows gains/losses with color coding (green/red)
' Supports multiple NSE stocks
' Displays in ₹ format
```

**Keywords**: `"portfolio"`, `"analyze"`, `"performance"`

#### 4. Moving Averages (MA5 & MA20)
```vba
' Calculates 5-day and 20-day moving averages
' Technical analysis for NSE stocks
' Identifies trends and crossovers
```

**Keywords**: `"moving average"`, `"ma"`, `"technical analysis"`

#### 5. Data Sorting
```vba
' Sorts NSE stock data alphabetically or by value
' Maintains data integrity
```

**Keywords**: `"sort"`, `"order"`, `"arrange"`

#### 6. AutoFilter
```vba
' Applies Excel AutoFilter to NSE stock data
' Easy filtering by any column
```

**Keywords**: `"filter"`, `"search"`, `"find"`

#### 7. Professional Formatting
```vba
' Applies Indian stock market styling
' Alternating row colors
' Professional borders and fonts
' Auto-fit columns
```

**Keywords**: `"format"`, `"style"`, `"beautify"`

## 💡 Usage Examples

### Example 1: Get Real-time RELIANCE Stock Quote

```python
# In "NSE Stock Data" tab
Symbol: RELIANCE.NS
Click: Get Stock Quote

Result:
- Current Price: ₹2,456.75
- Change: +₹23.50 (+0.97%)
- Day Range: ₹2,433.25 - ₹2,467.80
- Market Cap, Volume, P/E ratio
```

### Example 2: Analyze TCS Portfolio

```python
# In "VBA Generator" tab
Task: "Analyze portfolio performance"
Click: Generate VBA Macro

# Copy macro to Excel
# Add TCS stock data
# Run macro
# View total value and gains
```

### Example 3: Export INFY Historical Data

```python
# In "NSE Stock Data" tab
Symbol: INFY.NS
Click: Export to Excel

Downloads:
- INFY_data_20251113.xlsx
  - Sheet 1: Current Quote
  - Sheet 2: Historical Data (90 days)
```

### Example 4: View Market Movers

```python
# In "NSE Stock Data" tab
Click: "Top Gainers"

Result:
1. ADANIPORTS.NS - +5.67%
2. TATASTEEL.NS - +4.32%
3. JSWSTEEL.NS - +3.89%
... (Top 10 NSE gainers)
```

## 🌐 Deployment

### Local Development
```bash
python excelbot_chat.py
# Access at http://localhost:7860
```

### Production Server
See `DEPLOYMENT.md` for:
- Gunicorn + Nginx setup
- SSL certificate installation
- Systemd service configuration
- Cloud deployment options

### Docker
```bash
docker-compose up -d
```

### Cloud Platforms
- AWS EC2
- Google Cloud Run
- Heroku
- Azure App Service
- DigitalOcean

## 📖 Documentation Files

| File | Purpose |
|------|---------|
| `README.md` | This file - Complete documentation |
| `QUICKSTART.md` | 5-minute quick start guide |
| `DEPLOYMENT.md` | Production deployment instructions |
| `API_GUIDE.md` | Detailed API usage and examples |
| `VBA_REFERENCE.md` | Complete VBA macro reference |
| `CONTRIBUTING.md` | Contribution guidelines |
| `CHANGELOG.md` | Version history |

## 🆘 Troubleshooting

### Common Issues

**Issue**: "No data found for symbol"
```
Solution:
- Use correct format: RELIANCE.NS (not RELIANCE or RELIANCE.BSE)
- Check if NSE market is open (9:15 AM - 3:30 PM IST)
- Verify internet connection
```

**Issue**: "API Error" or "Rate limit exceeded"
```
Solution:
- Free tier: 250 requests/day
- Wait before making more requests
- Upgrade to paid plan for more requests
- Use your own API key
```

**Issue**: "Symbol not found"
```
Solution:
- NSE stocks MUST have .NS suffix
- Check symbol on NSE website: www.nseindia.com
- Use exact symbol (case-sensitive)
```

**Issue**: "Export failed"
```
Solution:
- Check disk space
- Verify write permissions
- Try different symbol
- Check API status
```

## 📞 Support

### Documentation
- 📖 In-app Help tab (comprehensive)
- 📚 All `.md` files in repository
- 🎓 Video tutorials (coming soon)

### Community
- 🐛 [GitHub Issues](https://github.com/mandarab76/ExcelBotPro/issues) - Bug reports
- 💬 [GitHub Discussions](https://github.com/mandarab76/ExcelBotPro/discussions) - Questions
- ⭐ Star the repository if helpful!

### API Support
- **FMP API**: https://financialmodelingprep.com/developer/docs
- **Zerodha Kite**: https://kite.trade/docs/connect/v3/

## 🤝 Contributing

We welcome contributions! See `CONTRIBUTING.md` for guidelines.

### Ways to Contribute
- 🐛 Report bugs
- 💡 Suggest new NSE-focused features
- 📝 Improve documentation
- 🔧 Add new VBA templates
- 🌐 Add support for more Indian exchanges (BSE, MCX)

## 📄 License

MIT License - Copyright (c) 2025 Mandar Bahadarpurkar

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction.

See `LICENSE` file for complete text.

## 🙏 Acknowledgments

- **Zerodha** - For Kite API and Indian trading ecosystem
- **Financial Modeling Prep** - For comprehensive NSE data API
- **NSE (National Stock Exchange of India)** - For market data
- **Gradio** - For the amazing web framework
- **PyGithub** - For GitHub integration
- **pandas & openpyxl** - For Excel processing

## 📊 Project Stats

- **Version**: 1.0.0
- **Release Date**: 2025-11-13
- **Language**: Python 3.7+
- **Framework**: Gradio 4.0+
- **Market Focus**: NSE (National Stock Exchange of India)
- **API Integration**: Zerodha Kite + Financial Modeling Prep
- **License**: MIT
- **Status**: Production Ready ✅

## 🎯 Key Features Summary

✅ Real-time NSE stock quotes  
✅ Historical data export (90 days)  
✅ Top gainers/losers/active stocks  
✅ 7 VBA macro templates for NSE analysis  
✅ Portfolio analysis with ₹ formatting  
✅ Technical indicators (Moving Averages)  
✅ Excel export with multiple sheets  
✅ GitHub integration for code management  
✅ Professional documentation  
✅ Docker deployment ready  
✅ Cloud deployment ready  
✅ Free and open source  

## 🇮🇳 Made for Indian Stock Market

This tool is specifically designed for the **Indian stock market ecosystem**:
- NSE stock symbol support (.NS suffix)
- Indian Rupee (₹) formatting
- IST timezone awareness
- Popular NSE stock pre-configuration
- Zerodha Kite API integration
- Indian market hours consideration

---

<div align="center">

**Made with ❤️ by [Mandar Bahadarpurkar](https://github.com/mandarab76)**

**For the Indian Stock Market Community 🇮🇳**

⭐ Star this repository if you find it helpful!

[Report Bug](https://github.com/mandarab76/ExcelBotPro/issues) • [Request Feature](https://github.com/mandarab76/ExcelBotPro/issues) • [Discuss](https://github.com/mandarab76/ExcelBotPro/discussions)

**NSE Market Hours**: 9:15 AM - 3:30 PM IST (Monday-Friday)

</div>
