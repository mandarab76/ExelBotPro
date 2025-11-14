# 🚀 ExcelBot Pro - NSE Stock Market Quick Start Guide

Get started with NSE stock market analysis in 5 minutes!

## ⚡ Super Quick Start

### Option 1: Automated Setup (Recommended)

**Linux/Mac:**
```bash
chmod +x setup.sh
./setup.sh
```

**Windows:**
```cmd
setup.bat
```

### Option 2: Manual Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Run application
python excelbot_chat.py

# Open browser
http://localhost:7860
```

### Option 3: Docker

```bash
docker-compose up -d
# Access at http://localhost:7860
```

---

## 🎯 First Steps with NSE Stocks

### 1. Get Your First Stock Quote (30 seconds)

1. Open **NSE Stock Data** tab
2. Enter: `RELIANCE.NS` (or just `RELIANCE`)
3. Click **"📈 Get Stock Quote"**
4. View real-time data:
   - Current Price: ₹2,456.75
   - Change: +₹23.50 (+0.97%)
   - Day Range, Market Cap, Volume

**Try these popular NSE stocks:**
- `TCS.NS` - Tata Consultancy Services
- `INFY.NS` - Infosys
- `HDFCBANK.NS` - HDFC Bank
- `ICICIBANK.NS` - ICICI Bank
- `HINDUNILVR.NS` - Hindustan Unilever

### 2. Generate Your First VBA Macro (1 minute)

1. Go to **VBA Generator** tab
2. Type: `"Create stock data template"`
3. Click **"🚀 Generate VBA Macro"**
4. Copy the generated code
5. In Excel: Press `Alt+F11`
6. Insert > Module
7. Paste code and press `F5` to run

**Result:** Excel worksheet with NSE stock data structure!

### 3. Export Stock Data to Excel (1 minute)

1. Go to **NSE Stock Data** tab
2. Enter: `TCS.NS`
3. Click **"📥 Export to Excel"**
4. Download file: `TCS_data_20251113.xlsx`
5. Open in Excel:
   - Sheet 1: Current Quote
   - Sheet 2: 90-day Historical Data

### 4. Check Market Movers (30 seconds)

In **NSE Stock Data** tab, click:
- **📈 Top Gainers** - Best performing stocks today
- **📉 Top Losers** - Worst performing stocks today
- **🔥 Most Active** - Highest volume stocks

---

## 📊 Popular NSE Stock Symbols

### Banking Sector
```
HDFCBANK.NS    - HDFC Bank
ICICIBANK.NS   - ICICI Bank
SBIN.NS        - State Bank of India
KOTAKBANK.NS   - Kotak Mahindra Bank
AXISBANK.NS    - Axis Bank
```

### IT Sector
```
TCS.NS         - Tata Consultancy Services
INFY.NS        - Infosys
WIPRO.NS       - Wipro
HCLTECH.NS     - HCL Technologies
TECHM.NS       - Tech Mahindra
```

### Consumer Goods
```
RELIANCE.NS    - Reliance Industries
HINDUNILVR.NS  - Hindustan Unilever
ITC.NS         - ITC Limited
NESTLEIND.NS   - Nestle India
TITAN.NS       - Titan Company
```

### Automobiles
```
MARUTI.NS      - Maruti Suzuki
TATAMOTORS.NS  - Tata Motors
M&M.NS         - Mahindra & Mahindra
BAJAJ-AUTO.NS  - Bajaj Auto
```

---

## 🔧 Example VBA Macro Tasks

Copy these into the VBA Generator:

### Stock Data Management
```
"Create stock data template"
"Fetch NSE stock prices"
"Format stock data table"
```

### Charts & Visualization
```
"Create stock price chart"
"Plot NSE stock trends"
"Visualize portfolio performance"
```

### Analysis & Calculations
```
"Calculate moving averages"
"Analyze portfolio"
"Calculate stock returns"
```

### Data Operations
```
"Sort stock data by price"
"Filter profitable stocks"
"Remove duplicate entries"
```

---

## 🎓 5-Minute Tutorial

### Complete Workflow: Analyze RELIANCE Stock

**Step 1: Get Live Data** (1 min)
```
1. NSE Stock Data tab
2. Enter: RELIANCE.NS
3. Click: Get Stock Quote
4. Note current price
```

**Step 2: Export Historical Data** (1 min)
```
1. Same tab
2. Enter: RELIANCE.NS in export section
3. Click: Export to Excel
4. Download file
```

**Step 3: Generate Analysis Macro** (1 min)
```
1. VBA Generator tab
2. Type: "Calculate moving averages"
3. Click: Generate
4. Copy code
```

**Step 4: Apply in Excel** (2 min)
```
1. Open exported RELIANCE file
2. Alt+F11 (VBA Editor)
3. Insert > Module
4. Paste macro
5. F5 to run
6. View MA5 and MA20 columns
```

---

## 💡 Pro Tips

### Tip 1: Symbol Format
- ✅ Correct: `RELIANCE.NS`, `TCS.NS`
- ❌ Wrong: `RELIANCE`, `TCS`, `RELIANCE.BSE`
- 💡 App auto-adds `.NS` if you forget!

### Tip 2: Market Hours
- NSE is open: 9:15 AM - 3:30 PM IST (Mon-Fri)
- Live data available during market hours
- Historical data always available

### Tip 3: API Limits
- Financial Modeling Prep: 250 free requests/day
- Plan your queries wisely
- Export data to Excel for offline analysis

### Tip 4: VBA Macros
- Always enable macros in Excel
- Test on sample data first
- Save workbooks as .xlsm (macro-enabled)

### Tip 5: Portfolio Tracking
- Create Excel file with portfolio
- Use "Analyze portfolio" macro
- Update prices from NSE Stock Data tab

---

## 🔑 API Keys Configuration

### Pre-configured Keys (Ready to Use!)

The application comes with working API keys:

**Zerodha Kite:**
```
API Key: kr8ob80gcmucrvph
Status: Configured ✅
```

**Financial Modeling Prep:**
```
API Key: rtD0v37SghQ4gMZNfX7q2Arv6RO7StUv
Status: Active ✅
Free Tier: 250 requests/day
```

### Optional: Use Your Own Keys

Create `.env` file:
```bash
cp .env.example .env
```

Edit `.env`:
```
ZERODHA_API_KEY=your_zerodha_key
FMP_API_KEY=your_fmp_key
```

Restart application.

---

## 🆘 Troubleshooting

### Problem 1: "No data found"
**Solution:**
- Check symbol format: Must be `SYMBOL.NS`
- Verify internet connection
- Try another popular stock (e.g., `TCS.NS`)

### Problem 2: "API Error"
**Solution:**
- Check if you've exceeded 250 requests/day
- Wait 5-10 seconds between requests
- Restart application

### Problem 3: "Python not found"
**Solution:**
```bash
# Install Python 3.7+ from python.org
# On Windows, check "Add Python to PATH"
python --version  # Should show 3.7+
```

### Problem 4: "Port 7860 in use"
**Solution:**
```bash
# Kill existing process
lsof -i :7860  # Find PID
kill -9 PID    # Kill it

# Or change port in code
# excelbot_chat.py: demo.launch(server_port=8080)
```

### Problem 5: "Excel macro not working"
**Solution:**
1. Enable macros: File > Options > Trust Center > Macro Settings
2. Save as .xlsm format
3. Check VBA code syntax
4. Test with sample data

---

## 📚 Next Steps

### Learn More (Choose Your Path)

**Beginner:**
1. ✅ Try all VBA templates
2. ✅ Export 3-4 popular stocks
3. ✅ Create simple portfolio in Excel
4. ✅ Read full README.md

**Intermediate:**
1. ✅ Analyze portfolio performance
2. ✅ Calculate technical indicators
3. ✅ Create custom VBA macros
4. ✅ Set up GitHub integration

**Advanced:**
1. ✅ Deploy to production server
2. ✅ Integrate with trading account (Zerodha)
3. ✅ Build automated trading strategies
4. ✅ Contribute to the project

### Documentation

- `README.md` - Complete documentation
- `API_GUIDE.md` - Detailed API usage
- `VBA_REFERENCE.md` - All VBA macros explained
- `DEPLOYMENT.md` - Production deployment
- Help tab - In-app comprehensive help

---

## 🎉 You're Ready to Trade Data!

### Quick Reference

**Get Live Quote:**
```
Tab: NSE Stock Data
Input: RELIANCE.NS
Button: Get Stock Quote
```

**Export Data:**
```
Tab: NSE Stock Data
Input: TCS.NS
Button: Export to Excel
```

**Generate VBA:**
```
Tab: VBA Generator
Input: "Create stock chart"
Button: Generate VBA Macro
```

**Market Movers:**
```
Tab: NSE Stock Data
Buttons: Top Gainers / Top Losers / Most Active
```

---

## 🇮🇳 Indian Stock Market Focus

This tool is built specifically for Indian traders:
- ✅ NSE stock symbols (.NS)
- ✅ Indian Rupee (₹) formatting
- ✅ IST timezone
- ✅ Popular Indian stocks pre-configured
- ✅ Zerodha Kite integration ready
- ✅ Cultural and market context

---

## 📞 Get Help

**In-App:** Click "Help" tab for comprehensive guide  
**Issues:** https://github.com/mandarab76/ExcelBotPro/issues  
**Discussions:** https://github.com/mandarab76/ExcelBotPro/discussions  

---

<div align="center">

**🚀 START ANALYZING NSE STOCKS NOW! 🚀**

Made with ❤️ for Indian Stock Market

⭐ Star the repo if this helps you!

</div>
