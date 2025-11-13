# 🚀 ExcelBot Pro - Quick Start Guide

Get ExcelBot Pro up and running in 5 minutes!

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

The script will:
- ✅ Check Python installation
- ✅ Install dependencies
- ✅ Configure environment
- ✅ Test the installation
- ✅ Start the application

### Option 2: Manual Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the application
python excelbot_chat.py

# 3. Open browser to http://localhost:7860
```

### Option 3: Docker

```bash
# Build and run with Docker
docker build -t excelbotpro .
docker run -p 7860:7860 excelbotpro

# Or use Docker Compose
docker-compose up -d
```

## 🎯 First Steps

### 1. Generate Your First VBA Macro (30 seconds)

1. Open **VBA Generator** tab
2. Type: `"Sort my data by first column"`
3. Click **Generate VBA Macro**
4. Copy the code
5. In Excel: Press `Alt+F11`, paste the code, press `F5`

### 2. Analyze an Excel File (1 minute)

1. Open **Excel Analyzer** tab
2. Upload `sample1.xlsx` or `sample2.xlsx`
3. Click **Analyze File**
4. View detailed statistics and data preview

### 3. Push Code to GitHub (2 minutes)

1. Get GitHub token from https://github.com/settings/tokens
2. Set environment variable:
   ```bash
   export GITHUB_TOKEN=your_token
   ```
3. Restart the app
4. Generate a macro
5. Go to **GitHub Integration** tab
6. Enter repo name and file name
7. Paste code and click **Push to GitHub**

## 📝 Example Tasks to Try

Copy and paste these into the VBA Generator:

### Basic Tasks
- `"Sort my data alphabetically"`
- `"Apply filters to my spreadsheet"`
- `"Format my table with colors"`

### Advanced Tasks
- `"Highlight all duplicate values"`
- `"Remove duplicate rows from my data"`
- `"Add sum totals to numeric columns"`
- `"Create a pivot table from this data"`

## 🔧 Troubleshooting

### Application won't start?
```bash
# Check Python version (need 3.7+)
python --version

# Reinstall dependencies
pip install --upgrade -r requirements.txt
```

### GitHub push fails?
```bash
# Check token is set
echo $GITHUB_TOKEN  # Linux/Mac
echo %GITHUB_TOKEN%  # Windows

# Set token
export GITHUB_TOKEN=your_token  # Linux/Mac
set GITHUB_TOKEN=your_token     # Windows
```

### Port 7860 already in use?
```bash
# Find and kill process using port 7860
lsof -i :7860  # Linux/Mac
netstat -ano | findstr :7860  # Windows

# Or change port in excelbot_chat.py:
# demo.launch(server_port=8080)
```

## 📚 Next Steps

- 📖 Read the full [README.md](README.md)
- 🚀 Check [DEPLOYMENT.md](DEPLOYMENT.md) for production setup
- 🤝 See [CONTRIBUTING.md](CONTRIBUTING.md) to contribute
- 📝 View [CHANGELOG.md](CHANGELOG.md) for version history

## 💡 Pro Tips

1. **Save time**: Bookmark frequently used macros
2. **GitHub integration**: Version control your VBA code
3. **Custom templates**: Add your own in `VBA_TEMPLATES` dictionary
4. **Batch processing**: Analyze multiple files in sequence
5. **Excel shortcuts**: `Alt+F11` (VBA Editor), `F5` (Run macro)

## 🎓 Learning Resources

### VBA Basics
- [Microsoft VBA Documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [Excel VBA Tutorial](https://www.excel-easy.com/vba.html)

### ExcelBot Pro Help
- In-app: Click **Help** tab
- GitHub: [Issues](https://github.com/mandarab76/ExcelBotPro/issues)
- GitHub: [Discussions](https://github.com/mandarab76/ExcelBotPro/discussions)

## ✅ Checklist for New Users

- [ ] Installed Python 3.7+
- [ ] Installed dependencies
- [ ] Started ExcelBot Pro
- [ ] Generated first VBA macro
- [ ] Tested macro in Excel
- [ ] Analyzed an Excel file
- [ ] (Optional) Set up GitHub integration
- [ ] (Optional) Deployed with Docker

## 🎉 You're Ready!

You now have a powerful Excel automation tool at your fingertips. Start automating your Excel tasks today!

---

**Need help?** Open an issue on GitHub or check the Help tab in the app.

**Enjoying ExcelBot Pro?** ⭐ Star the repository!
