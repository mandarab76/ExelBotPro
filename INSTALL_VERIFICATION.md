# ✅ ExcelBot Pro - Installation Verification Guide

This guide helps you verify that ExcelBot Pro is correctly installed and functioning.

## 📋 Pre-Installation Checklist

Before installing, verify you have:

- [ ] Python 3.7 or higher installed
- [ ] pip package manager installed
- [ ] Internet connection (for downloading dependencies)
- [ ] At least 500MB free disk space
- [ ] Modern web browser (Chrome, Firefox, Safari, or Edge)

### Check Python Version

```bash
python --version
# or
python3 --version
```

Expected output: `Python 3.7.x` or higher

### Check pip Version

```bash
pip --version
# or
pip3 --version
```

Expected output: `pip x.x.x from ...`

## 🔍 Installation Verification Steps

### Step 1: Verify Package Contents

Check that all files are present:

```bash
ls -la
```

You should see these files:
- ✓ excelbot_chat.py
- ✓ requirements.txt
- ✓ README.md
- ✓ QUICKSTART.md
- ✓ DEPLOYMENT.md
- ✓ CONTRIBUTING.md
- ✓ CHANGELOG.md
- ✓ PROJECT_SUMMARY.md
- ✓ PACKAGE_INFO.txt
- ✓ LICENSE
- ✓ .env.example
- ✓ .gitignore
- ✓ .dockerignore
- ✓ Dockerfile
- ✓ docker-compose.yml
- ✓ setup.sh
- ✓ setup.bat
- ✓ sample1.xlsx
- ✓ sample2.xlsx

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

Expected output:
```
Successfully installed gradio-x.x.x pandas-x.x.x openpyxl-x.x.x PyGithub-x.x.x python-dotenv-x.x.x
```

### Step 3: Verify Dependencies

Test import of all required packages:

```bash
python -c "import gradio; print(f'Gradio: {gradio.__version__}')"
python -c "import pandas; print(f'Pandas: {pandas.__version__}')"
python -c "import openpyxl; print(f'Openpyxl: {openpyxl.__version__}')"
python -c "from github import Github; print('PyGithub: OK')"
python -c "from dotenv import load_dotenv; print('python-dotenv: OK')"
```

All commands should complete without errors.

### Step 4: Syntax Check

Verify the main application has no syntax errors:

```bash
python -m py_compile excelbot_chat.py
```

No output means success!

### Step 5: Test Launch (Quick Test)

Start the application with a timeout:

```bash
timeout 10 python excelbot_chat.py
# or on Windows (Ctrl+C after a few seconds):
python excelbot_chat.py
```

Expected output should include:
```
Running on local URL:  http://127.0.0.1:7860
```

Press Ctrl+C to stop.

### Step 6: Full Application Test

1. **Start the application:**
```bash
python excelbot_chat.py
```

2. **Open browser to:** `http://localhost:7860`

3. **Verify UI loads** with all tabs visible:
   - [ ] VBA Generator tab
   - [ ] Excel Analyzer tab
   - [ ] GitHub Integration tab
   - [ ] Help tab

4. **Test VBA Generator:**
   - [ ] Type "Sort my data" in the text box
   - [ ] Click "Generate VBA Macro"
   - [ ] Verify VBA code appears in the output
   - [ ] Code should contain "Sub SortData()"

5. **Test Excel Analyzer:**
   - [ ] Upload sample1.xlsx or sample2.xlsx
   - [ ] Click "Analyze File"
   - [ ] Verify file statistics appear
   - [ ] Should show rows, columns, and data preview

6. **Check Help Tab:**
   - [ ] Click Help tab
   - [ ] Verify documentation is visible
   - [ ] Should show usage instructions

### Step 7: Optional - GitHub Integration Test

⚠️ Only if you want to use GitHub features:

1. **Get GitHub Token:**
   - Visit: https://github.com/settings/tokens
   - Generate new token with `repo` permissions

2. **Set Environment Variable:**
```bash
export GITHUB_TOKEN=your_token_here  # Linux/Mac
set GITHUB_TOKEN=your_token_here     # Windows
```

3. **Restart Application**

4. **Test Push:**
   - Generate a VBA macro
   - Go to GitHub Integration tab
   - Enter repository name (format: username/repo)
   - Enter file name (e.g., test_macro.bas)
   - Paste the code
   - Click "Push to GitHub"
   - Verify success message appears

## 🐛 Common Issues and Solutions

### Issue 1: "Python not found"
**Solution:**
- Install Python from https://python.org
- On Windows, ensure "Add Python to PATH" was checked during installation
- Restart your terminal/command prompt

### Issue 2: "pip not found"
**Solution:**
```bash
python -m ensurepip --default-pip
python -m pip install --upgrade pip
```

### Issue 3: "Permission denied" (Linux/Mac)
**Solution:**
```bash
chmod +x setup.sh
# or use pip with --user flag:
pip install --user -r requirements.txt
```

### Issue 4: "Port 7860 already in use"
**Solution:**
```bash
# Find process using port 7860
lsof -i :7860  # Linux/Mac
netstat -ano | findstr :7860  # Windows

# Kill the process or change port in excelbot_chat.py:
# demo.launch(server_port=8080)
```

### Issue 5: "Module not found" errors
**Solution:**
```bash
# Reinstall all dependencies
pip uninstall -y gradio pandas openpyxl PyGithub python-dotenv
pip install -r requirements.txt
```

### Issue 6: "Gradio connection error"
**Solution:**
- Check firewall settings
- Try accessing via `http://127.0.0.1:7860` instead of localhost
- Ensure no VPN is blocking local connections

### Issue 7: Excel file upload fails
**Solution:**
- Ensure file is valid .xlsx or .xls format
- Check file size (should be under 10MB)
- Try with provided sample files first
- Ensure file is not corrupted or password-protected

## ✅ Success Indicators

Your installation is successful if:

- ✓ All dependencies installed without errors
- ✓ Application starts without errors
- ✓ Browser UI loads correctly
- ✓ All four tabs are visible and clickable
- ✓ VBA Generator produces code output
- ✓ Excel Analyzer processes sample files
- ✓ No console errors appear
- ✓ Application responds to user input

## 🧪 Advanced Testing

### Test 1: Performance Test
```bash
# Test with a larger Excel file
python -c "
import pandas as pd
import numpy as np
df = pd.DataFrame(np.random.randn(1000, 10))
df.to_excel('test_large.xlsx', index=False)
print('Created test_large.xlsx with 1000 rows')
"
```

Upload this file and verify analysis completes in under 5 seconds.

### Test 2: Template Matching
Test each template keyword:
- "sort data" → Should return SortData macro
- "filter sheet" → Should return FilterData macro
- "highlight duplicates" → Should return HighlightDuplicates macro
- "remove duplicates" → Should return RemoveDuplicates macro
- "format table" → Should return FormatTable macro
- "sum columns" → Should return CalculateSums macro
- "create pivot" → Should return CreatePivotTable macro

### Test 3: Docker Installation (Optional)
```bash
# Build Docker image
docker build -t excelbotpro .

# Verify image created
docker images | grep excelbotpro

# Run container
docker run -p 7860:7860 excelbotpro

# Test access
curl http://localhost:7860
```

## 📊 Verification Checklist

Complete this checklist to confirm successful installation:

### Core Installation
- [ ] Python 3.7+ installed and verified
- [ ] pip working correctly
- [ ] All dependencies installed successfully
- [ ] No import errors when testing packages
- [ ] Main application passes syntax check

### Application Functionality
- [ ] Application starts without errors
- [ ] Web UI loads in browser
- [ ] All tabs are accessible
- [ ] VBA Generator produces code
- [ ] Excel Analyzer processes files
- [ ] Sample files load successfully
- [ ] Help documentation is visible

### Optional Features
- [ ] GitHub token configured (if using)
- [ ] GitHub integration tested (if using)
- [ ] Docker image builds (if using Docker)
- [ ] Docker container runs (if using Docker)

### Documentation
- [ ] README.md readable
- [ ] QUICKSTART.md reviewed
- [ ] DEPLOYMENT.md reviewed (if deploying)
- [ ] All .md files present

## 🎉 Installation Complete!

If you've completed all checks successfully, your installation is ready for production use!

### Next Steps:
1. Read QUICKSTART.md for usage examples
2. Try generating different types of VBA macros
3. Analyze your own Excel files
4. Configure GitHub integration (optional)
5. Read DEPLOYMENT.md if deploying to production

### Getting Help:
- 📖 Check README.md for detailed documentation
- 🐛 Report issues: https://github.com/mandarab76/ExcelBotPro/issues
- 💬 Ask questions: https://github.com/mandarab76/ExcelBotPro/discussions
- ❓ Use in-app Help tab

## 🔄 Re-verification

If you update or reinstall, re-run these verification steps:

```bash
# Quick verification script
python -c "
import sys
import gradio
import pandas
import openpyxl
from github import Github

print('✓ All imports successful')
print(f'✓ Python: {sys.version}')
print(f'✓ Gradio: {gradio.__version__}')
print(f'✓ Pandas: {pandas.__version__}')
print(f'✓ Openpyxl: {openpyxl.__version__}')
print('✓ PyGithub: OK')
print('\\n🎉 Installation verified!')
"
```

---

**Happy Automating with ExcelBot Pro! 🚀**
