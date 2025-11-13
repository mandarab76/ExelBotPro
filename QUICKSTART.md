# ExcelBot Pro - Quick Start Guide

## 🚀 Get Started in 5 Minutes

### Step 1: Extract the Package
Extract `ExcelBotPro.zip` to your desired location.

### Step 2: Install Dependencies

**On Linux/Mac:**
```bash
chmod +x setup.sh
./setup.sh
```

**On Windows:**
Double-click `setup.bat` or run it from Command Prompt.

**Manual Installation:**
```bash
pip install -r requirements.txt
```

### Step 3: Configure GitHub Token (Optional)

1. Get a GitHub Personal Access Token:
   - Go to https://github.com/settings/tokens
   - Click "Generate new token (classic)"
   - Select `repo` scope
   - Copy the token

2. Set the token:
   ```bash
   # Linux/Mac
   export GITHUB_TOKEN=your_token_here
   
   # Windows
   set GITHUB_TOKEN=your_token_here
   ```
   
   Or edit `.env` file and add:
   ```
   GITHUB_TOKEN=your_token_here
   ```

### Step 4: Run the Application

```bash
python excelbot_chat.py
```

### Step 5: Open in Browser

Navigate to: `http://localhost:7860`

## 📖 Basic Usage

1. **Generate VBA Code:**
   - Go to "VBA Generator" tab
   - Enter task description (e.g., "Format cells as bold")
   - Click "Generate VBA Macro"
   - Copy the generated code

2. **Upload Excel Files:**
   - Go to "Excel Upload" tab
   - Upload your Excel file
   - View file structure and sheet information

3. **Push to GitHub:**
   - Generate VBA code first
   - Go to "GitHub Integration" tab
   - Enter repository name and file name
   - Click "Push to GitHub"

## 🎯 Example Tasks

Try these example tasks:
- "Format selected cells as bold with gray background"
- "Calculate sum of selected range"
- "Sort data in ascending order"
- "Apply autofilter to selected range"
- "Delete selected rows"

## ❓ Need Help?

See the full [README.md](README.md) for detailed documentation.

## 🐛 Troubleshooting

**Port already in use?**
```bash
export SERVER_PORT=7861
python excelbot_chat.py
```

**GitHub not working?**
- Check your GITHUB_TOKEN is set correctly
- Verify token has `repo` scope
- Ensure repository exists and you have access

**Import errors?**
```bash
pip install --upgrade -r requirements.txt
```
