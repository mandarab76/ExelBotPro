# Installation Guide - ExcelBot Pro

## Quick Start

### Step 1: Download and Extract

1. Download the ExcelBot Pro zip file
2. Extract it to your desired location
3. Open a terminal/command prompt in the extracted folder

### Step 2: Install Python Dependencies

```bash
pip install -r requirements.txt
```

### Step 3: Run the Application

```bash
python excelbot_chat.py
```

The application will start at `http://localhost:7860`

## Detailed Installation

### Windows

1. **Install Python 3.8+**
   - Download from [python.org](https://www.python.org/downloads/)
   - Make sure to check "Add Python to PATH" during installation

2. **Open Command Prompt**
   - Press `Win + R`, type `cmd`, press Enter
   - Navigate to the ExcelBot Pro folder:
     ```cmd
     cd path\to\ExcelBotPro
     ```

3. **Create Virtual Environment (Optional but Recommended)**
   ```cmd
   python -m venv venv
   venv\Scripts\activate
   ```

4. **Install Dependencies**
   ```cmd
   pip install -r requirements.txt
   ```

5. **Run the Application**
   ```cmd
   python excelbot_chat.py
   ```

### macOS/Linux

1. **Install Python 3.8+**
   ```bash
   # macOS (using Homebrew)
   brew install python3
   
   # Linux (Ubuntu/Debian)
   sudo apt-get update
   sudo apt-get install python3 python3-pip python3-venv
   ```

2. **Open Terminal**
   - Navigate to the ExcelBot Pro folder:
     ```bash
     cd /path/to/ExcelBotPro
     ```

3. **Create Virtual Environment**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

4. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

5. **Run the Application**
   ```bash
   python excelbot_chat.py
   ```

## Optional: Configure API Keys

### GitHub Token (Optional)

1. Go to [GitHub Settings](https://github.com/settings/tokens)
2. Click "Generate new token" → "Generate new token (classic)"
3. Select `repo` scope
4. Copy the token
5. Set environment variable:
   ```bash
   # Windows
   set GITHUB_TOKEN=your_token_here
   
   # macOS/Linux
   export GITHUB_TOKEN=your_token_here
   ```

### OpenAI API Key (Optional)

1. Sign up at [OpenAI Platform](https://platform.openai.com/)
2. Navigate to API Keys
3. Create a new API key
4. Set environment variable:
   ```bash
   # Windows
   set OPENAI_API_KEY=your_key_here
   
   # macOS/Linux
   export OPENAI_API_KEY=your_key_here
   ```

## Troubleshooting

### Issue: "pip: command not found"

**Solution**: Install pip or use `python -m pip` instead:
```bash
python -m pip install -r requirements.txt
```

### Issue: "ModuleNotFoundError"

**Solution**: Make sure all dependencies are installed:
```bash
pip install --upgrade -r requirements.txt
```

### Issue: Port 7860 already in use

**Solution**: Change the port in `excelbot_chat.py`:
```python
demo.launch(server_port=7861)  # Use a different port
```

### Issue: Permission denied (macOS/Linux)

**Solution**: Use `python3` instead of `python`:
```bash
python3 excelbot_chat.py
```

## Verification

After installation, verify everything works:

1. Run `python excelbot_chat.py`
2. Open `http://localhost:7860` in your browser
3. You should see the ExcelBot Pro interface
4. Try generating a simple macro (e.g., "Format column A as currency")

## Next Steps

- Read the [README.md](README.md) for usage instructions
- Check out example tasks in the application
- Upload sample Excel files to test file analysis

## Support

If you encounter issues:
1. Check the [GitHub Issues](https://github.com/mandarab76/ExcelBotPro/issues)
2. Ensure Python version is 3.8 or higher: `python --version`
3. Verify all dependencies are installed: `pip list`
