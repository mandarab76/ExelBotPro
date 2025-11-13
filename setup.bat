@echo off
REM ExcelBot Pro Setup Script for Windows
REM This script helps set up the ExcelBot Pro environment

echo 🚀 ExcelBot Pro Setup
echo ====================
echo.

REM Check Python version
echo Checking Python version...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed. Please install Python 3.7 or higher.
    exit /b 1
)
python --version
echo ✅ Python found
echo.

REM Create virtual environment
echo Creating virtual environment...
if not exist "venv" (
    python -m venv venv
    echo ✅ Virtual environment created
) else (
    echo ⚠️  Virtual environment already exists
)
echo.

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
echo ✅ Virtual environment activated
echo.

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip >nul 2>&1
echo ✅ pip upgraded
echo.

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo ❌ Failed to install dependencies
    exit /b 1
)
echo ✅ Dependencies installed successfully
echo.

REM Check for .env file
if not exist ".env" (
    echo 📝 Creating .env file from template...
    copy .env.example .env >nul
    echo ✅ .env file created
    echo ⚠️  Please edit .env file and add your GITHUB_TOKEN
) else (
    echo ✅ .env file already exists
)
echo.

echo ✨ Setup complete!
echo.
echo Next steps:
echo 1. Edit .env file and add your GITHUB_TOKEN
echo 2. Activate virtual environment: venv\Scripts\activate
echo 3. Run the application: python excelbot_chat.py
echo.

pause
