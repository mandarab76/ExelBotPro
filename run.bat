@echo off
REM ExcelBot Pro - Launch Script for Windows
REM This script helps launch ExcelBot Pro with proper environment setup

echo 🚀 Starting ExcelBot Pro...

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Error: Python is not installed. Please install Python 3.8 or higher.
    pause
    exit /b 1
)

REM Display Python version
for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYTHON_VERSION=%%v
echo ✓ Python version: %PYTHON_VERSION%

REM Check if virtual environment exists
if exist "venv\Scripts\activate.bat" (
    echo ✓ Virtual environment found. Activating...
    call venv\Scripts\activate.bat
) else (
    echo ⚠ Virtual environment not found. Using system Python.
)

REM Check if dependencies are installed
python -c "import gradio" >nul 2>&1
if errorlevel 1 (
    echo ⚠ Dependencies not found. Installing...
    python -m pip install -r requirements.txt
)

REM Check for environment variables
if "%GITHUB_TOKEN%"=="" (
    echo ⚠ GITHUB_TOKEN not set. GitHub features will be disabled.
)

if "%OPENAI_API_KEY%"=="" (
    echo ⚠ OPENAI_API_KEY not set. Using template-based VBA generation.
)

REM Launch the application
echo.
echo 🎯 Launching ExcelBot Pro...
echo 📱 Open your browser at http://localhost:7860
echo.
echo Press Ctrl+C to stop the server
echo.

python excelbot_chat.py

pause
