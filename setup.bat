@echo off
REM ExcelBot Pro - Quick Setup Script for Windows
REM This script automates the installation process

echo ========================================
echo ExcelBot Pro - Installation Script
echo ========================================
echo.

REM Check if Python is installed
echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version') do set PYTHON_VERSION=%%i
echo [OK] Python %PYTHON_VERSION% found
echo.

REM Check pip
echo Checking pip installation...
pip --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip is not installed
    echo Installing pip...
    python -m ensurepip --default-pip
)
echo [OK] pip found
echo.

REM Create virtual environment (optional)
set /p CREATE_VENV="Do you want to create a virtual environment? (recommended) [y/N]: "
if /i "%CREATE_VENV%"=="y" (
    echo Creating virtual environment...
    python -m venv venv
    call venv\Scripts\activate.bat
    echo [OK] Virtual environment created and activated
    echo.
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)
echo [OK] Dependencies installed
echo.

REM Setup environment variables
if not exist .env (
    set /p SETUP_GITHUB="Do you want to configure GitHub integration? [y/N]: "
    if /i "%SETUP_GITHUB%"=="y" (
        copy .env.example .env
        echo [INFO] Please edit .env file and add your GitHub token
        echo Get your token from: https://github.com/settings/tokens
        pause
    )
) else (
    echo [OK] .env file already exists
)
echo.

REM Test the installation
echo Testing installation...
python -c "import gradio; import pandas; from github import Github" 2>nul
if errorlevel 1 (
    echo [ERROR] Some dependencies failed to import
    pause
    exit /b 1
)
echo [OK] All dependencies working correctly
echo.

REM Success message
echo ========================================
echo Installation completed successfully!
echo ========================================
echo.
echo To start ExcelBot Pro:
echo   python excelbot_chat.py
echo.
echo Then open your browser to:
echo   http://localhost:7860
echo.
echo For help and documentation, visit:
echo   https://github.com/mandarab76/ExcelBotPro
echo.

REM Ask if user wants to start now
set /p START_NOW="Do you want to start ExcelBot Pro now? [y/N]: "
if /i "%START_NOW%"=="y" (
    echo Starting ExcelBot Pro...
    python excelbot_chat.py
) else (
    pause
)
