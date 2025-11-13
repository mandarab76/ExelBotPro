#!/bin/bash

# ExcelBot Pro Setup Script
# This script helps set up the ExcelBot Pro environment

echo "🚀 ExcelBot Pro Setup"
echo "===================="
echo ""

# Check Python version
echo "Checking Python version..."
python_version=$(python3 --version 2>&1 | awk '{print $2}')
if [ $? -ne 0 ]; then
    echo "❌ Python 3 is not installed. Please install Python 3.7 or higher."
    exit 1
fi
echo "✅ Python version: $python_version"
echo ""

# Create virtual environment
echo "Creating virtual environment..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo "✅ Virtual environment created"
else
    echo "⚠️  Virtual environment already exists"
fi
echo ""

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate
echo "✅ Virtual environment activated"
echo ""

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip > /dev/null 2>&1
echo "✅ pip upgraded"
echo ""

# Install dependencies
echo "Installing dependencies..."
pip install -r requirements.txt
if [ $? -eq 0 ]; then
    echo "✅ Dependencies installed successfully"
else
    echo "❌ Failed to install dependencies"
    exit 1
fi
echo ""

# Check for .env file
if [ ! -f ".env" ]; then
    echo "📝 Creating .env file from template..."
    cp .env.example .env
    echo "✅ .env file created"
    echo "⚠️  Please edit .env file and add your GITHUB_TOKEN"
else
    echo "✅ .env file already exists"
fi
echo ""

echo "✨ Setup complete!"
echo ""
echo "Next steps:"
echo "1. Edit .env file and add your GITHUB_TOKEN"
echo "2. Activate virtual environment: source venv/bin/activate"
echo "3. Run the application: python excelbot_chat.py"
echo ""
