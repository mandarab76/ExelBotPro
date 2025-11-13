#!/bin/bash
# ExcelBot Pro - Quick Setup Script
# This script automates the installation process

set -e

echo "🤖 ExcelBot Pro - Installation Script"
echo "======================================"
echo ""

# Color codes
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Check if Python is installed
echo "Checking Python installation..."
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}Python 3 is not installed. Please install Python 3.7+ first.${NC}"
    exit 1
fi

PYTHON_VERSION=$(python3 --version | cut -d' ' -f2 | cut -d'.' -f1,2)
echo -e "${GREEN}✓ Python $PYTHON_VERSION found${NC}"

# Check pip
echo "Checking pip installation..."
if ! command -v pip3 &> /dev/null; then
    echo -e "${YELLOW}pip3 not found. Installing...${NC}"
    sudo apt-get install python3-pip -y
fi
echo -e "${GREEN}✓ pip found${NC}"

# Create virtual environment (optional)
read -p "Do you want to create a virtual environment? (recommended) [y/N]: " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
    source venv/bin/activate
    echo -e "${GREEN}✓ Virtual environment created and activated${NC}"
fi

# Install dependencies
echo ""
echo "Installing dependencies..."
pip3 install -r requirements.txt
echo -e "${GREEN}✓ Dependencies installed${NC}"

# Setup environment variables
echo ""
if [ ! -f .env ]; then
    read -p "Do you want to configure GitHub integration? [y/N]: " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        cp .env.example .env
        echo -e "${YELLOW}Please edit .env file and add your GitHub token${NC}"
        echo "Get your token from: https://github.com/settings/tokens"
        read -p "Press Enter to continue..."
    fi
else
    echo -e "${GREEN}✓ .env file already exists${NC}"
fi

# Test the installation
echo ""
echo "Testing installation..."
python3 -c "import gradio; import pandas; from github import Github" 2>/dev/null
if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓ All dependencies working correctly${NC}"
else
    echo -e "${RED}✗ Some dependencies failed to import${NC}"
    exit 1
fi

# Success message
echo ""
echo -e "${GREEN}======================================"
echo "✓ Installation completed successfully!"
echo "======================================${NC}"
echo ""
echo "To start ExcelBot Pro:"
echo "  python3 excelbot_chat.py"
echo ""
echo "Then open your browser to:"
echo "  http://localhost:7860"
echo ""
echo "For help and documentation, visit:"
echo "  https://github.com/mandarab76/ExcelBotPro"
echo ""

# Ask if user wants to start now
read -p "Do you want to start ExcelBot Pro now? [y/N]: " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo "Starting ExcelBot Pro..."
    python3 excelbot_chat.py
fi
