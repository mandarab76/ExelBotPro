#!/bin/bash

# ExcelBot Pro - Launch Script
# This script helps launch ExcelBot Pro with proper environment setup

echo "🚀 Starting ExcelBot Pro..."

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo "❌ Error: Python is not installed. Please install Python 3.8 or higher."
        exit 1
    else
        PYTHON_CMD=python
    fi
else
    PYTHON_CMD=python3
fi

# Check Python version
PYTHON_VERSION=$($PYTHON_CMD --version 2>&1 | awk '{print $2}')
echo "✓ Python version: $PYTHON_VERSION"

# Check if virtual environment exists
if [ -d "venv" ]; then
    echo "✓ Virtual environment found. Activating..."
    source venv/bin/activate
else
    echo "⚠ Virtual environment not found. Using system Python."
fi

# Check if dependencies are installed
if ! $PYTHON_CMD -c "import gradio" 2>/dev/null; then
    echo "⚠ Dependencies not found. Installing..."
    $PYTHON_CMD -m pip install -r requirements.txt
fi

# Check for environment variables
if [ -z "$GITHUB_TOKEN" ]; then
    echo "⚠ GITHUB_TOKEN not set. GitHub features will be disabled."
fi

if [ -z "$OPENAI_API_KEY" ]; then
    echo "⚠ OPENAI_API_KEY not set. Using template-based VBA generation."
fi

# Launch the application
echo ""
echo "🎯 Launching ExcelBot Pro..."
echo "📱 Open your browser at http://localhost:7860"
echo ""
echo "Press Ctrl+C to stop the server"
echo ""

$PYTHON_CMD excelbot_chat.py
