#!/bin/bash
# ExcelBot Pro - Mobile Access Launch Script

echo "╔══════════════════════════════════════════════════════════╗"
echo "║                                                          ║"
echo "║     📱 ExcelBot Pro - Mobile Access Launch 📱           ║"
echo "║                                                          ║"
echo "╚══════════════════════════════════════════════════════════╝"
echo ""
echo "🚀 Starting NSE Stock Market Analysis Suite..."
echo "📱 Creating PUBLIC share link for mobile access..."
echo ""
echo "⏳ Please wait 10-15 seconds for the link to generate..."
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "📋 You will see TWO URLs:"
echo "   1. Local URL (http://127.0.0.1:7860) - for desktop"
echo "   2. Public URL (https://xxxxx.gradio.live) - USE THIS ON YOUR PHONE!"
echo ""
echo "📱 COPY the https://xxxxx.gradio.live link and open it on your phone!"
echo ""
echo "⏹️  To stop: Press Ctrl+C"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

cd /workspace
python3 excelbot_chat.py
