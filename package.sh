#!/bin/bash

# ExcelBot Pro - Package Script
# Creates a ready-to-distribute zip file

PACKAGE_NAME="ExcelBotPro-v1.0.0"
TEMP_DIR="temp_package"
ZIP_FILE="${PACKAGE_NAME}.zip"

echo "📦 Packaging ExcelBot Pro..."

# Clean up previous builds
rm -rf "$TEMP_DIR" "$ZIP_FILE"

# Create temporary directory
mkdir -p "$TEMP_DIR"

# Copy essential files
echo "📋 Copying files..."
cp excelbot_chat.py "$TEMP_DIR/"
cp config.py "$TEMP_DIR/"
cp requirements.txt "$TEMP_DIR/"
cp setup.py "$TEMP_DIR/"
cp README.md "$TEMP_DIR/"
cp INSTALL.md "$TEMP_DIR/"
cp DEPLOYMENT.md "$TEMP_DIR/"
cp PACKAGE_INFO.md "$TEMP_DIR/"
cp LICENSE "$TEMP_DIR/"
cp .gitignore "$TEMP_DIR/"
cp run.sh "$TEMP_DIR/"
cp run.bat "$TEMP_DIR/"

# Copy sample files
if [ -f "sample1.xlsx" ]; then
    cp sample1.xlsx "$TEMP_DIR/"
fi
if [ -f "sample2.xlsx" ]; then
    cp sample2.xlsx "$TEMP_DIR/"
fi

# Make scripts executable
chmod +x "$TEMP_DIR/run.sh"

# Create zip file
echo "🗜️  Creating zip archive..."
cd "$TEMP_DIR"
zip -r "../$ZIP_FILE" . > /dev/null
cd ..

# Clean up
rm -rf "$TEMP_DIR"

echo "✅ Package created: $ZIP_FILE"
echo "📊 Package size: $(du -h $ZIP_FILE | cut -f1)"
echo ""
echo "🎉 Ready to distribute!"
