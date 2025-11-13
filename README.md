
# ExcelBot Pro 🤖

A powerful chatbot interface for automating Excel tasks using VBA, with GitHub integration and Excel file analysis.

## Features ✨
- **VBA Generation**: Generate VBA macros from natural language descriptions
- **Smart Templates**: Pre-built macros for common tasks (format, sort, filter, duplicates, totals, charts)
- **Excel Analysis**: Upload and analyze Excel files with pandas
- **GitHub Integration**: Push generated VBA code directly to your GitHub repositories
- **Modern UI**: Beautiful Gradio-based interface with easy-to-use controls

## Supported Tasks 📋
- **Format**: Auto-format data with borders, fonts, and auto-fit columns
- **Sort**: Sort data by columns
- **Filter**: Apply AutoFilter to data ranges
- **Remove Duplicates**: Clean up duplicate rows
- **Calculate Totals**: Add sum rows automatically
- **Create Charts**: Generate column charts from data

## Setup Instructions 🚀

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Set your GitHub token (optional for GitHub features):**
   ```bash
   export GITHUB_TOKEN=your_personal_access_token
   ```

3. **Run the chatbot:**
   ```bash
   python excelbot_chat.py
   ```

4. **Open your browser** and navigate to the local URL shown in the terminal (typically `http://127.0.0.1:7860`)

## Usage Examples 💡

1. **Generate a VBA macro:**
   - Type: "format my data" or "sort by first column"
   - Click "Generate VBA Macro"
   - Copy the generated VBA code

2. **Analyze an Excel file:**
   - Upload `sample1.xlsx` or `sample2.xlsx` (included in the project)
   - View file statistics and column information

3. **Push to GitHub:**
   - Generate a VBA macro
   - Enter your repository name (e.g., "username/repo-name")
   - Enter a file name (e.g., "my_macro.vba")
   - Click "Push to GitHub"

## Sample Excel Files 📂
Sample files `sample1.xlsx` and `sample2.xlsx` are included in the root directory for testing the upload and analysis features.

## Requirements 📦
- Python 3.7+
- gradio
- PyGithub
- pandas
- openpyxl

## License
MIT License - see LICENSE file for details
