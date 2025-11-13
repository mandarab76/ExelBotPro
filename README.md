
# ExcelBot Pro

A chatbot interface for automating Excel tasks using VBA, with GitHub integration and Excel file upload support.

## Features
- 🤖 **Intelligent VBA Generation**: Generate VBA macros from natural language descriptions
- 📊 **Excel File Analysis**: Upload and analyze Excel files to understand their structure
- 🔗 **GitHub Integration**: Push generated VBA code directly to GitHub repositories
- 🎨 **Modern UI**: Clean, tabbed interface built with Gradio

## Supported VBA Tasks

The bot can generate VBA code for common Excel tasks including:
- **Formatting**: Bold headers, highlight cells, color formatting
- **Data Operations**: Sort, filter, find/replace
- **Calculations**: Sum, totals, formulas
- **Data Management**: Delete empty rows/columns, copy/paste
- **Visualization**: Create charts and graphs

## Setup Instructions

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   Or manually:
   ```bash
   pip install gradio PyGithub pandas openpyxl
   ```

2. (Optional) Set your GitHub token for repository integration:
   ```bash
   export GITHUB_TOKEN=your_personal_access_token
   ```
   Note: The app works without GitHub token, but you won't be able to push to repositories.

3. Run the chatbot:
   ```bash
   python excelbot_chat.py
   ```
   Or with Python 3:
   ```bash
   python3 excelbot_chat.py
   ```

4. Open your browser to `http://localhost:7860` (or the URL shown in the terminal)

## Usage

1. **Generate VBA Macro**: 
   - Go to the "Generate VBA Macro" tab
   - Describe your Excel task in plain English (e.g., "Sort data by column A and make headers bold")
   - Click "Generate VBA Macro" to see the VBA code

2. **Analyze Excel Files**:
   - Go to the "Excel File Analysis" tab
   - Upload an Excel file (.xlsx or .xls)
   - View the file structure, row/column counts, and sheet names

3. **Push to GitHub**:
   - Generate a VBA macro first
   - Go to the "Push to GitHub" tab
   - Enter your repository name and desired file name
   - Click "Push to GitHub" to upload the code

## Example Tasks

- "Sort the data by the first column"
- "Make all headers bold and highlight them"
- "Calculate the sum of column B"
- "Delete all empty rows"
- "Create a chart from the data"
- "Filter the data to show only values greater than 100"

## Sample Excel Files

Sample Excel files (`sample1.xlsx`, `sample2.xlsx`) are included for testing the upload feature.

## Requirements

- Python 3.7+
- See `requirements.txt` for package dependencies

## License

See LICENSE file for details.
