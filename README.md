
# ExcelBot Pro

A chatbot interface for automating Excel tasks using VBA, with GitHub integration and Excel file upload support.

## Features
- **Intelligent VBA Generation**: Generate VBA macros from natural language descriptions
  - Supports common tasks: formatting, sorting, filtering, calculations, data manipulation
  - Includes error handling and best practices
- **Excel File Analysis**: Upload and analyze Excel files (.xls/.xlsx)
  - Automatically detects sheets, rows, and columns
  - Provides detailed file structure information
- **GitHub Integration**: Push generated VBA code to GitHub repositories
  - Create or update files in repositories
  - Automatic .bas file extension handling

## Setup Instructions
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   Or manually:
   ```bash
   pip install gradio PyGithub pandas openpyxl
   ```

2. (Optional) Set your GitHub token for GitHub integration:
   ```bash
   export GITHUB_TOKEN=your_personal_access_token
   ```
   Note: The app works without GitHub token, but GitHub push functionality will be disabled.

3. Run the chatbot:
   ```bash
   python excelbot_chat.py
   ```

4. Open your browser to `http://localhost:7860` (or the URL shown in the terminal)

## Usage Examples

### Generate VBA Macros
- "Format the header row with blue background"
- "Sort data by column A"
- "Delete empty rows"
- "Apply autofilter to the data"
- "Calculate sum of column A"

### Excel File Upload
Upload any `.xlsx` or `.xls` file to get detailed information about its structure, including:
- Number of sheets
- Sheet names
- Row and column counts for each sheet

### Push to GitHub
1. Generate a VBA macro
2. Enter your repository name
3. Specify a file name (automatically adds .bas extension)
4. Click "Push to GitHub"

## Sample Excel Files
Sample Excel files (`sample1.xlsx`, `sample2.xlsx`) are included for testing the upload feature.
