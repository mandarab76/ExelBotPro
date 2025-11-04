
# ExcelBot Pro

A chatbot interface for automating Excel tasks using VBA, with GitHub integration and Excel file upload support.

## Features
- Generate VBA macros from natural language
- Upload Excel files (.xls/.xlsx)
- Push generated code to GitHub repositories

## Setup Instructions
1. Install dependencies:
   ```bash
   pip install gradio PyGithub
   ```

2. Set your GitHub token:
   ```bash
   export GITHUB_TOKEN=your_personal_access_token
   ```

3. Run the chatbot:
   ```bash
   python excelbot_chat.py
   ```

## Sample Excel Files
Included in `sample_excel_files/` for testing the upload feature.
