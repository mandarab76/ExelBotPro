# 🤖 ExcelBot Pro - VBA Automation Suite

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Gradio](https://img.shields.io/badge/Gradio-4.0%2B-orange)

**A professional tool for generating VBA macros, processing Excel files, and managing code with GitHub integration**

[Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Documentation](#-documentation) • [License](#-license)

</div>

---

## 📋 Table of Contents
- [Overview](#-overview)
- [Features](#-features)
- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Usage](#-usage)
- [VBA Templates](#-vba-templates)
- [GitHub Integration](#-github-integration)
- [Screenshots](#-screenshots)
- [Contributing](#-contributing)
- [License](#-license)
- [Support](#-support)

## 🎯 Overview

ExcelBot Pro is an advanced automation tool that bridges the gap between natural language and Excel VBA programming. Whether you're a beginner or an expert, ExcelBot Pro helps you:

- 🚀 **Generate VBA Macros** from simple text descriptions
- 📊 **Analyze Excel Files** with detailed statistics and data quality reports
- 🐙 **Version Control** your VBA code with seamless GitHub integration
- ⚡ **Automate Common Tasks** like sorting, filtering, formatting, and more

## ✨ Features

### 🔧 VBA Macro Generator
- **Natural Language Processing**: Describe your task in plain English
- **Smart Template Matching**: Automatically selects the right macro template
- **7+ Pre-built Templates**: Sort, filter, highlight duplicates, format tables, and more
- **Custom Macro Generation**: Creates templates for unique tasks
- **Professional Code**: Production-ready VBA with error handling and comments

### 📊 Excel File Analyzer
- **File Upload Support**: Works with .xlsx and .xls files
- **Data Analytics**: Rows, columns, memory usage, and missing values
- **Statistical Analysis**: Comprehensive statistical summaries
- **Data Quality Reports**: Identify duplicates and data issues
- **Data Preview**: Quick view of your Excel data

### 🐙 GitHub Integration
- **Version Control**: Push macros directly to GitHub repositories
- **Auto-Update**: Detects existing files and updates them
- **Commit History**: Timestamped commits for tracking changes
- **Secure Authentication**: Uses Personal Access Tokens

### 🎨 Modern UI
- **Tabbed Interface**: Clean, organized workspace
- **Syntax Highlighting**: Easy-to-read code display
- **Responsive Design**: Works on desktop and mobile
- **Built-in Help**: Comprehensive documentation within the app

## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package manager)
- Git (optional, for cloning the repository)

### Step 1: Clone the Repository
```bash
git clone https://github.com/mandarab76/ExcelBotPro.git
cd ExcelBotPro
```

Or download the ZIP file and extract it.

### Step 2: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 3: Configure Environment (Optional)
For GitHub integration, create a `.env` file:
```bash
cp .env.example .env
```

Edit `.env` and add your GitHub Personal Access Token:
```
GITHUB_TOKEN=your_github_personal_access_token_here
```

**To get a GitHub token:**
1. Go to [GitHub Settings → Developer Settings → Personal Access Tokens](https://github.com/settings/tokens)
2. Click "Generate new token"
3. Select `repo` permissions
4. Copy the token to your `.env` file

### Step 4: Run the Application
```bash
python excelbot_chat.py
```

The application will start at `http://localhost:7860`

## ⚡ Quick Start

### 1. Generate Your First VBA Macro
1. Open the **VBA Generator** tab
2. Type a task description, e.g., "Sort my data by first column"
3. Click **Generate VBA Macro**
4. Copy the generated code to Excel's VBA editor (Alt+F11)

### 2. Analyze an Excel File
1. Open the **Excel Analyzer** tab
2. Upload your .xlsx or .xls file
3. Click **Analyze File** for quick stats
4. Click **Detailed Analysis** for comprehensive reports

### 3. Push Code to GitHub
1. Generate a VBA macro first
2. Open the **GitHub Integration** tab
3. Enter your repository name (e.g., `username/repo-name`)
4. Enter a file name (e.g., `my_macro.bas`)
5. Paste the code and click **Push to GitHub**

## 📖 Usage

### Using VBA Macros in Excel

Once you've generated a macro:

1. **Open Excel** with your data file
2. **Press Alt+F11** to open the VBA Editor
3. **Insert → Module** to create a new module
4. **Paste the generated code**
5. **Press F5** or click Run to execute the macro

### Command-Line Options

You can customize the server settings:

```python
# In excelbot_chat.py, modify the launch parameters:
demo.launch(
    server_name="0.0.0.0",  # Accept connections from any IP
    server_port=7860,        # Port number
    share=True,              # Create public share link
    show_error=True          # Show detailed errors
)
```

## 🛠️ VBA Templates

ExcelBot Pro includes these pre-built templates:

| Template | Description | Keywords |
|----------|-------------|----------|
| **Sort** | Sort data in ascending order | sort, order, arrange |
| **Filter** | Apply AutoFilter to data | filter, search, find |
| **Highlight Duplicates** | Highlight duplicate values with color | highlight, color, duplicate |
| **Remove Duplicates** | Delete duplicate rows | remove, delete, duplicate |
| **Format Table** | Apply professional table formatting | format, style, beautify |
| **Calculate Sums** | Add sum formulas to numeric columns | sum, total, calculate |
| **Create Pivot Table** | Generate pivot table from data | pivot |

### Example Tasks

Try these example descriptions:
- "Sort my data alphabetically"
- "Apply filters to my spreadsheet"
- "Highlight all duplicate entries"
- "Format my table with colors and borders"
- "Add sum totals to the bottom of numeric columns"
- "Create a pivot table from this data"

## 🐙 GitHub Integration

### Setup

1. **Create a Personal Access Token**:
   - Visit [GitHub Settings](https://github.com/settings/tokens)
   - Generate new token with `repo` scope
   - Copy the token

2. **Configure ExcelBot Pro**:
   ```bash
   export GITHUB_TOKEN=your_token_here
   ```
   Or add it to your `.env` file

3. **Create/Use a Repository**:
   - Create a repository on GitHub or use an existing one
   - Use format: `username/repository-name`

### Pushing Code

The GitHub integration will:
- ✅ Auto-detect if the file exists (updates it)
- ✅ Create new files if they don't exist
- ✅ Add timestamped commit messages
- ✅ Handle file extensions automatically (.bas, .vba, .txt)

## 📸 Screenshots

### VBA Macro Generator
Generate production-ready VBA code from natural language descriptions.

### Excel Analyzer
Comprehensive data analysis with statistics and quality reports.

### GitHub Integration
Push your macros to GitHub for version control and sharing.

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/AmazingFeature`
3. **Commit your changes**: `git commit -m 'Add some AmazingFeature'`
4. **Push to the branch**: `git push origin feature/AmazingFeature`
5. **Open a Pull Request**

### Ideas for Contributions
- Add more VBA templates
- Improve natural language processing
- Add support for other file formats
- Create video tutorials
- Improve documentation
- Add unit tests

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2025 Mandar Bahadarpurkar

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
```

## 💬 Support

### Documentation
- In-app help available in the **Help** tab
- [GitHub Issues](https://github.com/mandarab76/ExcelBotPro/issues) for bug reports
- [GitHub Discussions](https://github.com/mandarab76/ExcelBotPro/discussions) for questions

### Troubleshooting

**Issue**: Application won't start
- **Solution**: Ensure all dependencies are installed: `pip install -r requirements.txt`

**Issue**: GitHub push fails
- **Solution**: Check your `GITHUB_TOKEN` is set correctly and has `repo` permissions

**Issue**: Excel file upload fails
- **Solution**: Ensure the file is a valid .xlsx or .xls file and isn't corrupted

**Issue**: VBA macro doesn't work in Excel
- **Solution**: Enable macros in Excel: File → Options → Trust Center → Macro Settings

### System Requirements
- **OS**: Windows, macOS, or Linux
- **Python**: 3.7 or higher
- **Excel**: Microsoft Excel 2010 or later (for running VBA macros)
- **Browser**: Modern web browser (Chrome, Firefox, Safari, Edge)
- **RAM**: 2GB minimum, 4GB recommended
- **Disk Space**: 500MB for installation and dependencies

## 🎓 Advanced Usage

### Custom VBA Templates

You can add your own templates by editing the `VBA_TEMPLATES` dictionary in `excelbot_chat.py`:

```python
VBA_TEMPLATES["your_template"] = """Sub YourMacro()
    ' Your custom VBA code here
End Sub"""
```

### Running in Production

For production deployment:

```bash
# Install gunicorn
pip install gunicorn

# Run with multiple workers
gunicorn -w 4 -b 0.0.0.0:7860 excelbot_chat:demo
```

### Docker Deployment

Create a `Dockerfile`:

```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 7860
CMD ["python", "excelbot_chat.py"]
```

Build and run:
```bash
docker build -t excelbotpro .
docker run -p 7860:7860 -e GITHUB_TOKEN=your_token excelbotpro
```

## 🌟 Acknowledgments

- Built with [Gradio](https://gradio.app/) for the web interface
- Uses [PyGithub](https://github.com/PyGithub/PyGithub) for GitHub integration
- Excel processing powered by [pandas](https://pandas.pydata.org/) and [openpyxl](https://openpyxl.readthedocs.io/)

## 📊 Project Stats

- **Language**: Python
- **Framework**: Gradio
- **Version**: 1.0.0
- **Last Updated**: 2025-11-13

---

<div align="center">

**Made with ❤️ by [Mandar Bahadarpurkar](https://github.com/mandarab76)**

⭐ Star this repository if you find it helpful!

[Report Bug](https://github.com/mandarab76/ExcelBotPro/issues) • [Request Feature](https://github.com/mandarab76/ExcelBotPro/issues)

</div>
