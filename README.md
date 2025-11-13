# ExcelBot Pro

A powerful chatbot interface for automating Excel tasks using VBA, with GitHub integration and Excel file upload support.

## 🚀 Features

- **VBA Macro Generation**: Generate VBA code from natural language descriptions
- **Excel File Upload**: Upload and analyze Excel files (.xls/.xlsx)
- **GitHub Integration**: Push generated VBA macros directly to GitHub repositories
- **Smart Pattern Recognition**: Automatically detects common Excel tasks (formatting, sorting, filtering, etc.)
- **Modern Web Interface**: Beautiful Gradio-based UI for easy interaction

## 📋 Prerequisites

- Python 3.7 or higher
- GitHub Personal Access Token (for GitHub integration feature)

## 🛠️ Installation

### Option 1: Using pip

1. Clone or download this repository:
   ```bash
   git clone https://github.com/mandarab76/ExcelBotPro.git
   cd ExcelBotPro
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up environment variables:
   ```bash
   # Linux/Mac
   export GITHUB_TOKEN=your_personal_access_token
   
   # Windows
   set GITHUB_TOKEN=your_personal_access_token
   ```
   
   Or create a `.env` file (see `.env.example` for template)

4. Run the application:
   ```bash
   python excelbot_chat.py
   ```

### Option 2: Using virtual environment (Recommended)

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Linux/Mac:
source venv/bin/activate
# On Windows:
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Set GitHub token
export GITHUB_TOKEN=your_personal_access_token

# Run the application
python excelbot_chat.py
```

## 🔑 Getting a GitHub Personal Access Token

1. Go to GitHub Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Click "Generate new token (classic)"
3. Give it a name (e.g., "ExcelBot Pro")
4. Select scopes: `repo` (full control of private repositories)
5. Click "Generate token"
6. Copy the token and use it as `GITHUB_TOKEN`

## 📖 Usage

### Starting the Application

After installation, run:
```bash
python excelbot_chat.py
```

The application will start on `http://localhost:7860` by default.

### Using the VBA Generator

1. Navigate to the "VBA Generator" tab
2. Enter a task description (e.g., "Format selected cells as bold with gray background")
3. Click "Generate VBA Macro"
4. Copy the generated VBA code

### Uploading Excel Files

1. Navigate to the "Excel Upload" tab
2. Click "Upload" and select an Excel file (.xls or .xlsx)
3. View file information including sheet names and structure

### Pushing to GitHub

1. Generate VBA code using the VBA Generator
2. Navigate to the "GitHub Integration" tab
3. Enter your GitHub repository name
4. Enter a file name (will be saved as .bas file)
5. Click "Push to GitHub"

## 🎯 Supported Task Types

The bot recognizes and generates code for:

- **Formatting**: Cell formatting, styles, colors
- **Calculations**: Sum, totals, formulas
- **Sorting**: Data sorting operations
- **Filtering**: AutoFilter applications
- **Data Manipulation**: Copy, delete, remove operations
- **Generic Tasks**: Template-based code for custom tasks

## 📁 Project Structure

```
ExcelBotPro/
├── excelbot_chat.py      # Main application file
├── requirements.txt      # Python dependencies
├── README.md            # This file
├── LICENSE              # MIT License
├── .env.example         # Environment variables template
├── .gitignore           # Git ignore rules
└── sample*.xlsx         # Sample Excel files for testing
```

## ⚙️ Configuration

### Environment Variables

- `GITHUB_TOKEN`: Your GitHub Personal Access Token (required for GitHub features)
- `SERVER_NAME`: Server hostname (default: "0.0.0.0")
- `SERVER_PORT`: Server port (default: 7860)
- `SHARE`: Enable Gradio sharing (default: "False")

### Example .env file

```env
GITHUB_TOKEN=ghp_your_token_here
SERVER_NAME=0.0.0.0
SERVER_PORT=7860
SHARE=False
```

## 🐛 Troubleshooting

### GitHub Integration Not Working

- Ensure `GITHUB_TOKEN` is set correctly
- Verify the token has `repo` scope
- Check that the repository name is correct and exists
- Ensure you have write access to the repository

### Excel File Upload Issues

- Verify the file is a valid .xls or .xlsx format
- Check file permissions
- Ensure the file is not corrupted

### Port Already in Use

Change the port by setting the `SERVER_PORT` environment variable:
```bash
export SERVER_PORT=7861
python excelbot_chat.py
```

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📧 Support

For issues and questions, please open an issue on the GitHub repository.

## 🙏 Acknowledgments

- Built with [Gradio](https://www.gradio.app/)
- GitHub integration via [PyGithub](https://github.com/PyGithub/PyGithub)
- Excel processing with [openpyxl](https://openpyxl.readthedocs.io/) and [pandas](https://pandas.pydata.org/)
