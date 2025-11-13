# ExcelBot Pro

A powerful chatbot interface for automating Excel tasks using VBA macros, with GitHub integration and Excel file analysis capabilities.

![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Status](https://img.shields.io/badge/status-production--ready-brightgreen)

## 🚀 Features

- **🧠 AI-Powered VBA Generation**: Generate VBA macros from natural language descriptions
- **📊 Excel File Analysis**: Upload and analyze Excel files to understand their structure
- **🔗 GitHub Integration**: Push generated macros directly to your GitHub repositories
- **💾 Export Capabilities**: Download and save generated VBA code
- **🎨 Modern UI**: Clean, intuitive Gradio-based interface
- **⚡ Fast & Efficient**: Quick macro generation and file processing

## 📋 Prerequisites

- Python 3.8 or higher
- pip (Python package manager)
- (Optional) OpenAI API key for AI-powered generation
- (Optional) GitHub Personal Access Token for GitHub integration

## 🛠️ Installation

### 1. Clone the Repository

```bash
git clone https://github.com/mandarab76/ExcelBotPro.git
cd ExcelBotPro
```

### 2. Create Virtual Environment (Recommended)

```bash
python -m venv venv

# On Windows
venv\Scripts\activate

# On macOS/Linux
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure Environment Variables

Create a `.env` file in the project root (optional) or set environment variables:

```bash
# For GitHub integration (optional)
export GITHUB_TOKEN=your_github_personal_access_token

# For AI-powered VBA generation (optional)
export OPENAI_API_KEY=your_openai_api_key
```

**Note**: The application works without these keys, but with limited functionality:
- Without OpenAI API key: Uses template-based VBA generation
- Without GitHub token: GitHub push features will be disabled

## 🎯 Usage

### Basic Usage

Run the application:

```bash
python excelbot_chat.py
```

The application will start and be accessible at `http://localhost:7860`

### Using the Interface

1. **Describe Your Task**: Enter a natural language description of what you want to automate
   - Example: "Format all cells in column A as currency"
   - Example: "Sort data by the date column in ascending order"

2. **Upload Excel File (Optional)**: Upload an Excel file to help the AI understand your file structure

3. **Generate Macro**: Click "Generate VBA Macro" to create your VBA code

4. **Push to GitHub (Optional)**: Enter your repository name and file name, then click "Push to GitHub"

### Example Tasks

- Format cells based on conditions
- Sort and filter data
- Create summary reports
- Data validation and cleaning
- Automated data entry
- Chart generation
- Conditional formatting

## 📁 Project Structure

```
ExcelBotPro/
├── excelbot_chat.py      # Main application file
├── config.py             # Configuration management
├── requirements.txt      # Python dependencies
├── setup.py             # Package setup script
├── README.md            # This file
├── LICENSE              # MIT License
├── .gitignore          # Git ignore rules
├── sample1.xlsx        # Sample Excel file
└── sample2.xlsx        # Sample Excel file
```

## 🔧 Configuration

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `GITHUB_TOKEN` | GitHub Personal Access Token | No |
| `OPENAI_API_KEY` | OpenAI API Key | No |

### Getting a GitHub Token

1. Go to GitHub Settings → Developer settings → Personal access tokens
2. Generate a new token with `repo` scope
3. Copy the token and set it as `GITHUB_TOKEN`

### Getting an OpenAI API Key

1. Sign up at [OpenAI](https://platform.openai.com/)
2. Navigate to API Keys section
3. Create a new API key
4. Set it as `OPENAI_API_KEY`

## 🧪 Testing

Test the application with the included sample Excel files:

```bash
python excelbot_chat.py
```

Then upload `sample1.xlsx` or `sample2.xlsx` to test the file analysis feature.

## 📦 Deployment

### Local Deployment

```bash
python excelbot_chat.py
```

### Production Deployment

For production, consider using:

- **Gunicorn** (for WSGI)
- **Docker** (containerization)
- **Cloud platforms** (Heroku, AWS, Google Cloud, etc.)

Example with Gunicorn:

```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:7860 excelbot_chat:demo
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Mandar Bahadarpurkar**

- GitHub: [@mandarab76](https://github.com/mandarab76)
- Repository: [ExcelBotPro](https://github.com/mandarab76/ExcelBotPro)

## 🙏 Acknowledgments

- [Gradio](https://gradio.app/) for the amazing UI framework
- [PyGithub](https://pygithub.readthedocs.io/) for GitHub integration
- [OpenAI](https://openai.com/) for AI capabilities
- [openpyxl](https://openpyxl.readthedocs.io/) for Excel file handling

## 📞 Support

If you encounter any issues or have questions:

1. Check the [Issues](https://github.com/mandarab76/ExcelBotPro/issues) page
2. Create a new issue with detailed information
3. Include error messages and steps to reproduce

## 🔄 Version History

- **v1.0.0** (2025-01-XX)
  - Initial release
  - VBA macro generation
  - Excel file analysis
  - GitHub integration
  - Modern Gradio UI

## ⚠️ Disclaimer

This tool generates VBA code automatically. Always review and test generated macros before using them in production environments. The authors are not responsible for any data loss or issues arising from the use of generated code.

---

Made with ❤️ for Excel automation enthusiasts
