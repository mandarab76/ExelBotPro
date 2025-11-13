# Contributing to ExcelBot Pro

First off, thank you for considering contributing to ExcelBot Pro! It's people like you that make ExcelBot Pro such a great tool.

## 📋 Table of Contents
- [Code of Conduct](#code-of-conduct)
- [How Can I Contribute?](#how-can-i-contribute)
- [Development Setup](#development-setup)
- [Pull Request Process](#pull-request-process)
- [Coding Standards](#coding-standards)
- [Community](#community)

## 📜 Code of Conduct

This project and everyone participating in it is governed by our commitment to providing a welcoming and inclusive environment. By participating, you are expected to uphold this code.

### Our Standards

- ✅ Using welcoming and inclusive language
- ✅ Being respectful of differing viewpoints and experiences
- ✅ Gracefully accepting constructive criticism
- ✅ Focusing on what is best for the community
- ❌ Using sexualized language or imagery
- ❌ Trolling, insulting/derogatory comments, and personal attacks
- ❌ Public or private harassment

## 🤝 How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check existing issues to avoid duplicates. When you create a bug report, include as many details as possible:

**Bug Report Template:**
```markdown
## Description
A clear and concise description of the bug.

## Steps to Reproduce
1. Go to '...'
2. Click on '...'
3. Scroll down to '...'
4. See error

## Expected Behavior
What you expected to happen.

## Actual Behavior
What actually happened.

## Screenshots
If applicable, add screenshots.

## Environment
- OS: [e.g., Windows 10, Ubuntu 20.04]
- Python Version: [e.g., 3.9.7]
- ExcelBot Pro Version: [e.g., 1.0.0]

## Additional Context
Any other context about the problem.
```

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion, include:

- **Clear title and description**
- **Step-by-step description** of the suggested enhancement
- **Explain why** this enhancement would be useful
- **List similar features** in other tools (if applicable)

### Your First Code Contribution

Unsure where to begin? Look for issues labeled:
- `good first issue` - Simple issues for beginners
- `help wanted` - Issues where we need help
- `documentation` - Documentation improvements

### Pull Requests

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 🛠️ Development Setup

### Prerequisites
- Python 3.7+
- Git
- Text editor or IDE (VS Code recommended)

### Setup Instructions

1. **Fork and clone:**
```bash
git clone https://github.com/YOUR_USERNAME/ExcelBotPro.git
cd ExcelBotPro
```

2. **Create virtual environment:**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install dependencies:**
```bash
pip install -r requirements.txt
```

4. **Configure environment:**
```bash
cp .env.example .env
# Edit .env with your settings
```

5. **Run the application:**
```bash
python excelbot_chat.py
```

### Running Tests

```bash
# Install test dependencies
pip install pytest pytest-cov

# Run tests
pytest

# Run with coverage
pytest --cov=excelbot_chat
```

## 🔄 Pull Request Process

### Before Submitting

1. ✅ **Update documentation** if you changed functionality
2. ✅ **Add tests** for new features
3. ✅ **Ensure all tests pass**
4. ✅ **Update README.md** if needed
5. ✅ **Follow coding standards** (see below)
6. ✅ **Update CHANGELOG.md** with your changes

### PR Description Template

```markdown
## Description
Brief description of changes.

## Type of Change
- [ ] Bug fix (non-breaking change that fixes an issue)
- [ ] New feature (non-breaking change that adds functionality)
- [ ] Breaking change (fix or feature that would cause existing functionality to not work as expected)
- [ ] Documentation update

## How Has This Been Tested?
Describe the tests you ran.

## Checklist
- [ ] My code follows the style guidelines
- [ ] I have performed a self-review
- [ ] I have commented my code where needed
- [ ] I have made corresponding changes to documentation
- [ ] My changes generate no new warnings
- [ ] I have added tests that prove my fix/feature works
- [ ] New and existing unit tests pass locally

## Screenshots (if applicable)
Add screenshots to demonstrate the changes.
```

### Review Process

1. Maintainers will review your PR within 1-2 weeks
2. Address any requested changes
3. Once approved, a maintainer will merge your PR
4. Your contribution will be credited in the CHANGELOG

## 📐 Coding Standards

### Python Style Guide

Follow [PEP 8](https://www.python.org/dev/peps/pep-0008/) with these specifics:

```python
# Imports
import os
import sys
from typing import List, Dict

# Constants
MAX_FILE_SIZE = 10_000_000
DEFAULT_PORT = 7860

# Functions - use docstrings
def process_excel_file(file_path: str) -> Dict:
    """
    Process an Excel file and return analytics.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Dictionary containing file analytics
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file is invalid
    """
    pass

# Classes - use docstrings
class ExcelProcessor:
    """Handles Excel file processing and analysis."""
    
    def __init__(self, file_path: str):
        """Initialize processor with file path."""
        self.file_path = file_path
```

### VBA Style Guide

```vba
' Use descriptive variable names
Dim worksheetName As String
Dim lastRow As Long

' Add comments for complex logic
' Calculate the last row with data
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Error handling
On Error GoTo ErrorHandler

' ... code ...

Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
```

### Commit Messages

Follow the [Conventional Commits](https://www.conventionalcommits.org/) specification:

```
feat: add new pivot table template
fix: resolve GitHub authentication issue
docs: update installation instructions
style: format code with black
refactor: simplify VBA template matching
test: add tests for Excel analyzer
chore: update dependencies
```

### Code Organization

```
ExcelBotPro/
├── excelbot_chat.py       # Main application
├── requirements.txt       # Dependencies
├── README.md             # Project documentation
├── CONTRIBUTING.md       # This file
├── LICENSE              # MIT License
├── .env.example         # Environment template
├── .gitignore          # Git ignore rules
├── Dockerfile          # Docker configuration
├── docker-compose.yml  # Docker Compose config
├── setup.sh           # Linux/Mac setup script
├── setup.bat          # Windows setup script
└── tests/             # Test files
    ├── __init__.py
    ├── test_vba_generator.py
    └── test_excel_analyzer.py
```

## 🎨 Adding New VBA Templates

To add a new VBA template:

1. **Add to VBA_TEMPLATES dictionary:**
```python
VBA_TEMPLATES["your_template"] = """Sub YourMacro()
    ' Your VBA code here
End Sub"""
```

2. **Add template matching logic:**
```python
elif "your_keyword" in task_lower:
    return VBA_TEMPLATES["your_template"]
```

3. **Update documentation:**
   - Add to README.md template list
   - Add to in-app help
   - Add example usage

4. **Add tests:**
```python
def test_your_template():
    result = generate_vba_macro("your keyword")
    assert "YourMacro" in result
```

## 🐛 Debugging

### Enable Debug Mode

```python
# In excelbot_chat.py
demo.launch(
    debug=True,
    show_error=True
)
```

### Check Logs

```bash
# Redirect output to log file
python excelbot_chat.py > debug.log 2>&1
```

## 🌍 Community

### Communication Channels

- **GitHub Issues**: Bug reports and feature requests
- **GitHub Discussions**: Questions and general discussion
- **Pull Requests**: Code contributions

### Getting Help

If you need help:
1. Check existing [Issues](https://github.com/mandarab76/ExcelBotPro/issues)
2. Check [Discussions](https://github.com/mandarab76/ExcelBotPro/discussions)
3. Create a new issue with detailed information

## 🏆 Recognition

Contributors will be recognized in:
- README.md (Contributors section)
- CHANGELOG.md (for each release)
- GitHub Contributors page

## 📝 License

By contributing, you agree that your contributions will be licensed under the MIT License.

## 🙏 Thank You!

Your contributions make ExcelBot Pro better for everyone. We appreciate your time and effort!

---

**Questions?** Feel free to reach out by opening an issue or starting a discussion.

**Happy Contributing! 🎉**
