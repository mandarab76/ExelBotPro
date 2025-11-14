# ExcelBot Pro - Complete User Manual 📘

**Welcome to ExcelBot Pro!** This guide will help you install and use the application, even if you've never worked with Python or GitHub before.

---

## 📂 Project Structure

```
ExcelBot-Pro/
│
├── excelbot_chat.py          # Main application file (the chatbot)
├── requirements.txt           # List of required software packages
├── LICENSE                    # Legal license information
├── README.md                  # Quick project overview
├── USER_MANUAL.md            # This detailed guide
├── sample1.xlsx              # Sample Excel file for testing
└── sample2.xlsx              # Another sample Excel file
```

---

## 🎯 What Does This Application Do?

ExcelBot Pro is a **chatbot** that helps you create VBA macros (automation scripts) for Microsoft Excel. Instead of learning complex VBA programming, you simply tell the chatbot what you want in plain English, and it generates the code for you!

### Key Features:
1. **Generate VBA Macros** - Create Excel automation scripts from simple descriptions
2. **Analyze Excel Files** - Upload Excel files to see their structure and data
3. **GitHub Integration** - Save your generated code to GitHub (optional)

---

## 💻 System Requirements

### What You Need:
- **Operating System**: Windows, Mac, or Linux
- **Python**: Version 3.7 or higher (we'll help you install this)
- **Internet Connection**: Required for downloading software and (optionally) GitHub features
- **Web Browser**: Chrome, Firefox, Safari, or Edge

### Disk Space:
- Approximately 500 MB for Python and all required packages

---

## 🚀 Step-by-Step Installation Guide

### Step 1: Install Python

Python is the programming language that runs this application.

#### For Windows:
1. Go to https://www.python.org/downloads/
2. Click the yellow "Download Python" button (get version 3.11 or higher)
3. **IMPORTANT**: Run the installer and check the box that says "Add Python to PATH"
4. Click "Install Now"
5. Wait for installation to complete
6. Open Command Prompt (search for "cmd" in Start menu)
7. Type `python --version` and press Enter to verify installation

#### For Mac:
1. Go to https://www.python.org/downloads/
2. Download the macOS installer
3. Open the downloaded file and follow installation steps
4. Open Terminal (find it in Applications → Utilities)
5. Type `python3 --version` and press Enter to verify installation

#### For Linux (Ubuntu/Debian):
1. Open Terminal
2. Type: `sudo apt update`
3. Type: `sudo apt install python3 python3-pip`
4. Type: `python3 --version` to verify installation

---

### Step 2: Extract the ExcelBot Pro Files

1. Locate the `ExcelBot-Pro.zip` file you downloaded
2. **Windows**: Right-click → "Extract All" → Choose a location (like your Desktop or Documents folder)
3. **Mac**: Double-click the zip file (it extracts automatically)
4. **Linux**: Right-click → "Extract Here" or use command: `unzip ExcelBot-Pro.zip`

---

### Step 3: Install Required Packages

These are additional software components the application needs to run.

#### For Windows:
1. Open Command Prompt
2. Navigate to the ExcelBot Pro folder:
   ```
   cd Desktop\ExcelBot-Pro
   ```
   (Adjust the path based on where you extracted the files)
3. Install packages:
   ```
   pip install -r requirements.txt
   ```
4. Wait 2-5 minutes for installation to complete

#### For Mac/Linux:
1. Open Terminal
2. Navigate to the ExcelBot Pro folder:
   ```
   cd ~/Desktop/ExcelBot-Pro
   ```
   (Adjust the path based on where you extracted the files)
3. Install packages:
   ```
   pip3 install -r requirements.txt
   ```
4. Wait 2-5 minutes for installation to complete

---

## 🔑 Setting Up GitHub Integration (OPTIONAL)

If you want to save your generated VBA code to GitHub, follow these steps. **Skip this section if you don't need GitHub features.**

### What is GitHub?
GitHub is a website where programmers store and share their code. It's like Google Drive or Dropbox, but specifically designed for code.

### Step 1: Create a GitHub Account
1. Go to https://github.com
2. Click "Sign up" in the top-right corner
3. Enter your email, create a password, and choose a username
4. Verify your email address
5. Complete the setup wizard

### Step 2: Create a Personal Access Token

A token is like a special password that lets ExcelBot Pro access your GitHub account.

1. Log in to GitHub at https://github.com
2. Click your profile picture (top-right) → **Settings**
3. Scroll down on the left sidebar and click **Developer settings**
4. Click **Personal access tokens** → **Tokens (classic)**
5. Click **Generate new token** → **Generate new token (classic)**
6. You might need to enter your password again
7. Fill in the form:
   - **Note**: Type "ExcelBot Pro" (so you remember what this is for)
   - **Expiration**: Choose "No expiration" or select a time period
   - **Select scopes**: Check the box next to **"repo"** (this allows full repository access)
8. Scroll to the bottom and click **Generate token**
9. **IMPORTANT**: Copy the token (it looks like: `ghp_1234abcd...`)
   - Save it in a safe place (like a password manager or text file)
   - You won't be able to see it again!

### Step 3: Set Up the Token on Your Computer

#### For Windows:
1. Open Command Prompt
2. Type this command (replace `YOUR_TOKEN_HERE` with your actual token):
   ```
   setx GITHUB_TOKEN "ghp_your_actual_token_here"
   ```
3. Close and reopen Command Prompt for changes to take effect

#### For Mac/Linux:
1. Open Terminal
2. Type this command:
   ```
   nano ~/.bash_profile
   ```
   (On some systems, use `nano ~/.bashrc` or `nano ~/.zshrc`)
3. Add this line at the end (replace with your actual token):
   ```
   export GITHUB_TOKEN="ghp_your_actual_token_here"
   ```
4. Press `Ctrl+X`, then `Y`, then `Enter` to save
5. Type: `source ~/.bash_profile` to apply changes

### Step 4: Create a GitHub Repository (Where Your Code Will Be Saved)

1. Log in to GitHub
2. Click the **+** icon (top-right) → **New repository**
3. Repository name: Type something like "my-vba-macros"
4. Description: Optional (e.g., "VBA macros from ExcelBot Pro")
5. Choose **Public** or **Private**
6. **Do NOT** check "Initialize with README"
7. Click **Create repository**
8. Remember your repository name (you'll need it as: `your-username/my-vba-macros`)

---

## 🎮 How to Use ExcelBot Pro

### Starting the Application

#### For Windows:
1. Open Command Prompt
2. Navigate to the ExcelBot Pro folder:
   ```
   cd Desktop\ExcelBot-Pro
   ```
3. Start the application:
   ```
   python excelbot_chat.py
   ```

#### For Mac/Linux:
1. Open Terminal
2. Navigate to the ExcelBot Pro folder:
   ```
   cd ~/Desktop/ExcelBot-Pro
   ```
3. Start the application:
   ```
   python3 excelbot_chat.py
   ```

### What You'll See

After running the command, you'll see text like:
```
Running on local URL:  http://127.0.0.1:7860
```

**This means it's working!** The application is now running on your computer.

### Opening the Application in Your Browser

1. Open your web browser (Chrome, Firefox, Safari, or Edge)
2. Type this address in the URL bar:
   ```
   http://127.0.0.1:7860
   ```
3. Press Enter
4. The ExcelBot Pro interface will load!

---

## 📖 Using the Features

### Feature 1: Generate VBA Macros

VBA macros are small programs that automate tasks in Excel.

**Step-by-Step:**

1. In the text box labeled **"Describe your Excel task"**, type what you want to do. Examples:
   - "format my data"
   - "sort by first column"
   - "calculate totals"
   - "create a chart"
   - "remove duplicates"
   - "apply filter"

2. Click the **"Generate VBA Macro"** button

3. The VBA code will appear in the code box below

4. **To use this code in Excel:**
   - Open Microsoft Excel
   - Press `Alt+F11` (Windows) or `Fn+Option+F11` (Mac) to open the VBA Editor
   - Click **Insert** → **Module**
   - Copy the generated code from ExcelBot Pro
   - Paste it into the module window
   - Close the VBA Editor
   - Press `Alt+F8` to see your macros
   - Select the macro and click "Run"

### Feature 2: Upload and Analyze Excel Files

This feature lets you see information about your Excel file.

**Step-by-Step:**

1. Under **"Excel File Upload"**, click the upload area
2. Browse your computer and select an Excel file (.xls or .xlsx)
   - You can use `sample1.xlsx` or `sample2.xlsx` included in the package
3. The application will show you:
   - Filename
   - Number of rows
   - Number of columns
   - Names of all columns

**Why is this useful?**
- Quickly check file contents without opening Excel
- Verify data structure before creating macros
- Confirm column names for more specific VBA generation

### Feature 3: Push Code to GitHub (Optional)

If you set up GitHub, you can save your generated VBA code online.

**Step-by-Step:**

1. First, generate a VBA macro (see Feature 1)
2. In the **"Repository Name"** field, enter your repository in this format:
   ```
   your-username/my-vba-macros
   ```
   Example: `johnsmith/my-vba-macros`

3. In the **"File Name"** field, enter a name for your file:
   ```
   format_macro.vba
   ```
   (The `.vba` extension will be added automatically if you forget it)

4. Click **"Push to GitHub"**

5. You'll see a success message or an error if something went wrong

6. **To view your saved code:**
   - Go to https://github.com
   - Click on your repository
   - You'll see your VBA file listed
   - Click on it to view the code

---

## 🎓 Practical Examples

### Example 1: Format Sales Data

**Scenario**: You have a messy Excel spreadsheet with sales data and want to make it look professional.

1. Start ExcelBot Pro
2. Type: "format my data with borders and auto-fit columns"
3. Click "Generate VBA Macro"
4. Copy the generated code
5. Open your Excel file
6. Press `Alt+F11` to open VBA Editor
7. Insert → Module
8. Paste the code
9. Close VBA Editor
10. Press `Alt+F8`, select "FormatData", and click Run
11. Your data is now beautifully formatted!

### Example 2: Sort Employee List

**Scenario**: You have a list of employees and want to sort them alphabetically.

1. Type: "sort by first column"
2. Click "Generate VBA Macro"
3. Follow steps 4-11 from Example 1
4. Your data is now sorted!

### Example 3: Create a Chart

**Scenario**: You want to visualize your data with a chart.

1. Type: "create a chart from my data"
2. Click "Generate VBA Macro"
3. Follow the VBA implementation steps
4. A column chart will be created automatically!

---

## 🔧 Dependencies Explained

When you installed the application, these packages were installed:

### 1. **Gradio**
- **What it does**: Creates the web interface (the page you see in your browser)
- **Website**: https://www.gradio.app
- **Why we need it**: Makes the application easy to use with buttons and text boxes

### 2. **PyGithub**
- **What it does**: Connects Python to GitHub
- **Website**: https://pygithub.readthedocs.io
- **Why we need it**: Allows pushing your VBA code to GitHub repositories

### 3. **Pandas**
- **What it does**: Reads and analyzes Excel files
- **Website**: https://pandas.pydata.org
- **Why we need it**: Extracts information from your uploaded Excel files

### 4. **Openpyxl**
- **What it does**: Works with Excel file formats (.xlsx)
- **Website**: https://openpyxl.readthedocs.io
- **Why we need it**: Helps pandas read modern Excel files

---

## ❓ Troubleshooting Common Issues

### Issue 1: "Python is not recognized"

**Problem**: When you type `python` in Command Prompt/Terminal, you get an error.

**Solution**:
- **Windows**: Reinstall Python and make sure to check "Add Python to PATH"
- **Mac/Linux**: Use `python3` instead of `python`

---

### Issue 2: "Module not found" Error

**Problem**: Error message says a module (like gradio) is not found.

**Solution**:
```
pip install -r requirements.txt
```
or
```
pip3 install -r requirements.txt
```

---

### Issue 3: Application Won't Start

**Problem**: Nothing happens when you run the command.

**Solution**:
1. Make sure you're in the correct folder (use `cd` command)
2. Check if Python is installed: `python --version`
3. Try running with `python3` instead of `python`

---

### Issue 4: GitHub Push Fails

**Problem**: Error when trying to push to GitHub.

**Possible causes and solutions**:
1. **Token not set**: Go back to "Setting Up GitHub Integration" section
2. **Wrong repository name**: Make sure it's formatted as `username/repo-name`
3. **Repository doesn't exist**: Create the repository on GitHub first
4. **Token expired**: Generate a new token on GitHub

---

### Issue 5: Can't Open in Browser

**Problem**: The URL doesn't load.

**Solution**:
1. Make sure the application is still running in Command Prompt/Terminal
2. Try this URL instead: `http://localhost:7860`
3. Try a different browser
4. Check if another application is using port 7860

---

## 🛡️ Security Best Practices

### Protecting Your GitHub Token

1. **Never share your token** with anyone
2. **Don't post it online** (in forums, social media, etc.)
3. **Store it securely** (use a password manager)
4. **Revoke old tokens** if you think they're compromised:
   - Go to GitHub → Settings → Developer settings → Personal access tokens
   - Click "Delete" on any suspicious tokens

### Using the Application Safely

1. **Only upload your own Excel files** - don't upload sensitive company data without permission
2. **Review generated VBA code** before running it in Excel
3. **Back up important Excel files** before testing macros
4. **Use private GitHub repositories** for sensitive code

---

## 🎉 Quick Start Checklist

- [ ] Python installed and working
- [ ] ExcelBot Pro files extracted
- [ ] Dependencies installed (`pip install -r requirements.txt`)
- [ ] (Optional) GitHub account created
- [ ] (Optional) GitHub token generated and configured
- [ ] (Optional) GitHub repository created
- [ ] Application started successfully
- [ ] Browser opened to http://127.0.0.1:7860
- [ ] First VBA macro generated!

---

## 📞 Need More Help?

### Learning Resources

1. **Python Basics**: https://www.python.org/about/gettingstarted/
2. **Excel VBA Tutorial**: https://www.excel-easy.com/vba.html
3. **GitHub Guides**: https://guides.github.com
4. **Gradio Documentation**: https://www.gradio.app/docs/

### Testing the Sample Files

The package includes two sample Excel files you can use to test the upload feature:
- `sample1.xlsx` - Basic data file
- `sample2.xlsx` - Another sample for testing

Simply upload these files to see how the analysis feature works!

---

## 📋 Summary

**ExcelBot Pro** makes Excel automation accessible to everyone. You don't need to be a programmer to create powerful VBA macros. Just describe what you want in plain English, and the chatbot does the rest!

**Remember**:
1. Keep your GitHub token safe
2. Back up important files before testing macros
3. The application runs locally on your computer (your data stays private)
4. GitHub features are completely optional

**Enjoy automating your Excel tasks!** 🎊

---

*Version 1.0 - Created for ExcelBot Pro*
*For technical support, refer to the troubleshooting section above.*
