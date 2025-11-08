# ExcelBot Pro Installation Guide

This guide provides detailed step-by-step instructions for installing and configuring ExcelBot Pro on various platforms.

## Table of Contents

1. [System Requirements](#system-requirements)
2. [Windows Installation](#windows-installation)
3. [Mac Installation](#mac-installation)
4. [Excel Online Installation](#excel-online-installation)
5. [Troubleshooting](#troubleshooting)

## System Requirements

### For Office Add-in
- **Operating System**: Windows 10+, macOS 10.14+, or web browser for Excel Online
- **Excel Version**: 
  - Windows: Excel 2016 or later
  - Mac: Excel 2016 or later
  - Web: Excel Online (modern browser required)
- **Node.js**: Version 14.0 or higher
- **npm**: Version 6.0 or higher (included with Node.js)

### For VBA Modules Only
- **Excel Version**: Excel 2010 or later (Windows or Mac)
- **Macro Support**: Macros must be enabled

### For Python Chatbot
- **Python**: Version 3.7 or higher
- **pip**: Latest version

## Windows Installation

### Step 1: Install Node.js

1. Download Node.js from https://nodejs.org/
2. Run the installer and follow the prompts
3. Verify installation:
   ```cmd
   node --version
   npm --version
   ```

### Step 2: Clone the Repository

```cmd
git clone https://github.com/mandarab76/ExelBotPro.git
cd ExelBotPro
```

### Step 3: Install Dependencies

```cmd
npm install
```

### Step 4: Generate SSL Certificates

```cmd
npx office-addin-dev-certs install
```

When prompted, click "Yes" to trust the certificate.

### Step 5: Start the Add-in

```cmd
npm start
```

This will:
- Start the webpack dev server on https://localhost:3000
- Open Excel
- Automatically sideload the add-in

### Step 6: Use the Add-in

1. Look for the **"ExcelBot Pro"** button in the Home ribbon tab
2. Click **"Show Taskpane"** to open the add-in panel
3. Start using the features!

## Mac Installation

### Step 1: Install Node.js

1. Install Homebrew (if not already installed):
   ```bash
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
   ```

2. Install Node.js:
   ```bash
   brew install node
   ```

3. Verify installation:
   ```bash
   node --version
   npm --version
   ```

### Step 2: Clone the Repository

```bash
git clone https://github.com/mandarab76/ExelBotPro.git
cd ExelBotPro
```

### Step 3: Install Dependencies

```bash
npm install
```

### Step 4: Generate SSL Certificates

```bash
npx office-addin-dev-certs install
```

Follow the prompts to trust the certificate.

### Step 5: Start the Add-in

```bash
npm start
```

### Step 6: Use the Add-in

1. Excel should open automatically
2. Look for **"ExcelBot Pro"** in the Home ribbon
3. Click **"Show Taskpane"**

## Excel Online Installation

### Step 1: Setup Local Development

Follow steps 1-4 from either Windows or Mac installation to set up the local server.

### Step 2: Start the Server

```bash
npm run serve
```

This starts the server without trying to open Excel.

### Step 3: Upload Manifest to Excel Online

1. Open Excel Online at https://office.com
2. Create or open a workbook
3. Click **Insert** > **Add-ins** > **Upload My Add-in**
4. Click **Browse** and select `manifest.xml` from the repository
5. Click **Upload**

### Step 4: Use the Add-in

The add-in will now be available in the current workbook.

## Manual Sideloading (Alternative Method)

If automatic sideloading doesn't work, you can manually sideload the add-in:

### Windows

1. Start the dev server:
   ```cmd
   npm run serve
   ```

2. Open Excel
3. Go to **Insert** > **Get Add-ins**
4. Click **Upload My Add-in** (bottom right)
5. Browse to `manifest.xml` and click **Upload**

### Mac

1. Start the dev server:
   ```bash
   npm run serve
   ```

2. Open Excel
3. Go to **Insert** > **Add-ins** > **My Add-ins**
4. Click **+ Add a Custom Add-in**
5. Choose **manifest.xml** from the repository

## VBA Modules Installation (Without Office Add-in)

If you only want to use the VBA macros:

### Step 1: Enable Developer Tab

#### Windows:
1. File > Options > Customize Ribbon
2. Check "Developer" in the right column
3. Click OK

#### Mac:
1. Excel > Preferences > Ribbon & Toolbar
2. Check "Developer Tab"
3. Click Save

### Step 2: Enable Macros

#### Windows:
1. File > Options > Trust Center > Trust Center Settings
2. Click "Macro Settings"
3. Select "Enable all macros" (for development)
4. Click OK

#### Mac:
1. Excel > Preferences > Security & Privacy
2. Select "Enable all macros"

### Step 3: Import VBA Modules

1. Press `Alt + F11` (Windows) or `Fn + Option + F11` (Mac) to open VBA Editor
2. In VBA Editor, go to **File** > **Import File**
3. Navigate to `src/vba/` directory in the repository
4. Import `ExcelBotProCore.bas`
5. Import `ExcelBotProUtils.bas`

### Step 4: Run Macros

1. Press `Alt + F8` (Windows) or `Fn + Option + F8` (Mac)
2. Select a macro from the list
3. Click "Run"

## Python Chatbot Installation

### Step 1: Install Python

Download and install Python from https://www.python.org/downloads/

### Step 2: Install Dependencies

```bash
cd ExelBotPro
pip install -r requirements.txt
```

### Step 3: Set GitHub Token (Optional)

#### Windows (Command Prompt):
```cmd
set GITHUB_TOKEN=your_github_personal_access_token
```

#### Windows (PowerShell):
```powershell
$env:GITHUB_TOKEN="your_github_personal_access_token"
```

#### Mac/Linux:
```bash
export GITHUB_TOKEN=your_github_personal_access_token
```

### Step 4: Run the Chatbot

```bash
python excelbot_chat.py
```

### Step 5: Access the Interface

Open your browser and go to http://localhost:7860

## Troubleshooting

### Issue: SSL Certificate Errors

**Solution:**
```bash
npx office-addin-dev-certs install --force
```

### Issue: Port 3000 Already in Use

**Solution:**
Edit `webpack.config.js` and change the port number:
```javascript
devServer: {
  port: 3001  // Change to any available port
}
```

Also update the URLs in `manifest.xml` to use the new port.

### Issue: Add-in Doesn't Appear in Excel

**Solutions:**
1. Clear Office cache:
   - Windows: Delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Mac: Delete `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
2. Restart Excel
3. Try manual sideloading

### Issue: Macros Don't Run

**Solutions:**
1. Check if macros are enabled in Excel settings
2. Verify that the VBA modules are properly imported
3. Check for syntax errors in VBA Editor (press F8 to step through code)

### Issue: npm install Fails

**Solutions:**
1. Clear npm cache:
   ```bash
   npm cache clean --force
   ```
2. Delete `node_modules` folder and `package-lock.json`
3. Run `npm install` again

### Issue: Python Dependencies Installation Fails

**Solutions:**
1. Upgrade pip:
   ```bash
   pip install --upgrade pip
   ```
2. Install dependencies one by one to identify the problematic package
3. Check Python version compatibility

### Issue: Webpack Build Errors

**Solution:**
```bash
rm -rf node_modules package-lock.json
npm install
npm run build
```

## Getting Help

If you encounter issues not covered in this guide:

1. Check the [GitHub Issues](https://github.com/mandarab76/ExelBotPro/issues) page
2. Create a new issue with:
   - Your operating system and version
   - Excel version
   - Error messages or screenshots
   - Steps to reproduce the problem

## Next Steps

After installation:

1. Read the main [README.md](README.md) for feature overview
2. Try the sample Excel files included in the repository
3. Experiment with generating VBA macros
4. Explore the quick actions in the task pane
5. Customize the VBA modules for your needs

---

**Need more help?** Visit our [documentation](https://github.com/mandarab76/ExelBotPro) or open an issue on GitHub.
