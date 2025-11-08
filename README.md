
# ExcelBot Pro - Complete Office Add-in Package

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![License](https://img.shields.io/badge/license-MIT-green)

ExcelBot Pro is a comprehensive Excel Office Add-in that provides AI-powered VBA automation, task assistance, and GitHub integration. It includes both a modern Office Add-in with a custom task pane and a collection of powerful VBA macros for Excel automation.

## 📦 Package Contents

### Office Add-in Components
- **manifest.xml** - Office Add-in manifest file
- **src/taskpane/** - Task pane HTML, CSS, and JavaScript files
  - `taskpane.html` - User interface for the add-in
  - `taskpane.css` - Styling for the task pane
  - `taskpane.js` - JavaScript logic and Office.js integration

### VBA Modules
- **src/vba/ExcelBotProCore.bas** - Core automation functions
  - Highlight cells based on values
  - Sort and filter data
  - Create charts
  - Format ranges
  - Generate reports
  - Export to CSV

- **src/vba/ExcelBotProUtils.bas** - Utility functions
  - Calculate statistics
  - Remove duplicates
  - Find and replace
  - Create pivot tables
  - Generate random data
  - Data validation

### Additional Files
- **package.json** - Node.js dependencies and scripts
- **webpack.config.js** - Webpack configuration for bundling
- **assets/** - Icon files for the add-in (placeholders included)
- **excelbot_chat.py** - Optional Gradio chatbot interface
- **sample1.xlsx, sample2.xlsx** - Sample Excel files for testing

## 🚀 Installation

### Option 1: Office Add-in (Recommended)

#### Prerequisites
- Microsoft Excel (Windows, Mac, or Excel Online)
- Node.js (version 14 or higher)
- npm (comes with Node.js)

#### Steps

1. **Clone the repository:**
   ```bash
   git clone https://github.com/mandarab76/ExelBotPro.git
   cd ExelBotPro
   ```

2. **Install Node.js dependencies:**
   ```bash
   npm install
   ```

3. **Generate SSL certificates for local development:**
   ```bash
   npx office-addin-dev-certs install
   ```

4. **Start the development server:**
   ```bash
   npm start
   ```
   This will start the webpack dev server on https://localhost:3000 and automatically sideload the add-in into Excel.

5. **The add-in will appear in Excel:**
   - Look for the "ExcelBot Pro" button in the Home ribbon tab
   - Click "Show Taskpane" to open the add-in

#### Building for Production

To create a production build:
```bash
npm run build
```

The production files will be in the `dist/` directory.

### Option 2: VBA Modules Only

If you only want to use the VBA macros without the Office Add-in:

1. **Open Excel** and press `Alt + F11` to open the VBA Editor

2. **Import the VBA modules:**
   - In the VBA Editor, go to `File > Import File`
   - Navigate to the `src/vba/` directory
   - Import `ExcelBotProCore.bas`
   - Import `ExcelBotProUtils.bas`

3. **Enable macros:**
   - Go to `File > Options > Trust Center > Trust Center Settings > Macro Settings`
   - Select "Enable all macros" (for development) or "Disable all macros with notification"

4. **Run macros:**
   - Press `Alt + F8` to open the Macro dialog
   - Select a macro from the list and click "Run"

### Option 3: Python Chatbot Interface

For the optional Gradio chatbot interface:

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Set your GitHub token (optional, for GitHub integration):**
   ```bash
   export GITHUB_TOKEN=your_personal_access_token
   ```

3. **Run the chatbot:**
   ```bash
   python excelbot_chat.py
   ```

4. **Open the web interface:**
   - The chatbot will start on http://localhost:7860
   - Describe your Excel task and generate VBA macros

## 🎯 Features

### Office Add-in Task Pane
- **Generate VBA Macros**: Describe your task in natural language and get VBA code
- **Analyze Workbook**: Get statistics and information about your current workbook
- **GitHub Integration**: Push generated macros to GitHub repositories
- **Quick Actions**: 
  - Format selected range
  - Create charts
  - Sort data
  - Apply filters

### VBA Automation Functions

#### Core Functions (ExcelBotProCore.bas)
- `HighlightCells()` - Color cells based on value thresholds
- `SortData()` - Sort data by any column
- `FilterData()` - Apply AutoFilter to data
- `CreateChart()` - Generate various chart types
- `FormatRange()` - Apply professional formatting
- `GenerateSummaryReport()` - Create summary reports
- `ExportToCSV()` - Export data to CSV files

#### Utility Functions (ExcelBotProUtils.bas)
- `ShowRangeStatistics()` - Calculate and display statistics
- `RemoveDuplicates()` - Remove duplicate rows
- `FindAndReplaceText()` - Find and replace operations
- `CreatePivotTable()` - Generate pivot tables
- `GenerateRandomData()` - Create test data
- `ColorAlternateRows()` - Apply zebra striping
- `AddValidationList()` - Add data validation dropdowns
- And many more...

## 📚 Usage Examples

### Using the Task Pane Add-in

1. **Generate a macro to highlight cells:**
   - Open the ExcelBot Pro task pane
   - In "Describe your Excel task", type: "Highlight cells with values above 100"
   - Click "Generate Macro"
   - Copy the code or get instructions to insert it into VBA

2. **Analyze your workbook:**
   - Click "Analyze Workbook"
   - View sheet names, row counts, and column counts

3. **Quick formatting:**
   - Select a range in Excel
   - Click "Format Selection" in the Quick Actions section

### Running VBA Macros

1. **From the Macro Dialog:**
   - Press `Alt + F8`
   - Select macro (e.g., `HighlightCells`)
   - Click "Run"

2. **From the VBA Editor:**
   - Press `Alt + F11`
   - Press `F5` while cursor is in a macro

3. **Create a custom ribbon button:**
   - Right-click Excel ribbon > Customize the Ribbon
   - Add a new group and assign macros to buttons

## 🔧 Configuration

### Customize Manifest

Edit `manifest.xml` to change:
- Add-in name and description
- Icon URLs
- Default source location
- Permissions

### Customize VBA Macros

Edit the `.bas` files in `src/vba/` to:
- Modify color schemes
- Change default thresholds
- Add new automation functions
- Customize report formats

### Webpack Configuration

Edit `webpack.config.js` to:
- Change development server port
- Add new entry points
- Configure build optimization
- Add plugins

## 🛠️ Development

### Project Structure
```
ExelBotPro/
├── manifest.xml                 # Office Add-in manifest
├── package.json                 # Node.js dependencies
├── webpack.config.js            # Webpack configuration
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html       # Task pane UI
│   │   ├── taskpane.css        # Styles
│   │   └── taskpane.js         # JavaScript logic
│   └── vba/
│       ├── ExcelBotProCore.bas # Core VBA functions
│       └── ExcelBotProUtils.bas # Utility VBA functions
├── assets/                      # Icon files
├── excelbot_chat.py            # Python chatbot
└── sample1.xlsx, sample2.xlsx  # Sample files
```

### Available npm Scripts

- `npm start` - Start development server and sideload add-in
- `npm run build` - Create production build
- `npm run validate` - Validate manifest.xml
- `npm run serve` - Start webpack dev server only

### Testing

1. **Test the add-in in Excel:**
   - Use `npm start` to sideload automatically
   - Or manually sideload via Excel: Insert > Get Add-ins > Upload Manifest

2. **Test VBA macros:**
   - Create test data in Excel
   - Run macros and verify results
   - Use `Debug.Print` statements in VBA

3. **Validate manifest:**
   ```bash
   npm run validate
   ```

## 📝 Documentation

- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Office.js API Reference](https://learn.microsoft.com/en-us/javascript/api/office)
- [VBA Language Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel)

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🔗 Links

- **Repository**: https://github.com/mandarab76/ExelBotPro
- **Issues**: https://github.com/mandarab76/ExelBotPro/issues

## 🙏 Acknowledgments

- Built with [Office.js](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
- UI inspired by Microsoft Office design guidelines
- VBA automation patterns from Excel community

## 📞 Support

For support, please open an issue on GitHub or contact the maintainers.

---

**Version 1.0.0** | Made with ❤️ for Excel automation enthusiasts
