"""
ExcelBot Pro - Advanced VBA Automation Chatbot
A professional tool for generating VBA macros, processing Excel files, and managing code via GitHub
Author: Mandar Bahadarpurkar
License: MIT
"""

import os
import re
import gradio as gr
from github import Github
import pandas as pd
from datetime import datetime
import traceback

# VBA Macro Templates
VBA_TEMPLATES = {
    "sort": """Sub SortData()
    ' Automatically sorts data in the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Assuming data starts from A1
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A1"), Order:=xlAscending
    
    With ws.Sort
        .SetRange Range("A1:Z" & lastRow)
        .Header = xlYes
        .Apply
    End With
    
    MsgBox "Data sorted successfully!", vbInformation
End Sub""",
    
    "filter": """Sub FilterData()
    ' Applies AutoFilter to the data
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    Else
        Dim lastRow As Long
        Dim lastCol As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
        MsgBox "AutoFilter applied successfully!", vbInformation
    End If
End Sub""",
    
    "highlight": """Sub HighlightDuplicates()
    ' Highlights duplicate values in the selected range
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Please select a range first!", vbExclamation
        Exit Sub
    End If
    
    For Each cell In rng
        If cell.Value <> "" Then
            If dict.exists(cell.Value) Then
                cell.Interior.Color = RGB(255, 200, 200)
            Else
                dict.Add cell.Value, 1
            End If
        End If
    Next cell
    
    MsgBox "Duplicates highlighted!", vbInformation
End Sub""",
    
    "remove_duplicates": """Sub RemoveDuplicates()
    ' Removes duplicate rows from the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    dataRange.RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    MsgBox "Duplicates removed successfully!", vbInformation
End Sub""",
    
    "format": """Sub FormatTable()
    ' Applies professional formatting to the data
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Format headers
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Add borders
    tableRange.Borders.LineStyle = xlContinuous
    tableRange.Borders.Weight = xlThin
    
    ' Auto-fit columns
    ws.Columns.AutoFit
    
    MsgBox "Table formatted successfully!", vbInformation
End Sub""",
    
    "sum": """Sub CalculateSums()
    ' Adds sum formulas to numeric columns
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If IsNumeric(ws.Cells(2, col).Value) Then
            ws.Cells(lastRow + 1, col).Formula = "=SUM(" & ws.Cells(2, col).Address & ":" & ws.Cells(lastRow, col).Address & ")"
            ws.Cells(lastRow + 1, col).Font.Bold = True
        End If
    Next col
    
    MsgBox "Sum formulas added!", vbInformation
End Sub""",
    
    "pivot": """Sub CreatePivotTable()
    ' Creates a pivot table from the active data
    Dim ws As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim srcData As Range
    
    Set ws = ActiveSheet
    
    ' Define source data
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set srcData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Create new sheet for pivot table
    Dim newWs As Worksheet
    Set newWs = Worksheets.Add
    newWs.Name = "PivotTable_" & Format(Now, "hhmmss")
    
    ' Create pivot table
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcData)
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=newWs.Range("A1"), TableName:="PivotTable1")
    
    MsgBox "Pivot table created in new sheet!", vbInformation
End Sub"""
}

def generate_vba_macro(task_description):
    """
    Generate VBA macro based on task description with intelligent template matching
    """
    if not task_description or task_description.strip() == "":
        return "Please provide a task description."
    
    task_lower = task_description.lower()
    
    # Smart template matching
    if any(word in task_lower for word in ["sort", "order", "arrange"]):
        return VBA_TEMPLATES["sort"]
    elif any(word in task_lower for word in ["filter", "search", "find"]):
        return VBA_TEMPLATES["filter"]
    elif any(word in task_lower for word in ["highlight", "color", "mark"]) and "duplicate" in task_lower:
        return VBA_TEMPLATES["highlight"]
    elif any(word in task_lower for word in ["remove", "delete"]) and "duplicate" in task_lower:
        return VBA_TEMPLATES["remove_duplicates"]
    elif any(word in task_lower for word in ["format", "style", "beautify", "design"]):
        return VBA_TEMPLATES["format"]
    elif any(word in task_lower for word in ["sum", "total", "add", "calculate"]):
        return VBA_TEMPLATES["sum"]
    elif "pivot" in task_lower:
        return VBA_TEMPLATES["pivot"]
    else:
        # Generate custom template
        return f"""Sub CustomMacro()
    ' Task: {task_description}
    ' This is a custom macro template. Modify as needed.
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Your code here
    MsgBox "Task: {task_description}" & vbCrLf & "Please customize this macro for your specific needs.", vbInformation
    
End Sub"""

def handle_excel_upload(file):
    """
    Process uploaded Excel file and provide analytics
    """
    if file is None:
        return "No file uploaded."
    
    try:
        file_path = file.name
        file_name = os.path.basename(file_path)
        
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=0)
        
        # Generate file analytics
        analytics = f"""✅ File '{file_name}' uploaded successfully!

📊 File Analytics:
- Rows: {len(df)}
- Columns: {len(df.columns)}
- Column Names: {', '.join(df.columns.tolist())}
- Memory Usage: {df.memory_usage(deep=True).sum() / 1024:.2f} KB
- Missing Values: {df.isnull().sum().sum()}

📈 Data Preview:
{df.head().to_string()}
"""
        return analytics
    
    except Exception as e:
        return f"❌ Error processing file: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"

def analyze_excel_data(file):
    """
    Provide detailed statistical analysis of Excel data
    """
    if file is None:
        return "No file uploaded for analysis."
    
    try:
        df = pd.read_excel(file.name, sheet_name=0)
        
        # Statistical analysis
        stats = df.describe().to_string()
        dtypes = df.dtypes.to_string()
        
        analysis = f"""📊 Detailed Data Analysis:

🔢 Data Types:
{dtypes}

📈 Statistical Summary:
{stats}

🔍 Data Quality:
- Total Cells: {df.size}
- Non-Null Cells: {df.count().sum()}
- Null Cells: {df.isnull().sum().sum()}
- Duplicate Rows: {df.duplicated().sum()}
"""
        return analysis
    
    except Exception as e:
        return f"❌ Analysis Error: {str(e)}"

# Initialize GitHub client
github_token = os.getenv("GITHUB_TOKEN", "")
g = None
if github_token and github_token != "your_github_token":
    try:
        g = Github(github_token)
    except Exception as e:
        print(f"GitHub initialization warning: {e}")

def push_to_github(repo_name, file_name, code):
    """
    Push generated VBA code to GitHub repository
    """
    if not g:
        return "❌ GitHub token not configured. Please set GITHUB_TOKEN environment variable."
    
    if not repo_name or not file_name or not code:
        return "❌ Please provide repository name, file name, and code."
    
    try:
        # Ensure file has proper extension
        if not file_name.endswith(('.bas', '.vba', '.txt')):
            file_name += '.bas'
        
        repo = g.get_user().get_repo(repo_name)
        
        # Check if file exists
        try:
            contents = repo.get_contents(file_name)
            # File exists, update it
            repo.update_file(
                file_name,
                f"Update VBA macro - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                code,
                contents.sha
            )
            message = f"✅ Successfully updated '{file_name}' in '{repo_name}'"
        except:
            # File doesn't exist, create it
            repo.create_file(
                file_name,
                f"Add VBA macro - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                code
            )
            message = f"✅ Successfully created '{file_name}' in '{repo_name}'"
        
        return message
    
    except Exception as e:
        return f"❌ GitHub push failed: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"

def get_macro_templates_list():
    """Return list of available templates"""
    templates = list(VBA_TEMPLATES.keys())
    return f"Available templates: {', '.join(templates)}"

# Custom CSS for better UI
custom_css = """
#main-title {
    background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    padding: 20px;
    border-radius: 10px;
    color: white;
    text-align: center;
}
.gr-button-primary {
    background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
}
"""

# Build Gradio Interface
with gr.Blocks(css=custom_css, title="ExcelBot Pro") as demo:
    gr.Markdown("""
    <div id="main-title">
        <h1>🤖 ExcelBot Pro - VBA Automation Suite</h1>
        <p>Generate VBA macros, analyze Excel files, and manage code with GitHub integration</p>
    </div>
    """)
    
    with gr.Tabs():
        # Tab 1: VBA Macro Generator
        with gr.Tab("🔧 VBA Generator"):
            gr.Markdown("""
            ### Generate VBA Macros from Natural Language
            Describe your Excel automation task and get a ready-to-use VBA macro!
            
            **Supported Tasks:** sort, filter, highlight duplicates, remove duplicates, format, sum, pivot table
            """)
            
            with gr.Row():
                with gr.Column(scale=3):
                    task_input = gr.Textbox(
                        label="Describe your Excel task",
                        placeholder="e.g., 'Sort data by first column' or 'Remove duplicate rows'",
                        lines=3
                    )
                with gr.Column(scale=1):
                    generate_button = gr.Button("🚀 Generate VBA Macro", variant="primary")
            
            vba_output = gr.Code(
                label="Generated VBA Macro",
                language="vbnet",
                lines=20
            )
            
            gr.Markdown("**💡 Tip:** Copy the macro to Excel's VBA editor (Alt+F11) and run it!")
        
        # Tab 2: Excel File Processor
        with gr.Tab("📊 Excel Analyzer"):
            gr.Markdown("### Upload and Analyze Excel Files")
            
            excel_file = gr.File(
                label="Upload Excel File (.xlsx, .xls)",
                file_types=[".xlsx", ".xls"]
            )
            
            with gr.Row():
                upload_button = gr.Button("📤 Analyze File", variant="primary")
                detailed_button = gr.Button("🔍 Detailed Analysis")
            
            upload_status = gr.Textbox(
                label="File Information",
                lines=15,
                max_lines=20
            )
        
        # Tab 3: GitHub Integration
        with gr.Tab("🐙 GitHub Integration"):
            gr.Markdown("""
            ### Push VBA Code to GitHub
            Save your generated macros to a GitHub repository for version control and sharing.
            
            **Setup:** Set `GITHUB_TOKEN` environment variable with your GitHub Personal Access Token
            """)
            
            with gr.Row():
                repo_name = gr.Textbox(
                    label="Repository Name",
                    placeholder="e.g., username/repo-name"
                )
                file_name = gr.Textbox(
                    label="File Name",
                    placeholder="e.g., my_macro.bas"
                )
            
            code_to_push = gr.Code(
                label="Code to Push (paste or generate first)",
                language="vbnet",
                lines=10
            )
            
            push_button = gr.Button("📤 Push to GitHub", variant="primary")
            push_status = gr.Textbox(label="GitHub Status", lines=5)
        
        # Tab 4: Help & Documentation
        with gr.Tab("❓ Help"):
            gr.Markdown("""
            ## ExcelBot Pro - User Guide
            
            ### 🎯 Features
            1. **VBA Generator**: Create Excel automation macros from natural language descriptions
            2. **Excel Analyzer**: Upload and analyze Excel files with detailed statistics
            3. **GitHub Integration**: Version control your VBA macros
            
            ### 📝 How to Use VBA Macros
            1. Generate a macro using the VBA Generator tab
            2. Open your Excel file
            3. Press `Alt + F11` to open VBA Editor
            4. Insert > Module
            5. Paste the generated code
            6. Press `F5` to run the macro
            
            ### 🔑 GitHub Setup
            1. Create a GitHub Personal Access Token:
               - Go to GitHub Settings > Developer Settings > Personal Access Tokens
               - Generate new token with `repo` permissions
            2. Set environment variable:
               ```bash
               export GITHUB_TOKEN=your_token_here
               ```
            3. Restart the application
            
            ### 🛠️ Available Macro Templates
            - **Sort**: Organize data in ascending/descending order
            - **Filter**: Apply AutoFilter to your data
            - **Highlight Duplicates**: Mark duplicate values with color
            - **Remove Duplicates**: Delete duplicate rows
            - **Format**: Apply professional table formatting
            - **Sum**: Add sum formulas to numeric columns
            - **Pivot Table**: Create pivot table from your data
            
            ### 💻 System Requirements
            - Python 3.7+
            - Microsoft Excel (for running VBA macros)
            - GitHub account (for GitHub integration)
            
            ### 📄 License
            MIT License - Copyright (c) 2025 Mandar Bahadarpurkar
            """)
    
    # Event handlers
    generate_button.click(
        fn=generate_vba_macro,
        inputs=[task_input],
        outputs=[vba_output]
    )
    
    upload_button.click(
        fn=handle_excel_upload,
        inputs=[excel_file],
        outputs=[upload_status]
    )
    
    detailed_button.click(
        fn=analyze_excel_data,
        inputs=[excel_file],
        outputs=[upload_status]
    )
    
    push_button.click(
        fn=push_to_github,
        inputs=[repo_name, file_name, code_to_push],
        outputs=[push_status]
    )

if __name__ == "__main__":
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=False,
        show_error=True
    )
