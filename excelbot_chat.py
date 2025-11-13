
import os
import gradio as gr
from github import Github
import pandas as pd
import openpyxl

# VBA templates for common tasks
VBA_TEMPLATES = {
    "format": """Sub FormatData()
    ' Auto-format data range
    With ActiveSheet.UsedRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    MsgBox "Formatting applied successfully!"
End Sub""",
    
    "sort": """Sub SortData()
    ' Sort data by first column
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ActiveSheet.Range("A1").CurrentRegion.Sort _
        Key1:=Range("A1"), _
        Order1:=xlAscending, _
        Header:=xlYes
    
    MsgBox "Data sorted successfully!"
End Sub""",
    
    "filter": """Sub ApplyAutoFilter()
    ' Apply autofilter to data range
    Dim lastRow As Long, lastCol As Long
    
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ActiveSheet.Range(Cells(1, 1), Cells(lastRow, lastCol)).AutoFilter
    
    MsgBox "AutoFilter applied!"
End Sub""",
    
    "duplicate": """Sub RemoveDuplicates()
    ' Remove duplicate rows
    Dim lastRow As Long, lastCol As Long
    
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ActiveSheet.Range(Cells(1, 1), Cells(lastRow, lastCol)).RemoveDuplicates _
        Columns:=Array(1), Header:=xlYes
    
    MsgBox "Duplicates removed!"
End Sub""",
    
    "sum": """Sub CalculateTotals()
    ' Add sum row at the bottom
    Dim lastRow As Long, lastCol As Long
    
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    For col = 2 To lastCol
        If IsNumeric(Cells(2, col).Value) Then
            Cells(lastRow + 1, col).Formula = "=SUM(" & Cells(2, col).Address & ":" & Cells(lastRow, col).Address & ")"
        End If
    Next col
    
    Cells(lastRow + 1, 1).Value = "TOTAL"
    Cells(lastRow + 1, 1).Font.Bold = True
    
    MsgBox "Totals calculated!"
End Sub""",
    
    "chart": """Sub CreateChart()
    ' Create a column chart from selected data
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set dataRange = ActiveSheet.UsedRange
    Set chartObj = ActiveSheet.ChartObjects.Add(Left:=300, Width:=400, Top:=50, Height:=300)
    
    With chartObj.Chart
        .SetSourceData Source:=dataRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Data Analysis"
    End With
    
    MsgBox "Chart created!"
End Sub"""
}

def generate_vba_macro(task_description):
    """Generate VBA macro based on task description"""
    if not task_description:
        return "' Please enter a task description"
    
    task_lower = task_description.lower()
    
    # Match task description to templates
    if "format" in task_lower or "style" in task_lower:
        return VBA_TEMPLATES["format"]
    elif "sort" in task_lower:
        return VBA_TEMPLATES["sort"]
    elif "filter" in task_lower:
        return VBA_TEMPLATES["filter"]
    elif "duplicate" in task_lower:
        return VBA_TEMPLATES["duplicate"]
    elif "sum" in task_lower or "total" in task_lower or "calculate" in task_lower:
        return VBA_TEMPLATES["sum"]
    elif "chart" in task_lower or "graph" in task_lower:
        return VBA_TEMPLATES["chart"]
    else:
        # Generic macro with task description
        return f"""Sub CustomMacro()
    ' Task: {task_description}
    ' Generated VBA template - customize as needed
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Add your custom code here
    MsgBox "Macro for: {task_description}" & vbCrLf & _
           "Found " & lastRow & " rows of data"
End Sub"""

def handle_excel_upload(file):
    """Handle Excel file upload and provide analysis"""
    if file is None:
        return "No file uploaded"
    
    try:
        # Read the Excel file
        df = pd.read_excel(file.name)
        
        # Generate analysis
        rows, cols = df.shape
        columns_list = ", ".join(df.columns.tolist())
        
        analysis = f"""✅ File uploaded successfully!

📊 **File Analysis:**
- Filename: {os.path.basename(file.name)}
- Rows: {rows}
- Columns: {cols}
- Column names: {columns_list}

💡 Tip: Use the task input to generate VBA macros for this data!"""
        
        return analysis
    except Exception as e:
        return f"❌ Error reading file: {str(e)}"

github_token = os.getenv("GITHUB_TOKEN", "your_github_token")
g = Github(github_token)

def push_to_github(repo_name, file_name, code):
    """Push VBA code to GitHub repository"""
    if not repo_name or not file_name or not code:
        return "❌ Please provide repository name, file name, and generate code first"
    
    if not file_name.endswith('.vba') and not file_name.endswith('.bas'):
        file_name += '.vba'
    
    try:
        repo = g.get_user().get_repo(repo_name)
        
        # Check if file already exists
        try:
            contents = repo.get_contents(file_name)
            # File exists, update it
            repo.update_file(
                file_name,
                f"Update VBA macro - {file_name}",
                code,
                contents.sha
            )
            return f"✅ Successfully updated {file_name} in {repo_name}"
        except:
            # File doesn't exist, create it
            repo.create_file(file_name, f"Add VBA macro - {file_name}", code)
            return f"✅ Successfully created {file_name} in {repo_name}"
    except Exception as e:
        return f"❌ GitHub push failed: {str(e)}\n\nMake sure:\n- GITHUB_TOKEN is set correctly\n- Repository exists\n- You have write access"

with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("""# 🤖 ExcelBot Pro - VBA Automation Chatbot
    Generate VBA macros from natural language, analyze Excel files, and push code to GitHub!
    
    **Supported Tasks:** format, sort, filter, remove duplicates, calculate totals, create charts
    """)

    with gr.Row():
        with gr.Column(scale=3):
            task_input = gr.Textbox(
                label="📝 Describe your Excel task",
                placeholder="Example: 'format my data' or 'sort by first column' or 'create a chart'",
                lines=2
            )
        with gr.Column(scale=1):
            generate_button = gr.Button("🔧 Generate VBA Macro", variant="primary", size="lg")

    vba_output = gr.Code(label="💻 Generated VBA Macro", language="python", lines=15)
    
    gr.Markdown("---")
    gr.Markdown("### 📂 Excel File Upload (Optional)")

    with gr.Row():
        excel_file = gr.File(label="Upload Excel File (.xls/.xlsx)", file_types=[".xls", ".xlsx"])
        upload_status = gr.Textbox(label="📊 File Analysis", lines=8)

    gr.Markdown("---")
    gr.Markdown("### 🚀 Push to GitHub (Optional)")

    with gr.Row():
        repo_name = gr.Textbox(
            label="Repository Name",
            placeholder="username/repo-name"
        )
        file_name = gr.Textbox(
            label="File Name",
            placeholder="my_macro.vba"
        )
    
    with gr.Row():
        push_button = gr.Button("📤 Push to GitHub", variant="secondary")
    
    push_status = gr.Textbox(label="GitHub Status", lines=3)
    
    gr.Markdown("""
    ---
    ### 💡 Quick Tips
    - **No GitHub token?** Set environment variable: `export GITHUB_TOKEN=your_token`
    - **Sample files** are included: `sample1.xlsx` and `sample2.xlsx`
    - **VBA files** can be imported directly into Excel's VBA editor (Alt+F11)
    """)

    # Event handlers
    generate_button.click(fn=generate_vba_macro, inputs=[task_input], outputs=[vba_output])
    excel_file.change(fn=handle_excel_upload, inputs=[excel_file], outputs=[upload_status])
    push_button.click(fn=push_to_github, inputs=[repo_name, file_name, vba_output], outputs=[push_status])

if __name__ == "__main__":
    demo.launch()
