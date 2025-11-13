
import os
import re
import gradio as gr
from github import Github
import pandas as pd
import openpyxl
from openpyxl import load_workbook

def generate_vba_macro(task_description):
    """Generate VBA code based on task description using pattern matching and templates."""
    if not task_description or not task_description.strip():
        return "' Please provide a task description."
    
    task_lower = task_description.lower()
    
    # Pattern matching for common Excel tasks
    vba_code = f"Sub AutoMacro()\n    ' Task: {task_description}\n"
    
    # Format cells
    if any(word in task_lower for word in ['format', 'color', 'highlight', 'bold', 'italic']):
        if 'bold' in task_lower or 'header' in task_lower:
            vba_code += """    ' Format header row
    With Range("A1").CurrentRegion.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
"""
        if 'color' in task_lower or 'highlight' in task_lower:
            vba_code += """    ' Highlight cells
    Range("A1").CurrentRegion.Interior.Color = RGB(255, 255, 200)
"""
    
    # Sort data
    if any(word in task_lower for word in ['sort', 'order', 'arrange']):
        vba_code += """    ' Sort data by first column
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range("A1").CurrentRegion.Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlYes
"""
    
    # Filter data
    if any(word in task_lower for word in ['filter', 'show only', 'display']):
        vba_code += """    ' Apply filter
    Range("A1").CurrentRegion.AutoFilter
"""
    
    # Calculate/Sum
    if any(word in task_lower for word in ['sum', 'total', 'calculate', 'add']):
        vba_code += """    ' Calculate sum
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lastRow + 1, 1).Value = "Total"
    Cells(lastRow + 1, 2).Formula = "=SUM(B2:B" & lastRow & ")"
"""
    
    # Delete rows/columns
    if any(word in task_lower for word in ['delete', 'remove', 'clear']):
        if 'row' in task_lower:
            vba_code += """    ' Delete empty rows
    Dim i As Long
    For i = Cells(Rows.Count, 1).End(xlUp).Row To 1 Step -1
        If WorksheetFunction.CountA(Rows(i)) = 0 Then
            Rows(i).Delete
        End If
    Next i
"""
        elif 'column' in task_lower:
            vba_code += """    ' Delete empty columns
    Dim i As Long
    For i = Cells(1, Columns.Count).End(xlToLeft).Column To 1 Step -1
        If WorksheetFunction.CountA(Columns(i)) = 0 Then
            Columns(i).Delete
        End If
    Next i
"""
    
    # Copy/Paste
    if any(word in task_lower for word in ['copy', 'paste', 'duplicate']):
        vba_code += """    ' Copy and paste
    Range("A1").CurrentRegion.Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
"""
    
    # Find/Replace
    if any(word in task_lower for word in ['find', 'replace', 'search']):
        vba_code += """    ' Find and replace
    Cells.Replace What:="old", Replacement:="new", LookAt:=xlPart, MatchCase:=False
"""
    
    # Create chart
    if any(word in task_lower for word in ['chart', 'graph', 'plot']):
        vba_code += """    ' Create chart
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects.Add(Left:=100, Top:=100, Width:=400, Height:=300)
    chartObj.Chart.SetSourceData Source:=Range("A1").CurrentRegion
    chartObj.Chart.ChartType = xlColumnClustered
"""
    
    # Default action if no specific pattern matched
    if vba_code.count('\n') <= 2:
        vba_code += """    ' General macro execution
    MsgBox "Task completed: " & """ + f'"{task_description}"' + """
"""
    
    vba_code += "End Sub"
    return vba_code

def handle_excel_upload(file):
    """Process uploaded Excel file and provide analysis."""
    if file is None:
        return "No file uploaded."
    
    try:
        file_path = file.name if hasattr(file, 'name') else file
        
        # Try to read with pandas first
        try:
            df = pd.read_excel(file_path, sheet_name=0, nrows=5)
            num_rows = len(pd.read_excel(file_path, sheet_name=0))
            num_cols = len(df.columns)
            columns = list(df.columns)
            
            # Get workbook info with openpyxl
            wb = load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            
            analysis = f"""✅ File '{os.path.basename(file_path)}' uploaded successfully.

📊 Analysis:
- Sheets: {', '.join(sheet_names)}
- Rows: {num_rows}
- Columns: {num_cols}
- Column names: {', '.join([str(col)[:20] for col in columns[:10]])}{'...' if len(columns) > 10 else ''}
"""
            return analysis
        except Exception as e:
            return f"⚠️ File uploaded but analysis failed: {str(e)}"
            
    except Exception as e:
        return f"❌ Upload failed: {str(e)}"

# Initialize GitHub client
github_token = os.getenv("GITHUB_TOKEN")
if github_token and github_token != "your_github_token":
    try:
        g = Github(github_token)
    except Exception as e:
        g = None
        print(f"GitHub initialization warning: {e}")
else:
    g = None

def push_to_github(repo_name, file_name, code):
    """Push generated VBA code to GitHub repository."""
    if not g:
        return "⚠️ GitHub token not configured. Set GITHUB_TOKEN environment variable."
    
    if not repo_name or not repo_name.strip():
        return "❌ Please provide a repository name."
    
    if not file_name or not file_name.strip():
        return "❌ Please provide a file name."
    
    if not code or not code.strip():
        return "❌ No VBA code to push. Generate a macro first."
    
    # Ensure file has .bas extension for VBA
    if not file_name.endswith('.bas'):
        file_name = file_name + '.bas'
    
    try:
        repo = g.get_user().get_repo(repo_name)
        # Check if file exists
        try:
            contents = repo.get_contents(file_name)
            # Update existing file
            repo.update_file(file_name, "Update VBA macro", code, contents.sha)
            return f"✅ Successfully updated {file_name} in {repo_name}"
        except:
            # Create new file
            repo.create_file(file_name, "Add VBA macro", code)
            return f"✅ Successfully pushed {file_name} to {repo_name}"
    except Exception as e:
        return f"❌ GitHub push failed: {str(e)}"

with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("""
    # 🤖 ExcelBot Pro - VBA Automation Chatbot
    
    Automate Excel tasks by describing what you want to do in plain English. The bot will generate VBA macros for you!
    
    **Features:**
    - 📝 Generate VBA code from natural language descriptions
    - 📊 Analyze uploaded Excel files
    - 🔗 Push generated code to GitHub repositories
    
    **Example tasks:** "Sort data by column A", "Make headers bold", "Calculate sum of column B", "Create a chart"
    """)
    
    with gr.Tabs():
        with gr.TabItem("🔧 Generate VBA Macro"):
            gr.Markdown("### Describe your Excel automation task")
            task_input = gr.Textbox(
                label="Task Description",
                placeholder="e.g., Sort the data by the first column and make headers bold",
                lines=3
            )
            generate_button = gr.Button("Generate VBA Macro", variant="primary")
            vba_output = gr.Code(
                label="Generated VBA Macro",
                language="vbnet",
                lines=20
            )
            generate_button.click(fn=generate_vba_macro, inputs=[task_input], outputs=[vba_output])
        
        with gr.TabItem("📊 Excel File Analysis"):
            gr.Markdown("### Upload an Excel file to analyze its structure")
            excel_file = gr.File(
                label="Upload Excel File (.xls/.xlsx)",
                file_types=[".xlsx", ".xls"]
            )
            upload_status = gr.Textbox(
                label="File Analysis",
                lines=10,
                interactive=False
            )
            excel_file.change(fn=handle_excel_upload, inputs=[excel_file], outputs=[upload_status])
        
        with gr.TabItem("🔗 Push to GitHub"):
            gr.Markdown("### Push your generated VBA macro to a GitHub repository")
            gr.Markdown("**Note:** Make sure you've set the `GITHUB_TOKEN` environment variable.")
            with gr.Row():
                repo_name = gr.Textbox(
                    label="Repository Name",
                    placeholder="my-excel-macros"
                )
                file_name = gr.Textbox(
                    label="File Name",
                    placeholder="macro1",
                    value="AutoMacro.bas"
                )
            push_button = gr.Button("Push to GitHub", variant="primary")
            push_status = gr.Textbox(
                label="GitHub Status",
                lines=3,
                interactive=False
            )
            push_button.click(fn=push_to_github, inputs=[repo_name, file_name, vba_output], outputs=[push_status])

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=7860)
