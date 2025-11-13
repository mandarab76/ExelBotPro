"""
ExcelBot Pro - VBA Automation Chatbot
A Gradio-based interface for generating VBA macros and managing Excel automation tasks.
"""

import os
import re
import gradio as gr
from github import Github
import pandas as pd
from openpyxl import load_workbook
from typing import Optional, Tuple

# Initialize GitHub client
github_token = os.getenv("GITHUB_TOKEN")
g = None
if github_token and github_token != "your_github_token":
    try:
        g = Github(github_token)
    except Exception as e:
        print(f"Warning: GitHub initialization failed: {e}")


def generate_vba_macro(task_description: str) -> str:
    """
    Generate VBA macro code based on task description.
    This is an enhanced version that creates more useful VBA code.
    """
    if not task_description or not task_description.strip():
        return "' Please enter a task description to generate VBA code."
    
    task_lower = task_description.lower()
    
    # Pattern matching for common Excel tasks
    if "format" in task_lower or "style" in task_lower:
        return f"""Sub FormatCells()
    ' Task: {task_description}
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Format selected cells
    With Selection
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Font.Size = 12
    End With
    
    MsgBox "Formatting applied successfully!"
End Sub"""
    
    elif "sum" in task_lower or "total" in task_lower or "calculate" in task_lower:
        return f"""Sub CalculateSum()
    ' Task: {task_description}
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Calculate sum of selected range
    Dim total As Double
    total = Application.WorksheetFunction.Sum(Selection)
    
    ' Display result
    MsgBox "Total: " & Format(total, "#,##0.00")
    
    ' Optionally write to a cell
    ' ws.Range("A1").Value = total
End Sub"""
    
    elif "sort" in task_lower or "order" in task_lower:
        return f"""Sub SortData()
    ' Task: {task_description}
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Sort the selected range
    Selection.Sort Key1:=Selection.Cells(1, 1), Order1:=xlAscending, Header:=xlYes
    
    MsgBox "Data sorted successfully!"
End Sub"""
    
    elif "filter" in task_lower or "autofilter" in task_lower:
        return f"""Sub ApplyFilter()
    ' Task: {task_description}
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Apply autofilter to selected range
    Selection.AutoFilter
    
    MsgBox "Filter applied successfully!"
End Sub"""
    
    elif "delete" in task_lower or "remove" in task_lower:
        return f"""Sub DeleteRows()
    ' Task: {task_description}
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Delete selected rows
    Selection.EntireRow.Delete
    
    MsgBox "Rows deleted successfully!"
End Sub"""
    
    elif "copy" in task_lower or "duplicate" in task_lower:
        return f"""Sub CopyData()
    ' Task: {task_description}
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Copy selected range
    Selection.Copy
    
    ' Paste to destination (modify as needed)
    ' ws.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    MsgBox "Data copied successfully!"
End Sub"""
    
    else:
        # Generic template for other tasks
        return f"""Sub AutoMacro()
    ' Task: {task_description}
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' TODO: Implement your task here
    ' Example: Process selected range
    Dim cell As Range
    For Each cell In Selection
        ' Add your logic here
        ' cell.Value = ProcessCell(cell.Value)
    Next cell
    
    MsgBox "Task completed: {task_description}"
End Sub"""


def handle_excel_upload(file) -> str:
    """
    Handle Excel file upload and provide file information.
    """
    if file is None:
        return "No file uploaded."
    
    try:
        file_path = file.name if hasattr(file, 'name') else str(file)
        
        # Try to read with openpyxl
        try:
            wb = load_workbook(file_path, read_only=True)
            sheet_count = len(wb.sheetnames)
            sheet_names = ", ".join(wb.sheetnames[:5])
            if len(wb.sheetnames) > 5:
                sheet_names += f", ... ({sheet_count} total)"
            
            wb.close()
            return f"✅ File '{os.path.basename(file_path)}' uploaded successfully.\n" \
                   f"📊 Sheets: {sheet_count}\n" \
                   f"📋 Sheet names: {sheet_names}"
        except Exception as e:
            # Fallback to pandas
            try:
                df = pd.read_excel(file_path, nrows=1)
                return f"✅ File '{os.path.basename(file_path)}' uploaded successfully.\n" \
                       f"📊 Columns: {len(df.columns)}\n" \
                       f"📋 Column names: {', '.join(df.columns.astype(str)[:10])}"
            except Exception as e2:
                return f"⚠️ File uploaded but could not be read: {str(e2)}"
    
    except Exception as e:
        return f"❌ Upload failed: {str(e)}"


def push_to_github(repo_name: str, file_name: str, code: str) -> str:
    """
    Push generated VBA code to GitHub repository.
    """
    if not github_token or github_token == "your_github_token":
        return "❌ GitHub token not configured. Please set GITHUB_TOKEN environment variable."
    
    if not g:
        return "❌ GitHub client not initialized. Check your GITHUB_TOKEN."
    
    if not repo_name or not repo_name.strip():
        return "❌ Please provide a repository name."
    
    if not file_name or not file_name.strip():
        return "❌ Please provide a file name."
    
    if not code or not code.strip():
        return "❌ No code to push. Please generate VBA code first."
    
    # Ensure file has .bas extension for VBA modules
    if not file_name.endswith('.bas'):
        file_name = file_name.replace('.vba', '.bas').replace('.vb', '.bas')
        if not file_name.endswith('.bas'):
            file_name += '.bas'
    
    try:
        repo = g.get_user().get_repo(repo_name)
        
        # Check if file already exists
        try:
            contents = repo.get_contents(file_name)
            # File exists, update it
            repo.update_file(file_name, f"Update VBA macro: {file_name}", code, contents.sha)
            return f"✅ Successfully updated {file_name} in {repo_name}"
        except Exception:
            # File doesn't exist, create it
            repo.create_file(file_name, f"Add VBA macro: {file_name}", code)
            return f"✅ Successfully pushed {file_name} to {repo_name}"
    
    except Exception as e:
        error_msg = str(e)
        if "Not Found" in error_msg or "404" in error_msg:
            return f"❌ Repository '{repo_name}' not found. Please check the repository name and ensure it exists."
        elif "Bad credentials" in error_msg or "401" in error_msg:
            return "❌ GitHub authentication failed. Please check your GITHUB_TOKEN."
        else:
            return f"❌ GitHub push failed: {error_msg}"


def validate_github_token() -> Tuple[bool, str]:
    """Validate GitHub token and return status."""
    if not github_token or github_token == "your_github_token":
        return False, "GitHub token not configured"
    
    if not g:
        return False, "GitHub client not initialized"
    
    try:
        user = g.get_user()
        return True, f"✅ Connected as: {user.login}"
    except Exception as e:
        return False, f"❌ Token validation failed: {str(e)}"


# Create Gradio interface
with gr.Blocks(title="ExcelBot Pro", theme=gr.themes.Soft()) as demo:
    gr.Markdown("""
    # 🤖 ExcelBot Pro - VBA Automation Chatbot
    
    Generate VBA macros from natural language descriptions, upload Excel files, and push code to GitHub.
    """)
    
    # GitHub status
    github_status = gr.Markdown(value=validate_github_token()[1])
    
    with gr.Tab("VBA Generator"):
        gr.Markdown("### Describe your Excel task and generate VBA code")
        
        with gr.Row():
            task_input = gr.Textbox(
                label="Task Description",
                placeholder="e.g., Format selected cells as bold with gray background",
                lines=3
            )
        
        generate_button = gr.Button("Generate VBA Macro", variant="primary")
        
        vba_output = gr.Code(
            label="Generated VBA Macro",
            language="vbnet",
            lines=20
        )
        
        with gr.Row():
            clear_button = gr.Button("Clear", variant="secondary")
            copy_button = gr.Button("Copy Code", variant="secondary")
    
    with gr.Tab("Excel Upload"):
        gr.Markdown("### Upload Excel files to analyze structure")
        
        excel_file = gr.File(
            label="Upload Excel File (.xls/.xlsx)",
            file_types=[".xlsx", ".xls"]
        )
        
        upload_status = gr.Textbox(
            label="Upload Status",
            lines=5,
            interactive=False
        )
    
    with gr.Tab("GitHub Integration"):
        gr.Markdown("### Push generated VBA code to GitHub")
        
        with gr.Row():
            repo_name = gr.Textbox(
                label="GitHub Repository Name",
                placeholder="e.g., my-excel-macros"
            )
            file_name = gr.Textbox(
                label="File Name (will be saved as .bas)",
                placeholder="e.g., FormatMacro"
            )
        
        push_button = gr.Button("Push to GitHub", variant="primary")
        push_status = gr.Textbox(
            label="GitHub Status",
            lines=3,
            interactive=False
        )
    
    # Event handlers
    generate_button.click(
        fn=generate_vba_macro,
        inputs=[task_input],
        outputs=[vba_output]
    )
    
    clear_button.click(
        fn=lambda: ("", ""),
        outputs=[task_input, vba_output]
    )
    
    excel_file.change(
        fn=handle_excel_upload,
        inputs=[excel_file],
        outputs=[upload_status]
    )
    
    push_button.click(
        fn=push_to_github,
        inputs=[repo_name, file_name, vba_output],
        outputs=[push_status]
    )


if __name__ == "__main__":
    # Get server configuration from environment
    server_name = os.getenv("SERVER_NAME", "0.0.0.0")
    server_port = int(os.getenv("SERVER_PORT", "7860"))
    share = os.getenv("SHARE", "False").lower() == "true"
    
    demo.launch(
        server_name=server_name,
        server_port=server_port,
        share=share
    )
