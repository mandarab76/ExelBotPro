"""
ExcelBot Pro - VBA Automation Chatbot
A comprehensive tool for generating VBA macros from natural language descriptions
with Excel file analysis and GitHub integration.
"""

import os
import json
import gradio as gr
from github import Github
import pandas as pd
import openpyxl
from typing import Optional, Dict, Any
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configuration
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")


class ExcelAnalyzer:
    """Analyzes Excel files to extract structure and metadata."""
    
    @staticmethod
    def analyze_excel_file(file_path: str) -> Dict[str, Any]:
        """Analyze an Excel file and return its structure."""
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            analysis = {
                "file_name": os.path.basename(file_path),
                "sheets": [],
                "total_sheets": len(wb.sheetnames),
                "sheet_names": wb.sheetnames
            }
            
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                max_row = ws.max_row
                max_col = ws.max_column
                
                # Get sample data from first few rows
                sample_data = []
                for row in range(1, min(6, max_row + 1)):
                    row_data = []
                    for col in range(1, min(6, max_col + 1)):
                        cell_value = ws.cell(row, col).value
                        row_data.append(str(cell_value) if cell_value is not None else "")
                    sample_data.append(row_data)
                
                sheet_info = {
                    "name": sheet_name,
                    "max_row": max_row,
                    "max_column": max_col,
                    "sample_data": sample_data
                }
                analysis["sheets"].append(sheet_info)
            
            wb.close()
            return analysis
        except Exception as e:
            logger.error(f"Error analyzing Excel file: {e}")
            return {"error": str(e)}


class VBAGenerator:
    """Generates VBA macros using AI or template-based approach."""
    
    def __init__(self, use_openai: bool = False):
        self.use_openai = use_openai and bool(OPENAI_API_KEY)
        if self.use_openai:
            try:
                import openai
                self.openai_client = openai.OpenAI(api_key=OPENAI_API_KEY)
            except ImportError:
                logger.warning("OpenAI package not installed. Using template-based generation.")
                self.use_openai = False
    
    def generate_vba_macro(
        self, 
        task_description: str, 
        excel_analysis: Optional[Dict[str, Any]] = None
    ) -> str:
        """Generate VBA macro code from task description."""
        
        if self.use_openai:
            return self._generate_with_openai(task_description, excel_analysis)
        else:
            return self._generate_with_template(task_description, excel_analysis)
    
    def _generate_with_openai(
        self, 
        task_description: str, 
        excel_analysis: Optional[Dict[str, Any]] = None
    ) -> str:
        """Generate VBA macro using OpenAI API."""
        try:
            context = "Generate a VBA macro for Excel that accomplishes the following task. "
            context += f"Task: {task_description}\n\n"
            
            if excel_analysis and "sheets" in excel_analysis:
                context += "Excel File Structure:\n"
                for sheet in excel_analysis["sheets"]:
                    context += f"- Sheet '{sheet['name']}': {sheet['max_row']} rows, {sheet['max_column']} columns\n"
            
            context += "\nProvide only the VBA code, starting with 'Sub' and ending with 'End Sub'. "
            context += "Include proper error handling and comments."
            
            response = self.openai_client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are a VBA expert. Generate clean, well-commented VBA code for Excel automation."},
                    {"role": "user", "content": context}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            return response.choices[0].message.content.strip()
        except Exception as e:
            logger.error(f"OpenAI generation failed: {e}")
            return self._generate_with_template(task_description, excel_analysis)
    
    def _generate_with_template(
        self, 
        task_description: str, 
        excel_analysis: Optional[Dict[str, Any]] = None
    ) -> str:
        """Generate VBA macro using template-based approach."""
        template = f"""Sub AutoMacro()
    ' ExcelBot Pro - Generated VBA Macro
    ' Task: {task_description}
    '
    ' This macro was generated automatically.
    ' Please review and customize as needed.
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Set the active worksheet (modify as needed)
    Set ws = ActiveSheet
    
    ' Get the last row and column with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' TODO: Implement your task logic here
    ' Example: Process data from row 2 to lastRow
    ' For i = 2 To lastRow
    '     ' Your code here
    ' Next i
    
    MsgBox "Macro completed successfully!", vbInformation, "ExcelBot Pro"
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "ExcelBot Pro"
End Sub"""
        
        return template


class GitHubManager:
    """Manages GitHub operations for pushing generated code."""
    
    def __init__(self, token: str):
        self.token = token
        self.github = Github(token) if token else None
    
    def push_to_github(
        self, 
        repo_name: str, 
        file_name: str, 
        code: str,
        commit_message: Optional[str] = None
    ) -> str:
        """Push code to a GitHub repository."""
        if not self.github:
            return "Error: GitHub token not configured. Please set GITHUB_TOKEN environment variable."
        
        try:
            user = self.github.get_user()
            repo = user.get_repo(repo_name)
            
            # Check if file exists
            try:
                contents = repo.get_contents(file_name)
                repo.update_file(
                    file_name,
                    commit_message or f"Update {file_name}",
                    code,
                    contents.sha
                )
                return f"Successfully updated {file_name} in {repo_name}"
            except:
                # File doesn't exist, create it
                repo.create_file(
                    file_name,
                    commit_message or f"Add {file_name}",
                    code
                )
                return f"Successfully created {file_name} in {repo_name}"
        except Exception as e:
            logger.error(f"GitHub push failed: {e}")
            return f"GitHub push failed: {str(e)}"


# Initialize components
excel_analyzer = ExcelAnalyzer()
vba_generator = VBAGenerator(use_openai=bool(OPENAI_API_KEY))
github_manager = GitHubManager(GITHUB_TOKEN)


def generate_vba_macro_handler(task_description: str, excel_file) -> tuple:
    """Handle VBA macro generation request."""
    if not task_description.strip():
        return "", "Please provide a task description."
    
    excel_analysis = None
    if excel_file:
        try:
            excel_analysis = excel_analyzer.analyze_excel_file(excel_file.name)
            if "error" in excel_analysis:
                return "", f"Error analyzing Excel file: {excel_analysis['error']}"
        except Exception as e:
            logger.error(f"Error processing Excel file: {e}")
            return "", f"Error processing Excel file: {str(e)}"
    
    vba_code = vba_generator.generate_vba_macro(task_description, excel_analysis)
    status = "VBA macro generated successfully!"
    if excel_analysis:
        status += f"\nAnalyzed Excel file: {excel_analysis.get('file_name', 'Unknown')}"
        status += f"\nFound {excel_analysis.get('total_sheets', 0)} sheet(s)"
    
    return vba_code, status


def handle_excel_upload(file) -> str:
    """Handle Excel file upload."""
    if not file:
        return "No file uploaded."
    
    try:
        analysis = excel_analyzer.analyze_excel_file(file.name)
        if "error" in analysis:
            return f"Error: {analysis['error']}"
        
        result = f"File '{analysis['file_name']}' uploaded successfully!\n"
        result += f"Sheets: {', '.join(analysis['sheet_names'])}\n"
        result += f"Total sheets: {analysis['total_sheets']}"
        
        return result
    except Exception as e:
        logger.error(f"Upload error: {e}")
        return f"Error processing file: {str(e)}"


def push_to_github_handler(repo_name: str, file_name: str, vba_code: str) -> str:
    """Handle GitHub push request."""
    if not repo_name.strip():
        return "Please provide a repository name."
    if not file_name.strip():
        return "Please provide a file name."
    if not vba_code.strip():
        return "No VBA code to push. Please generate a macro first."
    
    return github_manager.push_to_github(repo_name, file_name, vba_code)


# Create Gradio interface
def create_interface():
    """Create and configure the Gradio interface."""
    with gr.Blocks(title="ExcelBot Pro", theme=gr.themes.Soft()) as demo:
        gr.Markdown("""
        # 🤖 ExcelBot Pro - VBA Automation Chatbot
        
        Generate VBA macros from natural language descriptions, analyze Excel files, 
        and push your code directly to GitHub repositories.
        
        **Features:**
        - 🧠 AI-powered VBA macro generation
        - 📊 Excel file structure analysis
        - 🔗 GitHub integration
        - 💾 Export generated macros
        """)
        
        with gr.Row():
            with gr.Column(scale=2):
                task_input = gr.Textbox(
                    label="Describe your Excel task",
                    placeholder="e.g., Format all cells in column A as currency, or Sort data by date column",
                    lines=3
                )
                generate_button = gr.Button("Generate VBA Macro", variant="primary", size="lg")
                
            with gr.Column(scale=1):
                excel_file = gr.File(
                    label="Upload Excel File (Optional)",
                    file_types=[".xlsx", ".xls"],
                    type="filepath"
                )
                upload_status = gr.Textbox(label="Upload Status", interactive=False)
        
        vba_output = gr.Code(
            label="Generated VBA Macro",
            language="vbnet",
            lines=20,
            interactive=True
        )
        
        status_output = gr.Textbox(
            label="Status",
            interactive=False,
            lines=3
        )
        
        with gr.Row():
            with gr.Column():
                gr.Markdown("### GitHub Integration")
                repo_name = gr.Textbox(
                    label="Repository Name",
                    placeholder="my-excel-automation"
                )
                file_name = gr.Textbox(
                    label="File Name (e.g., macro.bas)",
                    placeholder="macro.bas"
                )
                push_button = gr.Button("Push to GitHub", variant="secondary")
                push_status = gr.Textbox(label="GitHub Status", interactive=False)
        
        # Event handlers
        generate_button.click(
            fn=generate_vba_macro_handler,
            inputs=[task_input, excel_file],
            outputs=[vba_output, status_output]
        )
        
        excel_file.change(
            fn=handle_excel_upload,
            inputs=[excel_file],
            outputs=[upload_status]
        )
        
        push_button.click(
            fn=push_to_github_handler,
            inputs=[repo_name, file_name, vba_output],
            outputs=[push_status]
        )
        
        # Add examples
        gr.Examples(
            examples=[
                ["Format all cells in column A as currency"],
                ["Sort data by the date column in ascending order"],
                ["Delete all rows where column B is empty"],
                ["Apply conditional formatting to highlight values greater than 100"],
                ["Create a summary sheet with totals from all other sheets"]
            ],
            inputs=task_input
        )
    
    return demo


if __name__ == "__main__":
    # Check configuration
    if not GITHUB_TOKEN:
        logger.warning("GITHUB_TOKEN not set. GitHub features will be disabled.")
    if not OPENAI_API_KEY:
        logger.warning("OPENAI_API_KEY not set. Using template-based VBA generation.")
    
    # Launch the interface
    demo = create_interface()
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=False
    )
