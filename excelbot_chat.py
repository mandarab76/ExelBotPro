
import os
import gradio as gr
from github import Github


def generate_vba_macro(task_description):
    return f"""Sub AutoMacro()
    ' Task: {task_description}
    MsgBox "This is a placeholder macro for: {task_description}"
End Sub"""


def handle_excel_upload(file):
    return f"File '{file.name}' uploaded successfully."


github_token = os.getenv("GITHUB_TOKEN", "your_github_token")
g = Github(github_token)


def push_to_github(repo_name, file_name, code):
    try:
        repo = g.get_user().get_repo(repo_name)
        repo.create_file(file_name, "Add VBA macro", code)
        return f"Successfully pushed {file_name} to {repo_name}"
    except Exception as e:
        return f"GitHub push failed: {str(e)}"


with gr.Blocks() as demo:
    gr.Markdown("## 🤖 ExcelBot Pro - VBA Automation Chatbot")

    with gr.Row():
        task_input = gr.Textbox(label="Describe your Excel task")
        generate_button = gr.Button("Generate VBA Macro")

    vba_output = gr.Code(label="Generated VBA Macro", language="vbnet")

    with gr.Row():
        excel_file = gr.File(label="Upload Excel File (.xls/.xlsx)")
        upload_status = gr.Textbox(label="Upload Status")

    with gr.Row():
        repo_name = gr.Textbox(label="GitHub Repository Name")
        file_name = gr.Textbox(label="File Name to Push")
        push_button = gr.Button("Push to GitHub")
        push_status = gr.Textbox(label="GitHub Status")

    generate_button.click(fn=generate_vba_macro, inputs=[task_input], outputs=[vba_output])
    excel_file.change(fn=handle_excel_upload, inputs=[excel_file], outputs=[upload_status])
    push_button.click(fn=push_to_github, inputs=[repo_name, file_name, vba_output], outputs=[push_status])


demo.launch()
