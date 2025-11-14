# ExcelBot Pro - Cloud Deployment

This folder contains files configured for deploying ExcelBot Pro to Hugging Face Spaces (free cloud hosting).

## Quick Deploy to Hugging Face Spaces

1. Go to https://huggingface.co/spaces
2. Click "Create new Space"
3. Choose:
   - Space name: excelbot-pro
   - License: MIT
   - Space SDK: **Gradio**
4. Upload these files:
   - `app.py`
   - `requirements.txt`
   - `README.md`
5. Your app will be live at: `https://huggingface.co/spaces/YOUR_USERNAME/excelbot-pro`

## Configure GitHub Token (Optional)

1. Go to your Space settings
2. Click "Repository secrets"
3. Add secret:
   - Name: `GITHUB_TOKEN`
   - Value: Your GitHub personal access token

## Use with Android App

Once deployed, copy your Space URL and configure it in the Android app settings.
