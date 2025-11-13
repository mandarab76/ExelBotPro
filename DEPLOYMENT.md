# Deployment Guide - ExcelBot Pro

## Quick Deployment

### Option 1: Local Deployment

1. Extract the zip file
2. Run the launch script:
   ```bash
   # Linux/macOS
   ./run.sh
   
   # Windows
   run.bat
   ```
3. Open `http://localhost:7860` in your browser

### Option 2: Manual Deployment

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python excelbot_chat.py
```

## Cloud Deployment Options

### Heroku

1. Create a `Procfile`:
   ```
   web: python excelbot_chat.py
   ```

2. Deploy:
   ```bash
   heroku create excelbot-pro
   git push heroku main
   ```

### Docker

Create a `Dockerfile`:

```dockerfile
FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 7860

CMD ["python", "excelbot_chat.py"]
```

Build and run:
```bash
docker build -t excelbot-pro .
docker run -p 7860:7860 excelbot-pro
```

### AWS EC2

1. Launch an EC2 instance (Ubuntu recommended)
2. SSH into the instance
3. Install Python and dependencies
4. Run the application
5. Configure security groups to allow port 7860

### Google Cloud Platform

1. Create a Compute Engine instance
2. Install Python and dependencies
3. Run the application
4. Configure firewall rules

## Environment Variables

Set environment variables for production:

```bash
export GITHUB_TOKEN=your_token
export OPENAI_API_KEY=your_key
```

Or use a `.env` file (requires `python-dotenv`):

```bash
pip install python-dotenv
```

Then create `.env`:
```
GITHUB_TOKEN=your_token
OPENAI_API_KEY=your_key
```

## Production Considerations

1. **Security**: Never commit API keys to version control
2. **Performance**: Use a production WSGI server (Gunicorn, uWSGI)
3. **Monitoring**: Set up logging and monitoring
4. **Backup**: Regular backups of important data
5. **SSL**: Use HTTPS in production (consider using a reverse proxy)

## Reverse Proxy Setup (Nginx)

Example Nginx configuration:

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://localhost:7860;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

## Troubleshooting

### Port Already in Use

Change the port in `excelbot_chat.py`:
```python
demo.launch(server_port=8080)
```

### Memory Issues

For large Excel files, consider:
- Increasing server memory
- Processing files in chunks
- Using streaming for large datasets

### API Rate Limits

- Implement rate limiting
- Cache responses when possible
- Use API keys responsibly

## Support

For deployment issues, check:
- [GitHub Issues](https://github.com/mandarab76/ExcelBotPro/issues)
- Application logs
- System resource usage
