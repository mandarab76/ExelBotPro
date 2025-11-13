# 🚀 ExcelBot Pro - Deployment Guide

This guide covers various deployment options for ExcelBot Pro in production environments.

## 📋 Table of Contents
- [Local Development](#local-development)
- [Production Server](#production-server)
- [Docker Deployment](#docker-deployment)
- [Cloud Platforms](#cloud-platforms)
- [Security Best Practices](#security-best-practices)
- [Performance Optimization](#performance-optimization)

---

## 🖥️ Local Development

### Quick Start
```bash
# Clone and setup
git clone https://github.com/mandarab76/ExcelBotPro.git
cd ExcelBotPro
pip install -r requirements.txt

# Configure environment
cp .env.example .env
# Edit .env with your settings

# Run
python excelbot_chat.py
```

Access at: `http://localhost:7860`

---

## 🌐 Production Server

### Using Gunicorn (Recommended for Linux/Unix)

1. **Install Gunicorn**:
```bash
pip install gunicorn
```

2. **Create systemd service** (`/etc/systemd/system/excelbotpro.service`):
```ini
[Unit]
Description=ExcelBot Pro Service
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/opt/excelbotpro
Environment="GITHUB_TOKEN=your_token_here"
ExecStart=/usr/bin/python3 /opt/excelbotpro/excelbot_chat.py
Restart=always

[Install]
WantedBy=multi-user.target
```

3. **Enable and start**:
```bash
sudo systemctl daemon-reload
sudo systemctl enable excelbotpro
sudo systemctl start excelbotpro
```

### Using Nginx Reverse Proxy

1. **Install Nginx**:
```bash
sudo apt update
sudo apt install nginx
```

2. **Configure Nginx** (`/etc/nginx/sites-available/excelbotpro`):
```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:7860;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        
        # WebSocket support
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
    }
}
```

3. **Enable site**:
```bash
sudo ln -s /etc/nginx/sites-available/excelbotpro /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl restart nginx
```

### SSL with Let's Encrypt

```bash
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d your-domain.com
```

---

## 🐳 Docker Deployment

### Option 1: Basic Docker

1. **Create Dockerfile**:
```dockerfile
FROM python:3.9-slim

# Set working directory
WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Expose port
EXPOSE 7860

# Set environment variables
ENV GRADIO_SERVER_NAME=0.0.0.0
ENV GRADIO_SERVER_PORT=7860

# Run application
CMD ["python", "excelbot_chat.py"]
```

2. **Build and Run**:
```bash
# Build image
docker build -t excelbotpro:latest .

# Run container
docker run -d \
  --name excelbotpro \
  -p 7860:7860 \
  -e GITHUB_TOKEN=your_token_here \
  --restart unless-stopped \
  excelbotpro:latest
```

### Option 2: Docker Compose

1. **Create docker-compose.yml**:
```yaml
version: '3.8'

services:
  excelbotpro:
    build: .
    container_name: excelbotpro
    ports:
      - "7860:7860"
    environment:
      - GITHUB_TOKEN=${GITHUB_TOKEN}
    volumes:
      - ./data:/app/data
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:7860"]
      interval: 30s
      timeout: 10s
      retries: 3
```

2. **Deploy**:
```bash
docker-compose up -d
```

### Docker Best Practices

- Use `.dockerignore` to exclude unnecessary files:
```
__pycache__
*.pyc
*.pyo
*.pyd
.Python
env/
venv/
.git
.gitignore
.env
.env.local
*.log
```

---

## ☁️ Cloud Platforms

### AWS EC2

1. **Launch EC2 instance** (Ubuntu 20.04+)
2. **Install dependencies**:
```bash
sudo apt update
sudo apt install python3-pip python3-venv nginx
```

3. **Setup application**:
```bash
cd /opt
sudo git clone https://github.com/mandarab76/ExcelBotPro.git
cd ExcelBotPro
sudo python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

4. **Configure security group**:
   - Open port 80 (HTTP)
   - Open port 443 (HTTPS)
   - Open port 7860 (Application)

### Google Cloud Platform (Cloud Run)

1. **Create Dockerfile** (use Docker section above)

2. **Deploy**:
```bash
# Build and push to GCR
gcloud builds submit --tag gcr.io/PROJECT-ID/excelbotpro

# Deploy to Cloud Run
gcloud run deploy excelbotpro \
  --image gcr.io/PROJECT-ID/excelbotpro \
  --platform managed \
  --region us-central1 \
  --allow-unauthenticated \
  --set-env-vars GITHUB_TOKEN=your_token
```

### Heroku

1. **Create Procfile**:
```
web: python excelbot_chat.py
```

2. **Create runtime.txt**:
```
python-3.9.18
```

3. **Deploy**:
```bash
heroku create your-app-name
heroku config:set GITHUB_TOKEN=your_token
git push heroku main
```

### Azure App Service

1. **Create Web App** in Azure Portal
2. **Configure deployment**:
```bash
az webapp up --name excelbotpro --resource-group myResourceGroup
az webapp config appsettings set --name excelbotpro \
  --settings GITHUB_TOKEN=your_token
```

### DigitalOcean

1. **Create Droplet** (Ubuntu 20.04)
2. **Setup similar to AWS EC2**
3. **Optional: Use App Platform**:
   - Connect GitHub repository
   - Set environment variables
   - Deploy automatically

---

## 🔒 Security Best Practices

### Environment Variables
```bash
# Never commit .env files
echo ".env" >> .gitignore

# Use strong tokens
GITHUB_TOKEN=ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

### Application Security

1. **Enable HTTPS** (use Let's Encrypt)
2. **Set secure headers** in Nginx:
```nginx
add_header X-Frame-Options "SAMEORIGIN" always;
add_header X-Content-Type-Options "nosniff" always;
add_header X-XSS-Protection "1; mode=block" always;
```

3. **Rate limiting** in Nginx:
```nginx
limit_req_zone $binary_remote_addr zone=excelbotlimit:10m rate=10r/s;
limit_req zone=excelbotlimit burst=20 nodelay;
```

4. **Firewall rules**:
```bash
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
sudo ufw enable
```

### Access Control

For private deployments, add authentication:

```python
# In excelbot_chat.py
demo.launch(
    auth=("username", "password"),  # Basic auth
    # or
    auth_message="Login to ExcelBot Pro"
)
```

---

## ⚡ Performance Optimization

### 1. Caching

Enable Gradio caching:
```python
demo.launch(
    enable_queue=True,
    max_threads=10
)
```

### 2. Resource Limits

Set memory limits in Docker:
```bash
docker run -d \
  --memory="2g" \
  --cpus="2" \
  excelbotpro:latest
```

### 3. Load Balancing

Use multiple instances with Nginx:
```nginx
upstream excelbotpro {
    server 127.0.0.1:7860;
    server 127.0.0.1:7861;
    server 127.0.0.1:7862;
}

server {
    location / {
        proxy_pass http://excelbotpro;
    }
}
```

### 4. CDN

For static assets, use a CDN like Cloudflare:
1. Point DNS to Cloudflare
2. Enable caching rules
3. Enable "Always Online"

### 5. Monitoring

Install monitoring tools:
```bash
# Prometheus + Grafana
docker run -d -p 9090:9090 prom/prometheus
docker run -d -p 3000:3000 grafana/grafana
```

---

## 📊 Health Checks

Add a health check endpoint:

```python
# In excelbot_chat.py
import gradio as gr
from fastapi import FastAPI
from fastapi.responses import JSONResponse

app = FastAPI()

@app.get("/health")
def health_check():
    return JSONResponse({"status": "healthy"})

# Mount Gradio app
app = gr.mount_gradio_app(app, demo, path="/")
```

---

## 🔧 Troubleshooting

### Port Already in Use
```bash
# Find process using port 7860
lsof -i :7860

# Kill process
kill -9 PID
```

### Memory Issues
```bash
# Increase swap space
sudo fallocate -l 2G /swapfile
sudo chmod 600 /swapfile
sudo mkswap /swapfile
sudo swapon /swapfile
```

### Permission Errors
```bash
# Fix file permissions
sudo chown -R $USER:$USER /opt/excelbotpro
chmod +x excelbot_chat.py
```

---

## 📝 Deployment Checklist

- [ ] Environment variables configured
- [ ] SSL certificate installed
- [ ] Firewall rules configured
- [ ] Backup strategy in place
- [ ] Monitoring setup
- [ ] Error logging enabled
- [ ] Auto-restart configured
- [ ] Domain DNS configured
- [ ] Load testing completed
- [ ] Documentation updated

---

## 📧 Support

For deployment issues:
- 📖 [Documentation](https://github.com/mandarab76/ExcelBotPro)
- 🐛 [Issue Tracker](https://github.com/mandarab76/ExcelBotPro/issues)
- 💬 [Discussions](https://github.com/mandarab76/ExcelBotPro/discussions)

---

**Last Updated**: 2025-11-13  
**Version**: 1.0.0
