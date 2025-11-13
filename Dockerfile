# ExcelBot Pro - Docker Image
FROM python:3.9-slim

# Set metadata
LABEL maintainer="Mandar Bahadarpurkar"
LABEL description="ExcelBot Pro - VBA Automation Suite"
LABEL version="1.0.0"

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    git \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY excelbot_chat.py .
COPY sample1.xlsx .
COPY sample2.xlsx .
COPY README.md .
COPY LICENSE .

# Create directories for uploads and outputs
RUN mkdir -p /app/uploads /app/outputs

# Set environment variables
ENV PYTHONUNBUFFERED=1
ENV GRADIO_SERVER_NAME=0.0.0.0
ENV GRADIO_SERVER_PORT=7860

# Expose port
EXPOSE 7860

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:7860/ || exit 1

# Run the application
CMD ["python", "excelbot_chat.py"]
