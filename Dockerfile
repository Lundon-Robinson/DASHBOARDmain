FROM python:3.12-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    build-essential \
    pkg-config \
    libgl1-mesa-glx \
    libglib2.0-0 \
    libxext6 \
    libxrender-dev \
    libgomp1 \
    libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create necessary directories
RUN mkdir -p /app/logs /app/data /app/temp /app/exports

# Set environment variables
ENV PYTHONPATH=/app
ENV DISPLAY=:0

# Expose ports for web interface
EXPOSE 8080 8081

# Default command
CMD ["python", "main_dashboard.py"]