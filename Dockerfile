# Use official Python runtime as base image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies if needed
RUN apt-get update && apt-get install -y \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements-service.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements-service.txt

# Copy application code
COPY service.py .
COPY index.html .

# Expose port (Railway will set PORT env variable)
EXPOSE 8000

# Run the application
CMD ["python", "service.py"]
