# Simple Dockerfile for simplified-ai-parser
FROM python:3.12-slim

WORKDIR /app

# Install mupdf-tools for PDF processing
RUN apt-get update && apt-get install -y \
    mupdf-tools \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app/ ./app/
COPY server.py .

# Expose port
EXPOSE 7656

# Run the server
CMD ["python", "server.py"]
