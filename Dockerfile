FROM python:3.11-slim

# Required system packages
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    tesseract-ocr \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Set work directory
WORKDIR /app

# Create persistent folders
RUN mkdir -p /app/uploaded_conops /app/generated_draws
RUN chmod -R 777 /app/uploaded_conops /app/generated_draws
# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project files
COPY . .

# Expose API port
EXPOSE 10000

# Run FastAPI backend
CMD ["uvicorn", "api_server:app", "--host", "0.0.0.0", "--port", "10000"]
