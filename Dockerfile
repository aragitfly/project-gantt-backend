FROM python:3.9-slim

# Install system dependencies including ffmpeg
RUN apt-get update && apt-get install -y \
    ffmpeg \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Set environment variables
ENV ENVIRONMENT=production
ENV FFMPEG_PATH=/usr/bin/ffmpeg
ENV WHISPER_CACHE_DIR=/tmp/whisper_cache
ENV HF_HOME=/tmp/huggingface_cache

# Create cache directories
RUN mkdir -p /tmp/whisper_cache /tmp/huggingface_cache

# Pre-download Whisper model
RUN python -c "import whisper; whisper.load_model('base')"

# Expose port
EXPOSE 8000

# Start the application
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"] 