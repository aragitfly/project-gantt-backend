#!/bin/bash

# Update package list
apt-get update

# Install ffmpeg and other dependencies
apt-get install -y ffmpeg

# Set environment variables for better performance
export WHISPER_CACHE_DIR="/tmp/whisper_cache"
export HF_HOME="/tmp/huggingface_cache"

# Create cache directories
mkdir -p $WHISPER_CACHE_DIR
mkdir -p $HF_HOME

# Install Python dependencies
pip install -r requirements.txt

# Pre-download the large-v3 model to cache (this helps with deployment)
echo "Pre-downloading Whisper large-v3 model..."
python -c "import whisper; whisper.load_model('large-v3')"

# Start the application
uvicorn main:app --host 0.0.0.0 --port $PORT 