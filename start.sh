#!/bin/bash

# Install ffmpeg for Railway.app
echo "Installing ffmpeg..."
apt-get update
apt-get install -y ffmpeg

# Set ffmpeg path for Whisper
export FFMPEG_PATH=$(which ffmpeg)
echo "FFMPEG_PATH set to: $FFMPEG_PATH"

# Start the application
echo "Starting FastAPI application..."
uvicorn main:app --host 0.0.0.0 --port $PORT 