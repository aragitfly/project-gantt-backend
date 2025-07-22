# Project Gantt Chart Manager - Backend

FastAPI backend for the Project Gantt Chart Manager application with audio processing capabilities.

## Features

- Excel file upload and parsing
- Audio recording and transcription using OpenAI Whisper
- Dutch language support
- Task proposal generation
- Meeting summary generation
- CORS support for frontend integration

## Deployment to Railway.app

### Prerequisites

1. Railway.app account
2. Git repository with this code

### Deployment Steps

1. **Connect to Railway.app**
   - Go to [Railway.app](https://railway.app)
   - Create a new project
   - Connect your GitHub repository

2. **Environment Variables**
   Set the following environment variables in Railway.app:
   ```
   ENVIRONMENT=production
   PORT=8000
   ```

3. **Deploy**
   - Railway.app will automatically detect the Python project
   - It will use the `Procfile` to start the application
   - The `start.sh` script will install ffmpeg automatically

### Local Development

1. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Install ffmpeg**
   ```bash
   # macOS
   brew install ffmpeg
   
   # Or download the binary (already included in this repo)
   chmod +x ffmpeg
   ```

3. **Run the server**
   ```bash
   uvicorn main:app --host 127.0.0.1 --port 8000 --reload
   ```

## API Endpoints

- `GET /` - Health check
- `POST /upload-excel` - Upload and parse Excel file
- `POST /process-audio` - Process audio and generate transcript
- `POST /update-excel` - Update Excel file with changes
- `GET /download-excel` - Download updated Excel file

## Environment Variables

- `ENVIRONMENT` - Set to "production" for Railway.app deployment
- `PORT` - Port number (Railway.app sets this automatically)
- `FFMPEG_PATH` - Path to ffmpeg binary (set automatically in production) 