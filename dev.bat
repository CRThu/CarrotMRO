@echo off
echo Starting dev environment...
cd /d "%~dp0"

echo [1] Starting FastAPI backend on port 8000 (auto-reload)...
start "FastAPI Backend" cmd /k "cd /d %~dp0backend && uv run uvicorn main:app --reload --host 0.0.0.0 --port 8000"

echo [2] Starting Vite frontend dev server...
cd /d "%~dp0frontend"
bun dev

echo Dev servers stopped.
