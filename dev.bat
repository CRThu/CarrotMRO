@echo off
echo Starting dev environment...
cd /d "%~dp0"

echo [0] Cleaning up old processes on ports 8000 and 5173...
bunx kill-port 8000 2>nul
bunx kill-port 5173 2>nul
timeout /t 1 /nobreak >nul

echo [1] Starting FastAPI backend on port 8000 (auto-reload)...
start "FastAPI Backend" cmd /k "cd /d %~dp0backend && uv run uvicorn main:app --reload --host 0.0.0.0 --port 8000"

echo [2] Waiting for backend to start...
:wait_backend
timeout /t 1 /nobreak >nul
curl -s http://localhost:8000/docs >nul 2>&1
if errorlevel 1 goto wait_backend
echo Backend is ready.

echo [3] Starting Vite frontend dev server...
cd /d "%~dp0frontend"
bun dev

echo Dev servers stopped.
