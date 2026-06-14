@echo off
echo Starting backend...
cd /d "%~dp0backend"
uv run uvicorn main:app --host 0.0.0.0 --port 8000
