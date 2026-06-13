@echo off
echo Starting backend...
:: Set the working directory to the project root
cd /d "%~dp0"
:: Use uv to run the backend
uv run backend\main.py