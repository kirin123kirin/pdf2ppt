@echo off
cd /d "%~dp0"

where python >nul 2>&1
if not errorlevel 1 (
  python server.py
  exit /b %errorlevel%
)

set HF_HOME=%~dp0python_env\hf_models
python_env\python.exe server.py
