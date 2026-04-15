@echo off
cd /d "%~dp0"

where python >nul 2>&1
if not errorlevel 1 (
  python src\server.py
  exit /b %errorlevel%
)

set HF_HOME=%~dp0python\hf_models
python\python.exe src\server.py
