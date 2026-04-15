@echo off
cd /d "%~dp0"

where python >nul 2>&1
if not errorlevel 1 (
  python pdf2ppt.py %*
  exit /b %errorlevel%
)

set HF_HOME=%~dp0..\python\hf_models
..\python\python.exe pdf2ppt.py %*
