@echo off
cd /d "%~dp0"
set HF_HOME=%~dp0python_env\hf_models
python_env\python.exe pdf2ppt.py %*
