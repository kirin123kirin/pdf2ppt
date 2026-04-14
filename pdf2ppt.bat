@echo off
cd /d "%~dp0"

REM システム Python があればそのまま実行
where python >nul 2>&1
if not errorlevel 1 (
  python pdf2ppt.py %*
  exit /b %errorlevel%
)

REM なければポータブル環境を使用
set HF_HOME=%~dp0python_env\hf_models
python_env\python.exe pdf2ppt.py %*
