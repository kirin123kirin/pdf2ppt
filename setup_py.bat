@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo  pdf2ppt.py 用 ポータブル Python 環境 セットアップ
echo  Python 3.11 Embeddable + 必要ライブラリをインストールします
echo  ※ 初回のみ実行 / 合計ダウンロード目安: 1〜2 GB
echo ============================================================
echo.

set PY_VER=3.11.9
set PY_DIR=%~dp0python_env
set PY_EXE=%PY_DIR%\python.exe

REM ── Step 1: Python Embeddable をダウンロード・展開 ──────────────────────────
if exist "%PY_EXE%" (
  echo [SKIP] python_env\ は既に存在します。ライブラリの更新のみ行います。
  goto :pip_install
)

echo [1/4] Python %PY_VER% Embeddable Package をダウンロード中...
if not exist "%PY_DIR%" mkdir "%PY_DIR%"

powershell -NoProfile -Command ^
  "Invoke-WebRequest 'https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-embed-amd64.zip' -OutFile '%PY_DIR%\py_embed.zip'"
if errorlevel 1 (
  echo [ERROR] ダウンロード失敗。ネットワーク接続を確認してください。
  pause & exit /b 1
)

echo [2/4] 展開中...
powershell -NoProfile -Command ^
  "Expand-Archive '%PY_DIR%\py_embed.zip' '%PY_DIR%' -Force"
del "%PY_DIR%\py_embed.zip"

REM python311._pth の "#import site" を有効化（pip / site-packages を使えるようにする）
powershell -NoProfile -Command ^
  "(Get-Content '%PY_DIR%\python311._pth') -replace '#import site','import site' | Set-Content '%PY_DIR%\python311._pth'"

REM ── Step 3: pip をインストール ──────────────────────────────────────────────
echo [3/4] pip をインストール中...
powershell -NoProfile -Command ^
  "Invoke-WebRequest 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%PY_DIR%\get-pip.py'"
"%PY_EXE%" "%PY_DIR%\get-pip.py" --no-warn-script-location
del "%PY_DIR%\get-pip.py"

REM ── Step 4: ライブラリをインストール ────────────────────────────────────────
:pip_install
echo [4/4] ライブラリをインストール中...
echo.
echo  [4a] PyTorch CPU 版をインストール中（GPU 版回避でサイズを削減）...
"%PY_EXE%" -m pip install --no-warn-script-location ^
  torch torchvision --index-url https://download.pytorch.org/whl/cpu

echo.
echo  [4b] その他のライブラリをインストール中...
"%PY_EXE%" -m pip install --no-warn-script-location ^
  numpy ^
  scipy ^
  pymupdf ^
  pillow ^
  janome ^
  python-pptx ^
  "transformers==4.57.3" ^
  surya-ocr

echo.
echo ============================================================
echo  セットアップ完了！
echo.
echo  実行方法:
echo    python_env\python.exe pdf2ppt.py input.pdf
echo    python_env\python.exe pdf2ppt.py        (クリップボードから)
echo.
echo  ※ 初回実行時に surya OCR モデル (数百 MB) が自動ダウンロードされます
echo ============================================================
pause
