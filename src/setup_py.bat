@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo  pdf2ppt セットアップ
echo  Python 3.11 Embeddable + ライブラリ + surya モデルを準備します
echo  初回のみ実行 / 合計ダウンロード目安: 2-3 GB
echo ============================================================
echo.

set PY_VER=3.11.9
set PY_DIR=%~dp0..\python
set PY_EXE=%PY_DIR%\python.exe
set HF_MODELS_DIR=%PY_DIR%\hf_models

if exist "%PY_EXE%" (
  echo [SKIP] python\ は既に存在します。ライブラリ・モデルの更新のみ行います。
  goto :pip_install
)

echo [1/5] Python %PY_VER% をダウンロード中...
if not exist "%PY_DIR%" mkdir "%PY_DIR%"

powershell -NoProfile -Command ^
  "Invoke-WebRequest 'https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-embed-amd64.zip' -OutFile '%PY_DIR%\py_embed.zip'"
if errorlevel 1 (
  echo [ERROR] ダウンロード失敗。ネットワーク接続を確認してください。
  pause & exit /b 1
)

echo [2/5] 展開中...
powershell -NoProfile -Command ^
  "Expand-Archive '%PY_DIR%\py_embed.zip' '%PY_DIR%' -Force"
del "%PY_DIR%\py_embed.zip"

powershell -NoProfile -Command ^
  "(Get-Content '%PY_DIR%\python311._pth') -replace '#import site','import site' | Set-Content '%PY_DIR%\python311._pth'"

echo [3/5] pip をインストール中...
powershell -NoProfile -Command ^
  "Invoke-WebRequest 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%PY_DIR%\get-pip.py'"
"%PY_EXE%" "%PY_DIR%\get-pip.py" --no-warn-script-location
del "%PY_DIR%\get-pip.py"

:pip_install
echo [4/5] ライブラリをインストール中...
echo.
echo  [4a] PyTorch CPU 版をインストール中...
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
  flask ^
  svglib ^
  "transformers==4.57.3" ^
  surya-ocr

echo.
echo [5/5] surya OCR モデルをダウンロード中...
echo       モデルは python\hf_models\ に保存します
echo.

if not exist "%HF_MODELS_DIR%" mkdir "%HF_MODELS_DIR%"

echo import os > "%PY_DIR%\_dl_models.py"
echo os.environ["HF_HOME"] = r"%HF_MODELS_DIR%" >> "%PY_DIR%\_dl_models.py"
echo print("  [1/2] Detection モデルをダウンロード中...") >> "%PY_DIR%\_dl_models.py"
echo from surya.detection import DetectionPredictor >> "%PY_DIR%\_dl_models.py"
echo DetectionPredictor^(^) >> "%PY_DIR%\_dl_models.py"
echo print("  [2/2] Recognition モデルをダウンロード中...") >> "%PY_DIR%\_dl_models.py"
echo from surya.recognition import RecognitionPredictor >> "%PY_DIR%\_dl_models.py"
echo RecognitionPredictor^(^) >> "%PY_DIR%\_dl_models.py"
echo print("  surya モデルのダウンロード完了") >> "%PY_DIR%\_dl_models.py"

set HF_HOME=%HF_MODELS_DIR%
"%PY_EXE%" "%PY_DIR%\_dl_models.py"
del "%PY_DIR%\_dl_models.py"

echo.
echo [6/6] pdf2ppt パッケージをインストール中...
"%PY_EXE%" -m pip install --no-warn-script-location -e "%~dp0.."

echo.
echo ============================================================
echo  セットアップ完了
echo.
echo  Web UI 起動:  python\Scripts\webpdf2ppt.exe
echo  CLI 実行:     python\Scripts\pdf2ppt.exe input.pdf
echo ============================================================
pause
