@echo off
cd /d "%~dp0"
if not exist lib mkdir lib

REM PDF.js
curl -fL "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js" -o "lib\pdf.min.js"
curl -fL "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js" -o "lib\pdf.worker.min.js"

REM PptxGenJS
curl -fL "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js" -o "lib\pptxgen.bundle.js"

REM Transformers.js v3 (OCR エンジン本体)
curl -fL "https://cdn.jsdelivr.net/npm/@huggingface/transformers@3/dist/transformers.min.js" -o "lib\transformers.min.js"

REM -----------------------------------------------------------------------
REM manga-ocr-base モデルファイルの事前ダウンロードについて
REM   Transformers.js はモデルを初回使用時に自動的にブラウザの Cache API
REM   (Service Worker) へ保存するため、通常は追加 DL 不要です。
REM
REM   完全オフライン(air-gapped)環境が必要な場合は、以下のコマンドで
REM   HuggingFace CLI を使って ONNX モデルを lib\models\ に保存してください:
REM
REM     pip install huggingface_hub
REM     huggingface-cli download onnx-community/manga-ocr-base --local-dir lib\models\onnx-community\manga-ocr-base
REM -----------------------------------------------------------------------

echo Done.
