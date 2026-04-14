@echo off
cd /d "%~dp0"
if not exist lib mkdir lib

REM PDF.js
curl -fL "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js" -o "lib\pdf.min.js"
curl -fL "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js" -o "lib\pdf.worker.min.js"

REM PptxGenJS
curl -fL "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js" -o "lib\pptxgen.bundle.js"

REM Tesseract.js v5 (OCR エンジン本体)
REM  ※ 日本語言語データ (~10MB) は初回 OCR 実行時にブラウザが自動ダウンロード・キャッシュします
curl -fL "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js" -o "lib\tesseract.min.js"

REM -----------------------------------------------------------------------
REM Tesseract.js 日本語言語データについて
REM
REM   言語データ (jpn.traineddata, ~10MB) は初回 OCR 実行時に
REM   ブラウザが自動的にダウンロード・IndexedDB にキャッシュします。
REM   2回目以降はオフラインでも動作します。
REM
REM   完全オフライン環境で事前にダウンロードしたい場合:
REM     curl -fL "https://tessdata.projectnaptha.com/4.0.0/jpn.traineddata.gz" ^
REM       -o "lib\jpn.traineddata.gz"
REM   ※ 配置後は Tesseract.createWorker の langPath オプション設定が別途必要です
REM -----------------------------------------------------------------------

echo Done.
