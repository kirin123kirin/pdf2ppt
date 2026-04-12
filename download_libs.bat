@echo off
cd /d "%~dp0"
if not exist lib mkdir lib
curl -fL "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js" -o "lib\pdf.min.js"
curl -fL "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js" -o "lib\pdf.worker.min.js"
curl -fL "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js" -o "lib\pptxgen.bundle.js"
