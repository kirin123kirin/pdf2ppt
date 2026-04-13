@echo off
cd /d "%~dp0"
if not exist lib mkdir lib
if not exist lib\tessdata mkdir lib\tessdata
curl -fL "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js" -o "lib\pdf.min.js"
curl -fL "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js" -o "lib\pdf.worker.min.js"
curl -fL "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js" -o "lib\pptxgen.bundle.js"
curl -fL "https://unpkg.com/tesseract.js@5/dist/tesseract.min.js" -o "lib\tesseract.min.js"
curl -fL "https://unpkg.com/tesseract.js@5/dist/worker.min.js" -o "lib\tesseract.worker.min.js"
curl -fL "https://unpkg.com/tesseract.js-core@5/tesseract-core-lstm.wasm.js" -o "lib\tesseract-core-lstm.wasm.js"
curl -fL "https://tessdata.projectnaptha.com/4.0.0/jpn.traineddata.gz" -o "lib\tessdata\jpn.traineddata.gz"
curl -fL "https://tessdata.projectnaptha.com/4.0.0/eng.traineddata.gz" -o "lib\tessdata\eng.traineddata.gz"
