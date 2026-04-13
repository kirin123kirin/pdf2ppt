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
REM OCR モデル (onnx-community/manga-ocr-base) のダウンロード
REM
REM このモデルはゲート付き（HuggingFace アカウント + 利用規約同意が必要）。
REM 以下のいずれかの方法でダウンロードしてください。
REM
REM 【方法1】huggingface-cli（推奨）
REM   事前に: pip install huggingface_hub
REM           huggingface-cli login   ← ブラウザでトークン発行・貼り付け
REM
REM   実行:
REM     huggingface-cli download onnx-community/manga-ocr-base ^
REM       --local-dir lib\models\onnx-community\manga-ocr-base
REM
REM 【方法2】git clone（Git LFS 必須）
REM   事前に: git lfs install
REM   実行:
REM     git clone https://huggingface.co/onnx-community/manga-ocr-base ^
REM       lib\models\onnx-community\manga-ocr-base
REM
REM 【方法3】ブラウザで手動ダウンロード
REM   1. https://huggingface.co/onnx-community/manga-ocr-base/tree/main を開く
REM   2. ログイン → 利用規約に同意
REM   3. 以下のファイルをダウンロードし lib\models\onnx-community\manga-ocr-base\ に配置:
REM        config.json
REM        generation_config.json
REM        preprocessor_config.json
REM        special_tokens_map.json
REM        tokenizer.json
REM        tokenizer_config.json
REM        onnx\encoder_model_quantized.onnx   (または encoder_model.onnx)
REM        onnx\decoder_model_merged_quantized.onnx  (または decoder_model_merged.onnx)
REM
REM ダウンロード済みなら起動後はオフラインで動作します。
REM -----------------------------------------------------------------------

echo Done.
