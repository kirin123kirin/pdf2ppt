#!/usr/bin/env python3
"""
server.py — pdf2ppt Web UI サーバー

起動: python server.py
ブラウザが自動で開きます。Ctrl+C で停止。

対応入力: PDF / JPG / PNG / BMP / WebP / TIFF / SVG / クリップボード画像
"""

import gc
import io
import os
import sys
import uuid
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template_string
from PIL import Image

# ── pdf2ppt.py を同ディレクトリからインポート ─────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
import pdf2ppt

# ── surya モデルの事前ロード（サーバー起動直後にバックグラウンドで実行）────
_det = None
_rec = None
_model_ready = threading.Event()

def _preload_models():
    global _det, _rec
    print("[server] surya モデルをロード中... (初回のみ数十秒かかります)")
    try:
        _det, _rec = pdf2ppt.load_models()
        # load_models() を monkey-patch して変換ごとの再ロードを防ぐ
        pdf2ppt.load_models = lambda verbose=False: (_det, _rec)
        print("[server] surya モデル準備完了")
    except Exception as e:
        print(f"[server][WARN] モデルロード失敗: {e}")
    _model_ready.set()

threading.Thread(target=_preload_models, daemon=True).start()

# ── 生成 PPTX の一時保持 {token: Path} ──────────────────────────────────────
_pptx_store: dict[str, Path] = {}

ALLOWED_EXT = {
    '.pdf',
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff', '.tif',
    '.svg',
}

# ── Flask ────────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB

# ── HTML（埋め込み）────────────────────────────────────────────────────────
_HTML = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PDF → PowerPoint 変換</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: "メイリオ", Meiryo, "Yu Gothic", sans-serif;
  background: #f5f7fa; color: #1f2937;
  min-height: 100vh; display: flex; flex-direction: column;
  align-items: center; padding: 32px 16px;
}
h1 { font-size: 1.6rem; margin-bottom: 8px; color: #2563eb; }
.subtitle { font-size: 0.85rem; color: #6b7280; margin-bottom: 28px; }

#drop-zone {
  width: 100%; max-width: 640px;
  border: 2px dashed #9ca3af; border-radius: 12px;
  padding: 48px 24px; text-align: center; cursor: pointer;
  transition: all .2s; background: #fff; color: #6b7280; font-size: 1rem;
}
#drop-zone:hover, #drop-zone.drag-over {
  border-color: #2563eb; background: #eff6ff; color: #2563eb;
}
#drop-zone .icon { font-size: 3rem; display: block; margin-bottom: 12px; }
#file-input { display: none; }

.progress-wrap { width: 100%; max-width: 640px; margin-top: 24px; display: none; }
.progress-bar-bg {
  background: #e5e7eb; border-radius: 6px; height: 10px;
  overflow: hidden; margin-bottom: 8px;
}
.progress-bar {
  height: 100%; background: linear-gradient(90deg, #2563eb, #60a5fa);
  border-radius: 6px; width: 0%; transition: width .4s;
}
.progress-text { font-size: 0.85rem; color: #6b7280; text-align: right; }

#log {
  width: 100%; max-width: 640px; margin-top: 16px;
  background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px;
  padding: 16px; font-family: Consolas, monospace; font-size: 0.8rem;
  height: 200px; overflow-y: auto; display: none; line-height: 1.6; color: #166534;
}
#log .err  { color: #dc2626; }
#log .info { color: #1d4ed8; }
#log .warn { color: #92400e; }

.btn-wrap { margin-top: 20px; display: none; }
#btn-download {
  background: #2563eb; color: #fff; border: none; border-radius: 8px;
  padding: 10px 28px; font-size: 1rem; font-weight: bold;
  cursor: pointer; font-family: inherit;
}
#btn-download:hover { background: #1d4ed8; }

.model-badge {
  width: 100%; max-width: 640px; margin-top: 10px;
  font-size: 0.78rem; color: #9ca3af; text-align: right;
}
.model-badge.ready { color: #166534; }
</style>
</head>
<body>
<h1>PDF → PowerPoint 変換</h1>
<p class="subtitle">
  ファイルをドロップ / クリックして選択 / <kbd>Ctrl+V</kbd> でクリップボードから貼り付け
</p>

<div id="drop-zone">
  <span class="icon">📄</span>
  PDF・画像ファイルをここにドロップ<br>
  <small style="color:#9ca3af;margin-top:8px;display:block">
    PDF / JPG / PNG / BMP / WebP / TIFF / SVG
  </small>
</div>
<input type="file" id="file-input" accept=".pdf,image/*,.svg">

<div class="progress-wrap" id="progress-wrap">
  <div class="progress-bar-bg"><div class="progress-bar" id="progress-bar"></div></div>
  <div class="progress-text" id="progress-text">変換中...</div>
</div>

<div id="log"></div>
<div class="model-badge" id="model-badge">OCR モデル確認中...</div>

<div class="btn-wrap" id="btn-wrap">
  <button id="btn-download">⬇ PPTX をダウンロード</button>
</div>

<script>
const dropZone     = document.getElementById('drop-zone');
const fileInput    = document.getElementById('file-input');
const logEl        = document.getElementById('log');
const progressWrap = document.getElementById('progress-wrap');
const progressBar  = document.getElementById('progress-bar');
const progressText = document.getElementById('progress-text');
const btnWrap      = document.getElementById('btn-wrap');
const modelBadge   = document.getElementById('model-badge');

let currentToken = null, currentFilename = 'output.pptx';

// ── モデル状態ポーリング ──────────────────────────────────────────────────
async function pollModelStatus() {
  try {
    const r = await fetch('/model_status');
    const d = await r.json();
    if (d.ready) {
      modelBadge.textContent = '✓ OCR モデル準備完了';
      modelBadge.className = 'model-badge ready';
    } else {
      modelBadge.textContent = '⏳ OCR モデルロード中...';
      setTimeout(pollModelStatus, 2000);
    }
  } catch { setTimeout(pollModelStatus, 3000); }
}
pollModelStatus();

// ── ユーティリティ ─────────────────────────────────────────────────────────
function log(msg, cls = '') {
  logEl.style.display = 'block';
  const d = document.createElement('div');
  if (cls) d.className = cls;
  d.textContent = msg;
  logEl.appendChild(d);
  logEl.scrollTop = logEl.scrollHeight;
}
function setProgress(pct, label) {
  progressWrap.style.display = 'block';
  progressBar.style.width = pct + '%';
  progressText.textContent = label;
}
function reset() {
  logEl.innerHTML = ''; logEl.style.display = 'none';
  progressWrap.style.display = 'none'; progressBar.style.width = '0%';
  btnWrap.style.display = 'none'; currentToken = null;
}

// ── ファイル送信 ──────────────────────────────────────────────────────────
async function sendFile(file) {
  reset();
  const label = file.name || '(クリップボード画像)';
  log(`処理開始: ${label}`, 'info');
  setProgress(15, 'アップロード中...');

  const fd = new FormData();
  fd.append('file', file, file.name || 'clipboard.png');

  try {
    setProgress(40, 'OCR・変換中（数十秒かかります）...');
    const r = await fetch('/convert', { method: 'POST', body: fd });
    const d = await r.json();
    if (!r.ok || d.error) throw new Error(d.error || '変換エラー');
    currentToken = d.token; currentFilename = d.filename;
    setProgress(100, '完了');
    log(`完了: ${d.filename}  (${d.size_kb} KB)`, 'info');
    btnWrap.style.display = 'block';
  } catch(e) {
    progressWrap.style.display = 'none';
    log('[ERROR] ' + e.message, 'err');
  }
}

// ── イベント ────────────────────────────────────────────────────────────
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', e => { if (e.target.files[0]) sendFile(e.target.files[0]); });
dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) sendFile(e.dataTransfer.files[0]);
});
document.addEventListener('paste', e => {
  for (const item of e.clipboardData?.items ?? []) {
    if (item.type.startsWith('image/')) { sendFile(item.getAsFile()); break; }
  }
});
document.getElementById('btn-download').addEventListener('click', () => {
  if (!currentToken) return;
  const a = document.createElement('a');
  a.href = `/download/${currentToken}`; a.download = currentFilename; a.click();
});
</script>
</body>
</html>"""


# ── ルート ────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(_HTML)


@app.route('/model_status')
def model_status():
    return jsonify({'ready': _model_ready.is_set()})


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルが添付されていません'}), 400

    f      = request.files['file']
    name   = f.filename or 'upload.png'
    suffix = Path(name).suffix.lower()

    if suffix not in ALLOWED_EXT:
        return jsonify({'error': f'非対応形式: {suffix}'}), 400

    # モデルが準備できるまで待機（最大 120 秒）
    if not _model_ready.wait(timeout=120):
        return jsonify({'error': 'OCR モデルのロードがタイムアウトしました'}), 503

    tmp_dir   = Path(tempfile.mkdtemp(prefix='pdf2ppt_'))
    pptx_path = tmp_dir / (Path(name).stem + '.pptx')

    try:
        if suffix == '.pdf':
            src = tmp_dir / name
            f.save(str(src))
            pdf2ppt.process_pdf(
                src, tmp_dir,
                pptx_path=pptx_path, csv_path=None, verbose=False,
            )
        else:
            img = _load_image(f.read(), suffix)
            if img is None:
                return jsonify({'error': f'{suffix} を開けませんでした。SVG は svglib/cairosvg が必要です'}), 400
            pdf2ppt.process_image(
                img, tmp_dir,
                pptx_path=pptx_path, csv_path=None, verbose=False,
            )
            img.close()

        size_kb = round(pptx_path.stat().st_size / 1024)
        token   = str(uuid.uuid4())
        _pptx_store[token] = pptx_path
        return jsonify({'token': token, 'filename': pptx_path.name, 'size_kb': size_kb})

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        gc.collect()


def _load_image(data: bytes, suffix: str) -> "Image.Image | None":
    """バイト列から PIL Image を返す。SVG は PNG に変換して対応。"""
    if suffix == '.svg':
        # cairosvg を優先、なければ svglib を試みる
        try:
            import cairosvg
            png = cairosvg.svg2png(bytestring=data)
            return Image.open(io.BytesIO(png)).convert('RGB')
        except ImportError:
            pass
        try:
            from svglib.svglib import svg2rlg
            from reportlab.graphics import renderPM
            import tempfile as _tf, os as _os
            with _tf.NamedTemporaryFile(suffix='.svg', delete=False) as t:
                t.write(data); t.flush(); svg_tmp = t.name
            drawing = svg2rlg(svg_tmp)
            _os.unlink(svg_tmp)
            buf = io.BytesIO()
            renderPM.drawToFile(drawing, buf, fmt='PNG')
            buf.seek(0)
            return Image.open(buf).convert('RGB')
        except Exception:
            return None
    try:
        return Image.open(io.BytesIO(data)).convert('RGB')
    except Exception:
        return None


@app.route('/download/<token>')
def download(token):
    path = _pptx_store.get(token)
    if path is None or not path.exists():
        return 'Not found', 404
    return send_file(str(path), as_attachment=True, download_name=path.name)


# ── エントリーポイント ────────────────────────────────────────────────────

if __name__ == '__main__':
    port = 8765
    url  = f'http://127.0.0.1:{port}'
    print(f'[server] 起動: {url}')
    print('[server] 停止するには Ctrl+C を押してください')
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    app.run(host='127.0.0.1', port=port, debug=False, threaded=False)
