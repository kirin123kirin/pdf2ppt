# CLAUDE.md — 開発ガイド

## Git ワークフロー

**変更は必ず `main` ブランチに直接 push すること。**

```bash
git add <files>
git commit -m "..."
git push origin main
```

フィーチャーブランチは使用しない。

## 概要

`pdf2ppt.html` は単一 HTML ファイルで動作する PDF→PPTX 変換ツールです。  
外部サーバー通信なし、依存フレームワークなし、バンドラー不要です。

## ファイル構成

```
pdf2ppt.html          メインの変換ツール（HTML + CSS + JS をすべて内包）
pdf2ppt.py            Python 版（アルゴリズム参照実装）
download_libs.bat     Windows 用ライブラリダウンロードスクリプト
lib/                  ローカルライブラリ置き場（download_libs.bat で生成）
  pdf.min.js
  pdf.worker.min.js
  pptxgen.bundle.js
```

## 重要な定数（`pdf2ppt.py` と完全一致を維持すること）

```javascript
const RASTER_TARGET_WIDTH = 2048;   // PDF ラスタライズ幅
const NON_WHITE_THR       = 238;    // 非白ピクセル判定閾値
const BLOCK_GAP           = 10;     // ブロック統合ギャップ (px)
const MIN_BLOCK_AREA      = 50;     // 最小ブロック面積 (px²)
const TEXT_PAD_RATIO      = 0.15;   // テキストマスク周囲パディング率
const BG_MARGIN           = 8;      // 背景色サンプリングマージン (px)
```

## アーキテクチャ

### 座標系

- PDF 座標: ポイント単位 (pt)、Y 軸は下向き
- キャンバス座標: ピクセル単位 (px)、Y 軸は上向き
- PPTX 座標: インチ単位 (in)

変換式:
```
scale = RASTER_TARGET_WIDTH / page_width_pt
px    = pt * scale
in    = pt / 72
```

### 処理フロー（`processPDF` 関数）

```
1. PDF ロード (pdf.js)
2. PaddleOCR 初期化（一度だけ）
3. ページループ:
   a. ページをキャンバスにラスタライズ → origData = getImageData (1回のみ)
   b. PDF 内蔵テキスト抽出 (extractPdfNativeText)
   c. 内蔵テキストがなければ PaddleOCR (ocrModel.recognize)
   d. テキスト領域をマスク (maskText) ← origData.data を複製して加工
   e. 残った画像をブロック検出 (extractBlocks) ← マスク後のキャンバスから
   f. スライド構築 (pptx.addSlide)
4. PPTX 出力 (pptx.write)
```

### 主要関数

| 関数 | 役割 |
|------|------|
| `detectBlocks` | 行/列プロジェクションで非白ピクセル領域をブロックに分割 |
| `mergeBlocks` | 近接ブロックを統合 |
| `trimMargins` | ブロック境界の余白を除去 |
| `makeTransparentBg` | BFS で端から白ピクセルを透過 |
| `maskText` | テキスト行領域を背景色で塗りつぶし（origData を複製して使用） |
| `sampleBgColor` | BG_MARGIN 分の外周ピクセルから中央値で背景色を推定 |
| `getTextColor` | 輝度ベースでテキスト色を推定（整数座標が必要） |
| `processOcrResults` | PaddleOCR の 4 点ポリゴン bbox を軸平行矩形に変換 |
| `extractPdfNativeText` | pdf.js の getTextContent でベクタテキストを取得 |
| `snapFontSize` | フォントサイズを定義済みリストの最近傍にスナップ |

## バグ修正履歴

- **float 座標バグ**: `processOcrResults` が float 座標を返すと `pixelsInRect` が `data[float]` を参照し `undefined` になる。`Math.floor`/`Math.ceil` で整数化。`pixelsInRect` にも `| 0` の防御的コーションを追加。
- **重複 getImageData**: ページごとに `getImageData` を複数回呼んでいた（~24 MB のコピー × N回）。レンダリング直後に 1 回だけ取得した `origData` を使い回すよう修正。
- **maskText 内部 getImageData**: `maskText` が内部で `ctx.getImageData` を呼んでいた。`origData` を `Uint8ClampedArray` で複製して加工する方式に変更。

## OCR とオフラインキャッシュについて

PaddleOCR (`@paddle-js-models/ocr`) は `esm.sh` CDN 経由で動的 import します。

**Service Worker (`sw.js`)** が以下のオリジンへのリクエストをキャッシュします:
- `esm.sh` — PaddleOCR JS モジュール・依存ライブラリ
- `bj.bcebos.com` — PaddleOCR モデルファイル（ONNX 重みなど）

初回アクセス時に全ファイルをキャッシュし、以降はオフライン環境でも動作します。  
Service Worker は `http://` / `https://` 環境でのみ有効（`file://` は非対応）。

PDF 内蔵テキストがある場合は OCR をスキップします（`use-pdftext` チェックボックス）。

## Python 版との関係

`pdf2ppt.py` はサーバーサイド版の参照実装です。  
定数・アルゴリズムを変更する際は両ファイルを同時に更新してください。
