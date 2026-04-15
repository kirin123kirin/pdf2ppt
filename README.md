# pdf2ppt

PDF・画像を PowerPoint (PPTX) に変換するローカル Web アプリです。  
Flask サーバーがローカルで動作し、ブラウザから操作できます。ファイルは外部に送信されません。

## 機能

- **PDF 変換** — PDF の各ページをスライドに変換
- **画像変換** — JPG / PNG / BMP / WebP / TIFF / SVG に対応
- **クリップボード貼り付け** — `Ctrl+V` でクリップボード画像を直接変換
- **surya OCR** — テキストを高精度で認識しテキストボックスとして配置（日本語対応）
- **画像ブロック抽出** — テキスト以外の図形・画像をブロック単位でスライドに配置
- **PDF サイズ保持** — 元 PDF のページサイズを PPTX スライドサイズに反映

## セットアップ

### 初回のみ

```
setup_py.bat を実行
```

- Python 3.11 Embeddable + 必要ライブラリ + surya OCR モデルをダウンロードします
- 合計ダウンロードサイズ: 約 2〜3 GB
- インターネット接続が必要です（初回のみ）

## 使い方

### Web UI（推奨）

```
start.bat を実行
```

ブラウザが自動で開きます。PDF または画像をドロップして変換してください。

1. `start.bat` を実行
2. ブラウザでファイルをドロップ / クリックして選択 / `Ctrl+V` でペースト
3. 変換完了後「PPTX をダウンロード」ボタンをクリック

### CLI

```
pdf2ppt.bat input.pdf
pdf2ppt.bat input.jpg
```

## ファイル構成

```
pdf2ppt.py        変換ロジック（surya OCR・ブロック検出・PPTX 生成）
server.py         Flask Web サーバー
setup_py.bat      Python 環境・surya モデルのセットアップ（初回のみ）
start.bat         Web UI 起動
pdf2ppt.bat       CLI 実行
```

## 対応入力形式

| 形式 | 備考 |
|------|------|
| PDF | 複数ページ対応 |
| JPG / PNG / BMP / WebP / TIFF | 一般的な画像形式 |
| SVG | `cairosvg` または `svglib` が必要 |
| クリップボード画像 | `Ctrl+V` で貼り付け |

## 使用ライブラリ

| ライブラリ | 用途 |
|-----------|------|
| [surya-ocr](https://github.com/VikParuchuri/surya) | テキスト検出・認識 |
| [PyMuPDF](https://pymupdf.readthedocs.io/) | PDF レンダリング |
| [python-pptx](https://python-pptx.readthedocs.io/) | PPTX 生成 |
| [Pillow](https://pillow.readthedocs.io/) | 画像処理 |
| [Flask](https://flask.palletsprojects.com/) | Web サーバー |

## ライセンス

MIT License — 詳細は [LICENSE](LICENSE) を参照してください。
