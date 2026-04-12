# pdf2ppt

PDF を PowerPoint (PPTX) に変換するブラウザ完結ツールです。  
ファイルはサーバーに送信されません。すべての処理はブラウザ上で実行されます。

## 機能

- **PDF 内蔵テキスト抽出** — ベクタ PDF のテキストを高精度で取得
- **PaddleOCR** — 画像 PDF・図形内テキストを OCR で認識（日本語対応）
- **画像ブロック抽出** — テキスト以外の図形・画像をブロック単位でスライドに配置
- **透過背景** — 図形ブロックの白背景を透過処理
- **PDF サイズ保持** — 元 PDF のページサイズを PPTX スライドサイズに反映

## 使い方

### オンライン（CDN モード）

`pdf2ppt.html` をブラウザで直接開くだけで動作します。  
必要なライブラリは自動的に CDN から読み込まれます。

### オフライン（ローカルライブラリモード）

1. `download_libs.bat` を実行して `lib/` フォルダにライブラリをダウンロード
2. `pdf2ppt.html` をブラウザで開く

```
pdf2ppt.html          ← メインファイル
download_libs.bat     ← ライブラリダウンロード（要: Windows 10 1803 以降）
lib/
  pdf.min.js          ← PDF.js 3.11.174
  pdf.worker.min.js   ← PDF.js Worker
  pptxgen.bundle.js   ← PptxGenJS 3.12.0
```

> **注意:** PaddleOCR モデルは初回使用時のみ CDN（esm.sh）から読み込まれます。  
> 完全オフライン環境での OCR は利用できません。

### 操作方法

1. `pdf2ppt.html` をブラウザで開く
2. PDF ファイルをドロップ、またはクリックしてファイルを選択
3. 変換オプションを選択
4. 処理完了後、「PPTX をダウンロード」ボタンをクリック

## 変換オプション

| オプション | 説明 |
|-----------|------|
| テキストを OCR で抽出 | PaddleOCR でテキストを認識し PPTX テキストボックスとして配置 |
| PDF 内蔵テキストを優先使用 | ベクタ PDF の場合は埋め込みテキストを使用（OCR より高精度） |
| 図形ブロックの背景を透過 | 図形の白背景を透明化（スライド背景が透けて見える） |

## 対応ブラウザ

- Google Chrome 90+
- Microsoft Edge 90+
- Firefox 90+

## 使用ライブラリ

| ライブラリ | バージョン | 用途 |
|-----------|-----------|------|
| [PDF.js](https://mozilla.github.io/pdf.js/) | 3.11.174 | PDF レンダリング・テキスト抽出 |
| [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) | 3.12.0 | PPTX 生成 |
| [@paddle-js-models/ocr](https://github.com/PaddlePaddle/PaddleJS) | latest | OCR（DBNet + CRNN） |

## ライセンス

MIT License — 詳細は [LICENSE](LICENSE) を参照してください。
