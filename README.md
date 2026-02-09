# sozai-remotion-caption

CSVの文章をもとに、背景画像にキャプションを合成した縦長(1080x1920)の静止画PNGを一括生成します。

## Requirements

- Node.js 18+ (推奨)
- npm

## Setup

```bash
npm ci
```

## 使い方

1. `sozai/` に背景画像を1枚置く  
   - 対応: `.png` `.jpg` `.jpeg` `.webp`
   - 複数ある場合、ファイル名の昇順で最初の1枚を使います
2. `templates/captions.csv` を編集する  
   - ヘッダー: `filename,text`
   - `text` は `\\n` を書くと改行として扱われます
3. 生成する

```bash
npm run render:csv
```

出力は `output/` に `filename` の名前でPNGが作られます（`output/` はgit管理外）。

