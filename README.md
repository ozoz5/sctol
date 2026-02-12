# Apple Receipt Batch Renamer (Web)

Appleの領収書スクリーンショットを一括で命名整理するWebツールです。

## 主な機能

- Apple領収書プリセット: `APPLE-日付-合計-サービス名`
- サービス名の3層対応: レイアウト抽出 + 辞書正規化 + ローカル学習
- 重複候補の自動検出（同じ `日付+合計+サービス名`）
- ZIP出力時に月別フォルダへ自動振り分け（`YYYY-MM/`）
- `日付 / 合計 / サービス名` を個別編集して即再計算

## サービス名対応の更新方法

- `/Users/setohirokazu/Desktop/setoplan/screenshot-title-renamer/service-catalog.js` を編集
- `aliases`: OCRで拾う別名パターン（正規表現）を追加
- `display`: 正式表示名を追加
- `main.js` はこの定義を自動で読み込みます

## 使い方

1. フォルダでターミナルを開いて `python3 -m http.server 8787` を実行
2. ブラウザで `http://localhost:8787` を開く
3. 画像をドロップ（またはファイル選択）
4. 必要なら `日付 / 合計 / サービス名` を修正
5. `まとめて保存 (ZIP)` を押す

## Netlify公開（推奨）

- 設定ファイル: `/Users/setohirokazu/Desktop/setoplan/screenshot-title-renamer/netlify.toml`
- 手順書: `/Users/setohirokazu/Desktop/setoplan/screenshot-title-renamer/DEPLOY_NETLIFY.md`
- 先に `Formspree` の endpoint を設定してから本番公開してください

## 注意

- ブラウザ仕様上、元ファイルを直接リネームはできません。新しい名前で保存する方式です。
- OCRは `Tesseract.js` (CDN) を使うため、初回読み込み時にネット接続が必要です。
- 本サービスはApple社とは独立して運営しており、Apple社の提供・認定・保証を受けたものではありません。
