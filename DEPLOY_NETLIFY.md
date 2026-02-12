# Netlify公開手順（GitHub連携）

最終更新: 2026-02-12

このプロジェクトは静的サイトなので、`Netlify + GitHub` でそのまま公開できます。

## 0. 事前確認

- ルート: `/Users/setohirokazu/Desktop/setoplan/screenshot-title-renamer`
- `netlify.toml` は配置済み
- 問い合わせフォーム送信先は未設定（後述）

## 1. GitHubへpush

1. GitHubで新規リポジトリを作成（例: `screenshot-title-renamer`）
2. ローカルからpush

```bash
cd /Users/setohirokazu/Desktop/setoplan/screenshot-title-renamer
git init
git add .
git commit -m "Initial release"
git branch -M main
git remote add origin https://github.com/<your-id>/screenshot-title-renamer.git
git push -u origin main
```

## 2. Netlifyでサイト作成

1. Netlifyで `Add new site -> Import an existing project`
2. GitHubを接続して上記リポジトリを選択
3. Build settings:
   - Base directory: 空欄
   - Build command: 空欄（`netlify.toml`が優先）
   - Publish directory: 空欄（`netlify.toml`が優先）
4. `Deploy site`

## 3. 公開確認

- `https://<site-name>.netlify.app/`
- 次のページが開くことを確認
  - `/`
  - `/contact.html`
  - `/guide-explainer.html`
  - `/guide-comparison.html`
  - `/guide-howto.html`

## 4. Formspree接続（必須）

1. Formspreeでフォーム作成
2. Endpoint URL（`https://formspree.io/f/xxxxxxx`）を取得
3. 次を置換
   - `/Users/setohirokazu/Desktop/setoplan/screenshot-title-renamer/contact.html`
   - `/Users/setohirokazu/Desktop/setoplan/screenshot-title-renamer/contact-en.html`
   - `data-endpoint="https://formspree.io/f/REPLACE_WITH_YOUR_FORM_ID"`
4. commit & push
5. Netlify再デプロイ後にフォーム送信テスト

## 5. 独自ドメイン（任意）

1. Netlify: `Site settings -> Domain management -> Add custom domain`
2. DNS設定（Netlifyが表示するレコードを設定）
3. HTTPS証明書が有効化されるまで待つ

## 6. AdSense準備（最小）

1. `Privacy` `Terms` `Contact` が公開されていることを確認
2. ガイド3ページ（解説/比較/使い方）を公開した状態で申請
3. 申請通過後に `ads.txt` を追加
4. 免責や注意事項ページには広告を置かない

## 7. 変更運用

- 以後は `main` へpushするだけで自動デプロイ
- 反映後にハードリロードで確認（キャッシュ対策済み）
