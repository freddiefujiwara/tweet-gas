# tweet-gas

X（@freddiefujiwara）へ投稿するための Google Apps Script バックエンド。

- 投稿：`POST https://api.twitter.com/2/tweets`（OAuth 1.0a ユーザーコンテキスト、HMAC-SHA1 署名を自前実装）
- 認証：Web アプリ（doPost）は Google ID トークンを検証し、メールアドレスが
  `zegenvs@gmail.com` の場合だけ投稿を受け付ける
- 確認：投稿後にツイートIDで再取得し、本文を照合して返す
- フロントエンドは別リポジトリ `tweet`

## Script Properties（キーはコードに書かない）

Apps Script エディタの「プロジェクトの設定」で以下を設定する。

| キー | 値 |
|---|---|
| `X_CONSUMER_KEY` | X デベロッパーコンソールの Consumer Key |
| `X_CONSUMER_SECRET` | 同 Consumer Secret |
| `X_ACCESS_TOKEN` | 同 Access Token（Read and write） |
| `X_ACCESS_TOKEN_SECRET` | 同 Access Token Secret |
| `GOOGLE_CLIENT_ID` | （任意）IDトークンの aud を固定する場合に設定 |

## 初回セットアップ

1. https://script.google.com で新規プロジェクトを作成し、その Script ID を `.clasp.json` の `scriptId` に設定する
2. 上記の Script Properties を設定する
3. GitHub リポジトリの Secrets に `CLASPRC_JSON`（clasp の認証情報）を設定する
4. `.github/workflows/node.yaml` を追加する（このファイルは GitHub App の権限外のため、Web UI から手動で追加する。内容は time-table-gas と同一）
5. main へ push すると CI がテスト→ビルド→ `clasp push` →バージョン作成→デプロイまで行う
   - 既存デプロイを更新する場合は Secret `DEPLOYMENT_ID` も設定する
6. デプロイされた Web アプリの `/exec` URL を、フロントエンド `tweet` の `config.js` に設定する

## 動作確認

- エディタで `postTestTweet()` を実行すると、時刻入りのテスト投稿が行われ、ログにツイートURLが出る
- Web アプリの `/exec` へ GET すると `{"ok":true,"service":"tweet-gas",...}` が返る

## 開発

```bash
npm install
npm test        # vitest（署名は RFC 5849 の公式テストベクタで検証）
npm run build   # src/Code.js -> dist/Code.gs
```
