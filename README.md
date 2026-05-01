# 協力会社フォーム自動送信システム

## ファイル構成

```
form_sender/
├── app.py              # Webアプリ本体（Flask）
├── engine.py           # フォーム自動送信エンジン（Playwright + Claude）
├── requirements.txt
├── render.yaml         # Render デプロイ設定
└── templates/
    └── index.html      # 画面UI
```

## セットアップ手順

### 1. GitHub にアップロード

リポジトリを作成し、このフォルダを push してください。

### 2. Render.com にデプロイ

1. [Render](https://render.com) でアカウント作成
2. 「New Web Service」→ GitHub リポジトリを選択
3. 環境変数を設定（下記）
4. デプロイ完了後に発行された URL を開く

### 3. 環境変数

| 変数 | 必須 | 説明 |
|------|------|------|
| `ANTHROPIC_API_KEY` | はい | Claude API キー |
| `AUTH_USERNAME` | 推奨 | Basic 認証ユーザー名（未設定時は `admin`） |
| `AUTH_PASSWORD` | 推奨 | Basic 認証パスワード（未設定時は弱い既定値） |
| `GAS_URL` | いいえ | 追跡 URL 発行用の Google Apps Script Web アプリ URL。未設定またはプレースホルダのままの場合、資料 URL をそのまま本文に使います |
| `ANTHROPIC_MODEL` | いいえ | 既定は `claude-sonnet-4-6` |

### 4. 使い方

1. デプロイ後の URL をブラウザで開く（Basic 認証）
2. Excel（**1 行目ヘッダー、A 列＝会社名、B 列＝Web サイト URL**）をアップロード
3. **送信者・お問い合わせ本文**を画面に入力（Claude がフォーム項目へマッピング）
4. 必要なら **資料 URL** を入力（追跡用 GAS がある場合は本文に `{TRACKING_URL}` を含めるか、省略すると追跡リンク行を自動付与）
5. **テストモード**で動作確認後、本番モードで実行
6. 結果 Excel をダウンロード

## Excel フォーマット

- **先頭シート**の **A 列＝会社名、B 列＝サイト URL** を使用します（ヘッダー行は 1 行目。列名は任意で、位置で読みます）。
- `http` で始まる URL の行のみ処理します。
- 送信者情報は **画面入力**です（旧 README の複数シート専用ファイルとは異なります）。

## 注意事項

- CAPTCHA があるサイトは自動送信できません（結果の「手動対応リスト」に出ます）
- 営業禁止の記載があるページは AI 判定でスキップします
- 必ずテストモードで確認してから本番実行してください
- 本番は `gunicorn` で起動します（ローカルでは `python app.py` のままで可）
