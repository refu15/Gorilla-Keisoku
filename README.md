# AI活用度測定 - 日報作成アシスタント

Googleカレンダーから日報を自動生成し、各予定のAI活用度を記録するChrome拡張機能です。

## 機能

- 📅 **Googleカレンダー連携**: 当日の予定を自動取得
- 🤖 **AI活用度記録**: 各予定に対してAI活用有無と活用率を記録
- 📝 **日報自動生成**: 入力データからワンクリックで日報を生成
- 📊 **スプレッドシート保存**: Googleスプレッドシートへ自動保存
- 🔄 **オフライン対応**: ネットワーク切断時もデータをキューイング

## セットアップ

### 1. Google Cloud Projectの設定

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 新しいプロジェクトを作成（または既存のプロジェクトを選択）
3. 以下のAPIを有効化:
   - Google Calendar API
   - Google Sheets API
4. OAuth 2.0クライアントIDを作成:
   - 「認証情報」→「認証情報を作成」→「OAuthクライアントID」
   - アプリケーションの種類: **Chrome拡張機能**
   - 拡張機能ID（後で設定）

### 2. manifest.jsonの設定

`manifest.json` 内のクライアントIDを更新:

```json
"oauth2": {
  "client_id": "YOUR_CLIENT_ID.apps.googleusercontent.com",
  ...
}
```

### 3. 拡張機能のインストール

1. Chromeで `chrome://extensions` を開く
2. 「デベロッパーモード」を有効化
3. 「パッケージ化されていない拡張機能を読み込む」をクリック
4. このプロジェクトのフォルダを選択

### 4. スプレッドシートの設定

1. 日報保存用のGoogleスプレッドシートを作成
2. URLから`スプレッドシートID`を取得（`/d/`と`/edit`の間の文字列）
3. 拡張機能の設定画面でIDを入力

## 使い方

1. Chrome拡張機能のアイコンをクリック
2. Googleアカウントでログイン
3. 当日の予定一覧が表示される
4. 各予定に対して:
   - **AI活用余地**: No / Yes-使用中 / Yes-余地あり を選択
   - **活用率**: スライダーで0-100%を設定
   - **メモ**: 使用したツールや詳細を記入（任意）
5. 「日報プレビュー」で内容を確認
6. 「送信」ボタンでスプレッドシートに保存

## ファイル構成

```
keisokuGoogle/
├── manifest.json          # 拡張機能設定
├── background.js          # サービスワーカー
├── sidepanel/
│   ├── index.html         # サイドパネルUI
│   ├── styles.css         # スタイル
│   └── app.js             # メインロジック
├── src/
│   ├── auth/oauth.js      # OAuth認証
│   ├── api/
│   │   ├── calendar.js    # Calendar API
│   │   └── sheets.js      # Sheets API
│   └── utils/
│       ├── ai-estimator.js    # AI推定
│       ├── offline-queue.js   # オフラインキュー
│       └── storage.js         # ストレージ管理
└── icons/                 # アイコン
```

## データ構造

スプレッドシートに保存されるデータ:

| 列 | 内容 |
|---|---|
| A | ユーザーメール |
| B | 日付 |
| C | 予定タイトル |
| D | 開始時刻 |
| E | 終了時刻 |
| F | カレンダー名 |
| G | AI活用フラグ |
| H | 活用率(%) |
| I | メモ |
| J | 送信日時 |

## 開発

### 必要な権限

- `identity`: OAuth認証
- `storage`: データ保存
- `sidePanel`: サイドパネル表示
- `activeTab`: 現在のタブへのアクセス

### OAuthスコープ

- `calendar.readonly`: カレンダー読み取り
- `userinfo.email`: ユーザーメール取得
- `userinfo.profile`: ユーザープロフィール取得
- `spreadsheets`: スプレッドシート読み書き

## ライセンス

社内利用限定
