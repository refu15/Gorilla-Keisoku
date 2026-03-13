# 🚀 AI活用度測定 - セットアップガイド

配布されたZIPファイルからChrome拡張機能をインストール・設定するためのガイドです。

---

## 📋 ロードマップ（全体の流れ）

```
① ZIPファイルを解凍する
    ↓
② Google Cloud Projectを作成・設定する
    ↓
③ Chrome拡張機能をインストールする
    ↓
④ 拡張機能ID を Google Cloud に登録する
    ↓
⑤ Googleスプレッドシートを準備する
    ↓
⑥ 拡張機能の初期設定を行う
    ↓
⑦（任意）Gemini API / Notion連携を設定する
    ↓
✅ 使い始める！
```

---

## ① ZIPファイルを解凍する

1. 配布された `keisokuGoogle.zip` をダウンロードします
2. ZIPファイルを右クリック → **「すべて展開」** を選択
3. 展開先フォルダを選んで **「展開」** をクリック
4. 展開されたフォルダ内に以下の構成が確認できればOKです：

```
keisokuGoogle/
├── manifest.json          ← 拡張機能の設定ファイル（後で編集）
├── background.js
├── sidepanel/
│   ├── index.html
│   ├── styles.css
│   └── app.js
├── src/
│   ├── auth/
│   ├── api/
│   └── utils/
└── icons/
```

> [!IMPORTANT]
> `.git` フォルダや `website` フォルダ、`vercel.json` は配布ZIPに含めないでください（開発用ファイルのため）。

---

## ② Google Cloud Projectの設定

### 2-1. プロジェクトを作成する

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 画面上部のプロジェクト選択 → **「新しいプロジェクト」** をクリック
3. プロジェクト名を入力（例：`AI活用度測定`）して **「作成」**

### 2-2. APIを有効化する

1. 左メニューの **「APIとサービス」** → **「ライブラリ」**
2. 以下の3つのAPIを検索して **「有効にする」** をクリック：

| API名 | 用途 |
|--------|------|
| **Google Calendar API** | カレンダーの予定を取得 |
| **Google Sheets API** | スプレッドシートに日報を保存 |

### 2-3. OAuth同意画面を設定する

1. **「APIとサービス」** → **「OAuth同意画面」**
2. ユーザータイプ：**「内部」**（組織内のみ）または **「外部」** を選択
3. アプリ名、サポートメール等を入力
4. スコープの追加：
   - `Google Calendar API` → `.../auth/calendar.readonly`
   - `Google Sheets API` → `.../auth/spreadsheets`
   - `UserInfo` → `email`, `profile`
5. 設定を保存

### 2-4. OAuthクライアントIDを作成する

1. **「APIとサービス」** → **「認証情報」** → **「認証情報を作成」**
2. **「OAuthクライアントID」** を選択
3. アプリケーションの種類：**「Chrome拡張機能」**
4. 拡張機能のID：（③のインストール後に入力するので、ここでは一旦スキップしてもOK）
5. **「作成」** をクリック
6. 表示される **クライアントID** をコピーしておく

> [!TIP]
> クライアントIDは `XXXXXXXXX.apps.googleusercontent.com` の形式です。

---

## ③ Chrome拡張機能をインストールする

### 3-1. manifest.jsonを編集する

1. 解凍したフォルダ内の `manifest.json` をテキストエディタで開く
2. `client_id` を、②で取得した **自分のクライアントID** に書き換える：

```json
"oauth2": {
  "client_id": "ここに自分のクライアントIDを貼り付け.apps.googleusercontent.com",
  "scopes": [
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/spreadsheets"
  ]
}
```

3. ファイルを保存

### 3-2. Chromeに拡張機能を読み込む

1. Chromeブラウザで `chrome://extensions` を開く
2. 右上の **「デベロッパーモード」** をONにする
3. **「パッケージ化されていない拡張機能を読み込む」** をクリック
4. 解凍した `keisokuGoogle` フォルダを選択
5. 拡張機能が一覧に表示されればインストール完了！

> [!NOTE]
> インストール後に表示される **拡張機能ID**（`abcdefg...` のような文字列）をメモしてください。次のステップで使います。

---

## ④ 拡張機能IDをGoogle Cloudに登録する

1. [Google Cloud Console](https://console.cloud.google.com/) に戻る
2. **「APIとサービス」** → **「認証情報」**
3. 作成したOAuthクライアントIDをクリックして編集
4. **「アプリケーションID」** 欄に、③でメモした **拡張機能ID** を入力
5. **「保存」** をクリック

> [!CAUTION]
> この手順を忘れると、Googleアカウントでのログインが失敗します。必ず設定してください。

---

## ⑤ Googleスプレッドシートを準備する

1. [Google スプレッドシート](https://sheets.google.com) で新規スプレッドシートを作成
2. スプレッドシートの**URL**からIDを取得：

```
https://docs.google.com/spreadsheets/d/【ここがスプレッドシートID】/edit
```

| URL の例 | IDの部分 |
|----------|---------|
| `https://docs.google.com/spreadsheets/d/1aBcDeFg.../edit` | `1aBcDeFg...` |

3. このIDをメモしておく

> [!TIP]
> シートのヘッダー行は自動生成されるので、空のスプレッドシートのままでOKです。

---

## ⑥ 拡張機能の初期設定

1. Chromeのツールバーにある拡張機能アイコン（パズルピース）をクリック
2. **「AI活用度測定」** をクリックしてサイドパネルを開く
3. **Googleアカウントでログイン** する
4. 設定画面（⚙️アイコン）を開く
5. **スプレッドシートID** を入力して保存

---

## ⑦（任意）追加設定

### Gemini API連携（AI分析機能）

AI分析機能を使うには、Gemini APIキーが必要です。

1. [Google AI Studio](https://aistudio.google.com/apikey) にアクセス
2. **「APIキーを作成」** をクリック
3. 取得したAPIキーを拡張機能の設定画面で入力

### Notion連携（日報をNotionに保存）

1. [Notion Integrations](https://www.notion.so/my-integrations) でインテグレーションを作成
2. **Internal Integration Token** をコピー
3. 日報を保存したいNotionデータベースにインテグレーションを接続
4. データベースIDを取得（NotionページURLから）
5. 拡張機能の設定画面で **Notionトークン** と **データベースID** を入力

---

## ✅ 使い方（基本操作）

```
1. Chromeツールバーの拡張機能アイコンをクリック
   ↓
2. サイドパネルが開き、本日の予定一覧が表示
   ↓
3. 各予定に対して：
   • AI活用余地 → No / Yes-使用中 / Yes-余地あり を選択
   • 活用率 → スライダーで 0〜100% を設定
   • メモ → 使用したAIツール名など（任意）
   ↓
4. 「日報プレビュー」で確認
   ↓
5. 「送信」でスプレッドシート（/Notion）に保存！
```

---

## 🔧 トラブルシューティング

| 症状 | 対処法 |
|------|--------|
| ログインできない | ④の拡張機能ID登録を確認 |
| 予定が表示されない | Google Calendar APIの有効化を確認 |
| スプレッドシートに保存できない | スプレッドシートIDの入力ミスを確認 / Sheets APIの有効化を確認 |
| AI分析が使えない | Gemini APIキーの入力を確認 |
| 拡張機能が読み込めない | `manifest.json` の編集でJSON構文が壊れていないか確認 |

---

## 📦 配布用ZIPの作り方（管理者向け）

配布用ZIPには以下のファイル・フォルダのみ含めてください：

```
含めるもの：
✅ manifest.json
✅ background.js
✅ sidepanel/（フォルダごと）
✅ src/（フォルダごと）
✅ icons/（フォルダごと）
✅ scripts/（フォルダごと）
✅ README.md
✅ SETUP_GUIDE.md（このファイル）

除外するもの：
❌ .git/
❌ website/
❌ vercel.json
```

### PowerShellでZIPを作成するコマンド：

```powershell
Compress-Archive -Path manifest.json, background.js, sidepanel, src, icons, scripts, README.md, SETUP_GUIDE.md -DestinationPath keisokuGoogle.zip
```

---

> 質問やトラブルがあれば管理者にお問い合わせください。
