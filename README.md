# 🏠 みえる化家計簿

LINE Bot連携・Gemini AI分析・Googleスプレッドシートをバックエンドにした、スマートフォン最適化の家計簿Webアプリです。

---

## 📋 主な機能

| 機能 | 説明 |
|------|------|
| **LINE Bot連携** | LINEから「ランチ 1200」と送るだけで自動記録 |
| **ダッシュボード** | 月次サマリー、円グラフ、サンキー図、年間収支レポート |
| **AI家計分析** | Gemini AIによる客観的な家計アドバイスと週次・月次レポート |
| **固定費自動記録** | 毎月指定日に固定費を自動でスプレッドシートへ記録 |
| **予算アラート** | 予算の80%・100%超過時にLINEでPush通知 |
| **Gmail連携** | 三井住友カード・PayPayカードの利用通知メールを自動取り込み |
| **設定管理** | 月間予算・カテゴリ・固定費・口座をダッシュボードから設定 |

---

## 🏗️ アーキテクチャ

```
Google Apps Script (サーバーサイド)
├── config.js         定数・ScriptProperties（APIキー等）
├── main.js           doGet / doPost エントリポイント
├── line_bot.js       LINE Bot Webhook処理・メッセージ送信
├── data_access.js    スプレッドシート読み書き・バリデーション
├── settings.js       設定データ取得・保存
├── dashboard_api.js  ダッシュボードAPI（月次・年次・サンキー）
├── ai_analysis.js    Gemini AI分析・アドバイス生成
├── triggers.js       タイムドリブントリガー（固定費・アラート・レポート）
├── gmail_integration.js  カードメール解析・自動記録
└── utils.js          共通ユーティリティ・バリデーション

フロントエンド (HtmlService テンプレート)
├── index.html        テンプレート（構造・インクルード指示）
├── styles.html       CSS（Kawaii & Clean パステルテーマ）
├── helpers.html      ユーティリティJS（animateNumber, escapeHtml等）
├── charts.html       チャート描画（D3.js: 円グラフ、サンキー図）
├── records.html      取引一覧表示
├── modals.html       設定・入力モーダル
└── scripts.html      メインロジック（API呼び出し、loadData等）
```

バックエンド: **Google Apps Script**（無料サーバーレス）
データベース: **Google スプレッドシート**
フロントエンド: **HTML/CSS/JS** + D3.js + Marked.js
AI: **Gemini API**（gemini-2.5-flash）
通知: **LINE Messaging API**

---

## ⚙️ セットアップ

### 1. スクリプトプロパティの設定

GASエディタの「プロジェクトの設定」→「スクリプトプロパティ」に以下を設定：

| キー | 値 |
|------|-----|
| `LINE_ACCESS_TOKEN` | LINE Messaging APIのチャネルアクセストークン |
| `LINE_CHANNEL_SECRET` | LINE Messaging APIのチャネルシークレット |
| `SPREADSHEET_ID` | GoogleスプレッドシートのID |
| `GEMINI_API_KEY` | Google AI StudioのAPIキー |

### 2. スプレッドシートの準備

GASエディタで `createDatabase()` を実行すると、必要なシートが自動作成されます。

### 3. デプロイ

```bash
# Dockerコンテナ経由でデプロイ
cd env && docker-compose exec -w /workspace/kakeibo-dashboard -T agent-env clasp push -f
```

GASエディタで「デプロイ」→「新しいデプロイ」→ 種類: **ウェブアプリ**

- 実行ユーザー: **自分**
- アクセス権: **全員**

### 4. LINE Webhook設定

デプロイで発行されたURLをLINE Developers ConsoleのWebhook URLに設定してください。

### 5. トリガー設定

GASエディタで以下を手動実行してタイムドリブントリガーを登録：

```javascript
setupAITriggers(); // 週次・月次レポートトリガー
```

固定費の自動記録は `autoRecordFixedExpenses`、予算アラートは `checkBudgetAndAlert` を毎日実行するトリガーをGASエディタから手動設定してください。

---

## 📁 ファイル構成

```
kakeibo-dashboard/
├── .clasp.json           clasp設定（scriptId, filePushOrder）
├── appsscript.json       GASマニフェスト（タイムゾーン, oauthスコープ）
├── README.md             このファイル
├── ANTIGRAVITY.md        開発メモ・機能詳細
├── design_system.md      CSSデザインシステム
├── spreadsheet_schema.md スプレッドシート構造
├── setup_guide_phase1.md セットアップガイド
│
├── [GAS バックエンド]
│   ├── config.js
│   ├── main.js
│   ├── line_bot.js
│   ├── data_access.js
│   ├── settings.js
│   ├── dashboard_api.js
│   ├── ai_analysis.js
│   ├── triggers.js
│   ├── gmail_integration.js
│   └── utils.js
│
└── [フロントエンド]
    ├── index.html
    ├── styles.html
    ├── helpers.html
    ├── charts.html
    ├── records.html
    ├── modals.html
    └── scripts.html
```

---

## 🔐 セキュリティ

- APIキー・シークレットはすべて**ScriptProperties**で管理（コードにハードコードしない）
- LINE Webhookのリクエストは**HMAC-SHA256署名**で検証
- フロントエンドのユーザー入力は**XSSエスケープ**処理済み
- GAS_API_URLはテンプレート内で**動的生成**（ハードコード不使用）

---

## 💡 LINE Bot コマンド例

| 入力例 | 動作 |
|--------|------|
| `ランチ 1200` | 食費 1,200円を記録 |
| `ランチ 1200 カフェ` | 食費 1,200円（カフェ）を記録 |
| `給与 300000 収入` | 収入 300,000円を記録 |
| `今月` | 今月のサマリーを返信 |
| `分析` | AI家計分析レポートを返信 |
