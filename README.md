# X Auto Post System

Google Sheetsで管理した投稿内容を、指定した日時に自動でX（旧Twitter）へ投稿するGoogle Apps Scriptシステムです。サーバー不要・完全無料で動作します。

## 機能

- **予約投稿** — スプレッドシートに日時を入力するだけで自動投稿
- **画像添付** — セル内画像・Google Drive URLの両方に対応（最大4枚）
- **スレッド投稿** — グループIDで複数ツイートをスレッド化
- **note連携** — 投稿後にnote記事URLをリプライとして自動送信
- **エンゲージメント計測** — いいね・リポスト・返信・引用・インプレッション数を自動取得
- **文字数カウント** — Xの仕様（全角140文字、URL短縮）に準拠

## 構成ファイル

```
src/
├── post_x.js          # メイン処理（投稿・エンゲージメント取得）
├── setup_template.js  # スプレッドシート初期セットアップ用
└── appsscript.json    # Apps Script マニフェスト
```

## セットアップ

### 1. スプレッドシートの準備

`setup_template.js` の内容を新しいスプレッドシートのApps Scriptに貼り付けて `setupTemplate()` を実行すると、シートが自動作成されます。

手動で作成する場合は `sheet` という名前のシートを用意し、以下の列構成にしてください。

| 列 | 項目名 | 説明 |
|---|---|---|
| A | 投稿本文 | ツイートの内容（必須） |
| B | 文字数 | 自動計算（X仕様準拠） |
| C | 画像 | セル内画像 or Google Drive URL（最大4枚） |
| D | noteリンク | 投稿後にリプライとして投稿されるURL（任意） |
| E | 日にち | 投稿予定日（シリアル値） |
| F | 時 | 投稿予定時刻の「時」（0〜23） |
| G | 分 | 投稿予定時刻の「分」（0〜59） |
| H | 投稿URL | 投稿完了後に自動記録されるツイートURL |
| I | スレッドグループID | 同じIDを持つ行を連続投稿（任意） |
| J | ステータス | pending / posting / posted / failed |
| K | いいね数 | エンゲージメント指標（自動更新） |
| L | リポスト数 | エンゲージメント指標（自動更新） |
| M | 返信数 | エンゲージメント指標（自動更新） |
| N | 引用数 | エンゲージメント指標（自動更新） |
| O | インプレッション数 | 有料プラン専用（自動更新） |
| P | ブックマーク数 | 有料プラン専用（自動更新） |
| Q | 最終更新日時 | エンゲージメント更新日時 |

### 2. GASプロジェクトへのコード配置

[Google Apps Script](https://script.google.com) でスプレッドシートのコンテナバインドプロジェクトを開き、`src/post_x.js` の内容を貼り付けてください。

claspを使う場合:

```bash
npm install -g @google/clasp
clasp login
clasp push
```

### 3. スクリプトプロパティの設定

Apps Scriptエディタの「プロジェクトの設定 → スクリプトプロパティ」に以下を追加します。

| プロパティ名 | 取得元 |
|---|---|
| `X_API_KEY` | X Developer Portal → アプリの Consumer Key |
| `X_API_SECRET` | X Developer Portal → アプリの Consumer Secret |
| `X_ACCESS_TOKEN` | X Developer Portal → Access Token |
| `X_ACCESS_TOKEN_SECRET` | X Developer Portal → Access Token Secret |

> APIは **Read and Write** 権限が必要です。

### 4. トリガーの設定

Apps Scriptエディタの「トリガー」から以下を追加します。

| 関数名 | 推奨頻度 |
|---|---|
| `postNextScheduledToX` | 1分〜1時間ごと（投稿頻度に合わせて） |
| `dailyEngagementUpdate` | 1日1回（例: 毎朝9時） |

## 画像の投稿

セルへの画像挿入は以下のどちらでも対応しています。

- **セル内に挿入**: 「挿入 → 画像 → セル内に画像を挿入」
- **セル上に配置**: 「挿入 → 画像 → セル上に画像を配置」

## セキュリティ

APIキーなどの認証情報はGASのスクリプトプロパティで管理し、コードには直接書かないでください。詳細は [SECURITY.md](SECURITY.md) を参照してください。

## ライセンス

MIT License
