# セキュリティガイド

このドキュメントでは、X自動投稿GASシステムを安全に使用するためのセキュリティ対策について説明します。

## 重要な注意事項

### 🚨 絶対に公開してはいけない情報

以下の情報は**絶対にGitHubや公開リポジトリにコミットしないでください**：

1. **X API認証情報**
   - `X_API_KEY`
   - `X_API_SECRET`
   - `X_ACCESS_TOKEN`
   - `X_ACCESS_TOKEN_SECRET`

2. **設定ファイル**
   - `.env` ファイル
   - `config.json`
   - `credentials.json`
   - `script-properties.json`

3. **スプレッドシートURL**
   - 実際に使用しているスプレッドシートのURL
   - スプレッドシートID

## セキュリティ対策

### 1. .gitignoreの使用

`.gitignore` ファイルに機密情報を含むファイルを追加してください。すでに以下が設定されています：

```
.env
.env.local
config.json
credentials.json
secrets.json
script-properties.json
```

### 2. 環境変数の管理

機密情報は環境変数またはGASのスクリプトプロパティで管理してください：

#### ローカル開発時（将来的な拡張用）
1. `.env.example` を `.env` にコピー
2. `.env` に実際の認証情報を記載
3. `.env` は絶対にコミットしない

#### GAS本番環境
1. GASエディタで「プロジェクトの設定」を開く
2. 「スクリプト プロパティ」で認証情報を設定
3. スクリプトプロパティは自動的に暗号化されて保存される

### 3. APIキーの権限設定

X Developer Portalで、APIキーの権限を必要最小限に設定してください：

- **Read and Write**: 投稿のみの場合
- **Read, Write, and Direct Messages**: DM機能が不要な場合は設定しない

### 4. トークンのローテーション

定期的に（3〜6ヶ月ごと）APIトークンを再生成することを推奨します：

1. X Developer Portalでトークンを再生成
2. GASのスクリプトプロパティを更新
3. 古いトークンは無効化

### 5. アクセス制御

#### スプレッドシート
- 共有設定を「制限付き」に設定
- 必要最小限のユーザーのみに共有
- 「編集者」権限は信頼できるユーザーのみ

#### GASプロジェクト
- スクリプトの共有は慎重に行う
- デプロイ時は「自分のみ」または「特定ユーザー」に制限

### 6. ログの取り扱い

ログに機密情報が含まれないよう注意してください：

```javascript
// ❌ 悪い例
console.log(`API Key: ${apiKey}`);

// ✅ 良い例
console.log('API認証成功');
```

本スクリプトでは、APIキーやトークンはログ出力されないように設計されています。

### 7. コミット前のチェック

GitHubにpushする前に、以下を確認してください：

```bash
# .env ファイルが存在しないことを確認
ls -la | grep .env

# gitで追跡されているファイルを確認
git status

# .gitignoreが正しく機能しているか確認
git check-ignore .env
```

### 8. 誤ってコミットした場合の対処

万が一、機密情報をコミットしてしまった場合：

1. **すぐにトークンを無効化**
   - X Developer Portalでトークンを再生成

2. **Gitの履歴から削除**
   ```bash
   # 特定ファイルを履歴から完全削除
   git filter-branch --force --index-filter \
     "git rm --cached --ignore-unmatch .env" \
     --prune-empty --tag-name-filter cat -- --all

   # 強制プッシュ
   git push origin --force --all
   ```

3. **GitHubでリポジトリを削除して再作成**（最も確実）

## 推奨されるワークフロー

### 初回セットアップ

1. このリポジトリをクローン
2. `.env.example` を `.env` にコピー
3. `.env` に実際の認証情報を記載（絶対にコミットしない）
4. GASのスクリプトプロパティに認証情報を設定
5. `.gitignore` が正しく機能していることを確認

### 定期的なメンテナンス

- **月1回**: ログを確認して不審なアクティビティがないかチェック
- **3ヶ月ごと**: APIトークンのローテーション
- **6ヶ月ごと**: アクセス権限の見直し

## 脆弱性の報告

セキュリティ上の問題を発見した場合は、公開のIssueではなく、リポジトリ管理者に直接連絡してください。

## 参考リンク

- [X Developer Portal](https://developer.twitter.com/)
- [Google Apps Script ベストプラクティス](https://developers.google.com/apps-script/guides/security)
- [OAuth 1.0a 仕様](https://oauth.net/core/1.0a/)

---

**最終更新**: 2025年1月
