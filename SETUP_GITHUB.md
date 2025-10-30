# GitHubセットアップガイド

このプロジェクトをGitHubに安全にアップロードするための手順です。

## 📋 事前準備

### 1. Gitのインストール確認

```bash
git --version
```

インストールされていない場合は [Git公式サイト](https://git-scm.com/) からダウンロードしてください。

### 2. GitHubアカウントの作成

[GitHub](https://github.com/) でアカウントを作成してください（既にある場合はスキップ）。

## 🚀 アップロード手順

### ステップ1: ローカルリポジトリの初期化

```bash
# プロジェクトディレクトリに移動
cd C:\Users\skyeu\codex\GAS

# Gitリポジトリを初期化
git init

# .gitignoreが正しく設定されているか確認
cat .gitignore
```

### ステップ2: 機密情報の確認

**重要**: 以下のファイルが存在しないことを確認してください：

```bash
# 機密情報を含むファイルがないか確認
ls -la | grep -E "\.env$|credentials|secrets|properties\.json"
```

もし `.env` や `credentials.json` などが存在する場合は、削除するか `.gitignore` に追加されていることを確認してください。

### ステップ3: ファイルをステージング

```bash
# すべてのファイルを追加（.gitignoreで除外されたファイルは自動的にスキップ）
git add .

# 追加されたファイルを確認
git status
```

**確認ポイント**:
- ✅ `Code.gs` が追加されている
- ✅ `README.md` が追加されている
- ✅ `.gitignore` が追加されている
- ✅ `.env.example` が追加されている
- ❌ `.env` は**追加されていない**
- ❌ `credentials.json` は**追加されていない**

### ステップ4: 初回コミット

```bash
git commit -m "Initial commit: X自動投稿GASシステム"
```

### ステップ5: GitHubでリポジトリを作成

1. [GitHub](https://github.com/new) にアクセス
2. リポジトリ名を入力（例: `x-auto-post-gas`）
3. 説明を入力（任意）
4. **Private** を選択（推奨）
5. 「Create repository」をクリック

### ステップ6: リモートリポジトリを追加

GitHubで表示される指示に従って、リモートリポジトリを追加します：

```bash
# リモートリポジトリを追加（URLは自分のリポジトリURLに置き換え）
git remote add origin https://github.com/YOUR_USERNAME/x-auto-post-gas.git

# ブランチ名をmainに変更（必要に応じて）
git branch -M main

# GitHubにプッシュ
git push -u origin main
```

### ステップ7: 認証設定

初回プッシュ時に認証が求められます：

#### Personal Access Token（推奨）

1. [GitHub Settings > Developer settings > Personal access tokens](https://github.com/settings/tokens) にアクセス
2. 「Generate new token (classic)」をクリック
3. スコープで `repo` を選択
4. トークンをコピー
5. Gitのプッシュ時にパスワードの代わりに使用

#### SSH Key（上級者向け）

```bash
# SSH鍵を生成
ssh-keygen -t ed25519 -C "your_email@example.com"

# 公開鍵をGitHubに追加
cat ~/.ssh/id_ed25519.pub
# 表示された内容をGitHub Settings > SSH and GPG keys に追加
```

## 🔒 セキュリティチェック

プッシュ前に必ず以下を確認してください：

### チェックリスト

- [ ] `.gitignore` が正しく設定されている
- [ ] `.env` ファイルがGitで追跡されていない
- [ ] API認証情報がコードに直接書かれていない
- [ ] スプレッドシートURLが含まれていない
- [ ] `git status` で不要なファイルが含まれていない
- [ ] リポジトリをPrivateに設定した

### 追跡されているファイルを確認

```bash
# 追跡されているすべてのファイルを表示
git ls-files
```

### .envが追跡されていないか確認

```bash
# .envが無視されているか確認（"!!"と表示されればOK）
git check-ignore -v .env
```

## 🔄 更新の反映

コードを変更した後、GitHubに反映する方法：

```bash
# 変更されたファイルを確認
git status

# 変更をステージング
git add .

# コミット
git commit -m "機能追加: スレッド投稿機能を実装"

# GitHubにプッシュ
git push
```

## ⚠️ トラブルシューティング

### 誤って.envをコミットしてしまった

```bash
# ファイルをGit管理から削除（実ファイルは残す）
git rm --cached .env

# .gitignoreに追加されているか確認
echo ".env" >> .gitignore

# コミット
git commit -m "Remove .env from tracking"

# プッシュ
git push
```

**注意**: 一度でもコミットしてプッシュした場合は、APIトークンを再生成してください。

### リモートとの競合

```bash
# リモートの変更を取得
git pull origin main

# 競合を解決してから再度プッシュ
git push
```

## 📚 参考リンク

- [Git公式ドキュメント](https://git-scm.com/doc)
- [GitHub Docs](https://docs.github.com/)
- [.gitignore テンプレート](https://github.com/github/gitignore)

---

**次のステップ**: [SECURITY.md](SECURITY.md) でセキュリティ対策を確認してください。
