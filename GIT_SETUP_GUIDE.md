# Git初期化 & GitHub プッシュ手順書

ぎゆう様へ：以下の手順を実行してください。

---

## 【Step 1】ターミナル/PowerShell を開く

### Windows 10/11
```
1. Windowsキー + R
2. "cmd" または "powershell" と入力
3. Enter キーを押す
4. 表示されたウィンドウで以下のコマンドを実行
```

---

## 【Step 2】リポジトリディレクトリに移動

```bash
cd C:\Users\user\Desktop\Excel-SpecificationChecker
```

※ パスはぎゆう様の環境に合わせて修正してください。

---

## 【Step 3】Git初期化

```bash
git init
```

**出力例**：
```
Initialized empty Git repository in C:\Users\user\Desktop\Excel-SpecificationChecker\.git/
```

---

## 【Step 4】リモートリポジトリを登録

```bash
git remote add origin https://github.com/yoshidadan/Excel-SpecificationChecker.git
```

**確認コマンド**：
```bash
git remote -v
```

**出力例**：
```
origin  https://github.com/yoshidadan/Excel-SpecificationChecker.git (fetch)
origin  https://github.com/yoshidadan/Excel-SpecificationChecker.git (push)
```

---

## 【Step 5】ファイルをステージングエリアに追加

```bash
git add .
```

**確認コマンド**：
```bash
git status
```

**出力例**：
```
On branch main

No commits yet

Changes to be committed:
  new file:   .gitignore
  new file:   CHANGELOG.md
  new file:   README.md
  new file:   docs/Architecture.md
  new file:   vba/SpecificationChecker_Complete.txt
```

---

## 【Step 6】初回コミット

```bash
git commit -m "Initial commit: 仕様書差分チェックマクロ v1.0

- LCS ベース差分検出（WinMerge同等の精度）
- ハンガリアンアルゴリズム最適ペアリング
- スコア1.0先行確定＋重複使用防止
- 赤文字インライン処理対応
- 最大100,000行サポート
- 7マッチタイプ自動分類
- 条件付き書式による色分け表示
- 完全ドキュメント完備"
```

**出力例**：
```
[main (root-commit) a1b2c3d] Initial commit: 仕様書差分チェックマクロ v1.0
 5 files changed, 2847 insertions(+)
 create mode 100644 .gitignore
 create mode 100644 CHANGELOG.md
 create mode 100644 README.md
 create mode 100644 docs/Architecture.md
 create mode 100644 vba/SpecificationChecker_Complete.txt
```

---

## 【Step 7】GitHub にプッシュ

```bash
git branch -M main
git push -u origin main
```

**初回実行時の認証方法**：

### 方法A：GitHub Personal Access Token（推奨）
```
1. https://github.com/settings/tokens に移動
2. [Generate new token] → [Generate new token (classic)]
3. Scopes で "repo" にチェック
4. Generate token でトークンコピー
5. ターミナルで "git push -u origin main" 実行
6. Username: your_github_username
7. Password: （コピーしたトークンをペースト）
```

### 方法B：SSH Key（上級）
```
ssh-keygen -t ed25519 -C "your_email@example.com"
```

詳細は: https://docs.github.com/en/authentication

---

## 【Step 8】GitHub で確認

```
1. https://github.com/yoshidadan/Excel-SpecificationChecker を開く
2. ファイル一覧が表示されることを確認
3. README.md が自動的にプレビュー表示される
```

---

## 【トラブルシューティング】

### Q. "fatal: could not read Username"

**原因**: 認証情報が登録されていない

**対処**:
```bash
# Windows Credential Manager をリセット
# [コントロールパネル] → [資格情報マネージャー] → GitHub関連を削除
# 再度 "git push" を実行して認証
```

### Q. "Permission denied (publickey)"

**原因**: SSH Keyが正しく設定されていない

**対処**:
```bash
# Personal Access Token で再度試す（方法Aを使用）
```

### Q. "Everything up-to-date"

**確認事項**:
```bash
git log
git remote -v
```

---

## 【確認コマンド一覧】

```bash
# ログを確認
git log --oneline

# ファイル一覧
git ls-files

# リモート情報
git remote -v

# ブランチ確認
git branch -a

# 最後のプッシュ時刻確認
git log --oneline origin/main
```

---

## 【次のステップ】

### テンプレートファイルの作成
```
templates/Template_SpecChecker_Base.xlsm を作成
（別途手順を提供します）
```

### サンプルデータの追加
```
templates/SampleData_Demo.xlsx を作成
（仕様書①②のサンプル）
```

### 更新時の手順
```bash
# 修正後のコミット
git add .
git commit -m "Fix: [修正内容]"
git push origin main
```

---

**完了したら、ぎゆう様から連絡をください！**

次のステップ:
1. ✅ GitHub初期化完了
2. ⏳ テンプレートXLSMの作成（次回）
3. ⏳ サンプルデータの追加（次回）

