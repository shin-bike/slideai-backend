# SlideAI Backend

Kimuraさんの手順通りに実装したPowerPoint自動生成バックエンドです。

## 処理フロー

```
① 構成設計        Claude APIでスライド構成を自動設計
② テンプレート選択  941枚のメタデータから最適なスライドを選択
③ コンテンツ生成   Claude APIで各スライドの内容を詳細生成
④ pptx構築       実際のpptxテンプレートからスライドをコピー → テキスト流し込み
⑤ 視覚QA        Claude Visionでレイアウト崩れを自動検出
```

## Railwayへのデプロイ手順

### 1. GitHubリポジトリを作成

```bash
cd slideai_backend
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_NAME/slideai-backend.git
git push -u origin main
```

### 2. Railwayでデプロイ

1. https://railway.app にアクセス
2. 「New Project」→「Deploy from GitHub repo」
3. 作成したリポジトリを選択
4. 環境変数は不要（APIキーはフロントエンドから毎回送信）
5. デプロイ完了後、「Settings」→「Domains」でURLを取得

### 3. フロントエンドに設定

`frontend.html` の「バックエンドURL」欄にRailwayのURLを入力して使用。

## ローカル開発

```bash
# 依存関係インストール
pip install -r requirements.txt

# LibreOffice（視覚QAに必要）
# Mac: brew install libreoffice
# Ubuntu: sudo apt install libreoffice

# poppler（PDF→画像変換）
# Mac: brew install poppler
# Ubuntu: sudo apt install poppler-utils

# 起動
uvicorn main:app --reload --port 8000
```

API確認: http://localhost:8000/health

## ファイル構成

```
slideai_backend/
├── main.py              # FastAPIバックエンド（メインロジック）
├── requirements.txt     # Python依存関係
├── Procfile             # Railway/Heroku用
├── railway.json         # Railway設定
├── nixpacks.toml        # LibreOffice/poppler自動インストール
├── frontend.html        # フロントエンド（バックエンドAPIを呼ぶ）
├── metadata.json        # 941枚のテンプレートメタデータ
├── DTC.pptx             # DTCテンプレートライブラリ（2013）
├── BCG.pptx             # BCGテンプレート
├── McKinsey.pptx        # McKinseyテンプレート
└── PowerLibrary2009.pptx # PowerLibrary（2009）
```

## 注意事項

- `metadata.json` と `.pptx` ファイルはGitにコミットしてください（Railway側に必要）
- pptxファイルは合計約12MB、Railway無料枠で問題なし
- 視覚QAはLibreOfficeとpopplerが必要（nixpacks.tomlで自動インストール）
　
