# GitHub Pagesへのデプロイ手順

## 1. GitHubリポジトリの設定

### GitHub Pagesを有効化

1. GitHubのリポジトリページに移動: https://github.com/masaspc/OutlookMailAddin
2. **Settings** タブをクリック
3. 左サイドバーから **Pages** をクリック
4. **Source** セクションで以下を選択:
   - Source: **GitHub Actions**

## 2. デプロイの実行

### 自動デプロイ

mainブランチにpushすると、自動的にGitHub Actionsが実行され、GitHub Pagesにデプロイされます。

```bash
git add .
git commit -m "Deploy to GitHub Pages"
git push origin main
```

### 手動デプロイ

GitHubリポジトリの **Actions** タブから、**Deploy to GitHub Pages** ワークフローを手動実行できます。

## 3. デプロイ後の確認

デプロイが完了すると、以下のURLでアクセス可能になります:

- **Add-in URL**: https://masaspc.github.io/OutlookMailAddin/
- **Manifest URL**: https://masaspc.github.io/OutlookMailAddin/manifest.xml

## 4. Outlookへのインストール

### Web版Outlook

1. Outlook on the web (https://outlook.office.com) にアクセス
2. 設定 (⚙️) → **すべてのOutlook設定を表示**
3. **全般** → **アドインの管理**
4. **+ アドインを追加** → **URLから追加**
5. 以下のURLを入力:
   ```
   https://masaspc.github.io/OutlookMailAddin/manifest.xml
   ```
6. **追加** をクリック

### デスクトップ版Outlook (Windows/Mac)

1. Outlookを起動
2. **ファイル** → **アドインの取得**
3. **マイ アドイン** → **カスタム アドイン** → **URLから追加**
4. 以下のURLを入力:
   ```
   https://masaspc.github.io/OutlookMailAddin/manifest.xml
   ```

## 5. トラブルシューティング

### デプロイが失敗する場合

1. **Actions** タブでエラーログを確認
2. npm installやビルドエラーの場合:
   ```bash
   npm install
   npm run build
   ```
   をローカルで実行して確認

### アドインが動作しない場合

1. ブラウザの開発者ツールでコンソールエラーを確認
2. manifest.xmlのURLが正しいか確認
3. HTTPS接続であることを確認

## 6. 開発環境とのURL切り替え

- **開発環境**: `manifest.xml` (localhost:3000)
- **本番環境**: `manifest-production.xml` (GitHub Pages)

開発時はlocalhost版を使用し、デプロイ時は自動的に本番版が使用されます。

## 7. カスタムドメインの設定（オプション）

独自ドメインを使用したい場合:

1. GitHub Pagesの設定で **Custom domain** を設定
2. `manifest-production.xml` のURLを変更
3. 再デプロイ

## 参考リンク

- [GitHub Pages Documentation](https://docs.github.com/pages)
- [Office Add-ins Documentation](https://learn.microsoft.com/office/dev/add-ins/)
