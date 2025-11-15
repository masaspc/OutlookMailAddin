# ファイルからのアドインインストール手順

URLベースのインストールが動作しない場合、manifest.xmlファイルを直接アップロードしてインストールできます。

## 1. manifest.xmlファイルの取得

### 方法A: GitHubリポジトリからダウンロード

1. GitHubリポジトリにアクセス: https://github.com/masaspc/OutlookMailAddin
2. `manifest-production.xml` ファイルを開く
3. 「Raw」ボタンをクリック
4. 右クリック → 「名前を付けて保存」で `manifest.xml` として保存

### 方法B: ローカルファイルを使用

プロジェクトフォルダの `manifest-production.xml` をコピーして `manifest.xml` にリネームします。

```bash
cp manifest-production.xml manifest.xml
```

## 2. Outlook on the web でのインストール

1. Outlook on the web (https://outlook.office.com) にアクセス
2. 設定 (⚙️) → **すべてのOutlook設定を表示**
3. **全般** → **アドインの管理**
4. **+ カスタムアドインを追加** → **ファイルから追加**
5. ダウンロードした `manifest.xml` ファイルを選択
6. **インストール** をクリック

## 3. Outlook Desktop (Windows/Mac) でのインストール

### Windows版

1. Outlookを起動
2. **ファイル** → **Office アドインの管理**
3. **マイ アドイン** タブをクリック
4. **+ カスタムアドインを追加** → **ファイルから追加...**
5. ダウンロードした `manifest.xml` ファイルを選択
6. **OK** をクリック

### Mac版

1. Outlookを起動
2. **ツール** → **Office アドインの管理**
3. **マイ アドイン** タブをクリック
4. **+ カスタムアドインを追加** → **ファイルから追加...**
5. ダウンロードした `manifest.xml` ファイルを選択
6. **OK** をクリック

## 4. インストールの確認

1. 新規メール作成画面を開く
2. メールを作成して「送信」ボタンをクリック
3. Outlook Send Guardの確認ダイアログが表示されることを確認

## 5. トラブルシューティング

### アドインが表示されない場合

1. **Outlookを再起動**してみてください
2. **ブラウザのキャッシュをクリア**してください（Web版の場合）
3. manifest.xmlのURLが正しいか確認:
   - 全てのURL: `https://masaspc.github.io/OutlookMailAddin/`

### エラーが表示される場合

#### "アドインマニフェストが無効です"
- manifest.xmlファイルが破損していないか確認
- XMLの形式が正しいか確認

#### "このアドインをインストールできません"
- Outlookのバージョンが要件を満たしているか確認（Mailbox 1.12以上）
- 管理者権限が必要な場合があります

#### アドインは表示されるが動作しない
1. ブラウザの開発者ツール (F12) を開く
2. コンソールでエラーを確認
3. GitHub Pagesのデプロイが完了しているか確認:
   - https://masaspc.github.io/OutlookMailAddin/commands.html にアクセスできるか
   - https://masaspc.github.io/OutlookMailAddin/taskpane.html にアクセスできるか

### GitHub Pagesが有効になっているか確認

1. GitHubリポジトリに移動: https://github.com/masaspc/OutlookMailAddin
2. **Settings** → **Pages**
3. **Source** が **GitHub Actions** になっているか確認
4. **Actions** タブで最新のデプロイが成功しているか確認

## 6. URLベースのインストールを試す前に

GitHub Pagesのデプロイが完了するまで数分かかる場合があります。以下を確認してください:

1. **Actions** タブでワークフローが完了しているか（緑のチェックマーク）
2. https://masaspc.github.io/OutlookMailAddin/manifest.xml にアクセスできるか

デプロイが完了していれば、URLベースのインストールも動作するはずです。

## 参考リンク

- [Office アドインのサイドロード](https://learn.microsoft.com/ja-jp/office/dev/add-ins/testing/test-debug-office-add-ins)
- [アドインマニフェストのトラブルシューティング](https://learn.microsoft.com/ja-jp/office/dev/add-ins/testing/troubleshoot-manifest)
