# Outlook Send Guard - 誤送信防止アドイン

Outlook on the web向けの誤送信防止アドインです。メール送信時に確認ダイアログを表示し、外部宛先への誤送信や添付ファイル忘れを防ぎます。

## 主要機能

### 送信前確認ダイアログ

メール送信時に自動的に確認ダイアログを表示します。すべての確認項目にチェックを入れないと送信できません。

- **外部ドメイン宛先の確認**
  - 社外メールアドレスを一覧表示（TO/CC/BCC区別あり）
  - 各アドレスを個別にチェックして確認
  - 社内ドメインは設定ファイルで管理

- **件名の確認**
  - 件名の内容を表示
  - チェックボックスで確認

- **本文の確認**
  - 本文の先頭部分を表示（最大1000文字）
  - チェックボックスで確認

- **添付ファイルの確認**
  - 添付ファイルを一覧表示（ファイル名、サイズ）
  - 各ファイルを個別にチェックして確認

### 送信前チェック機能

- **添付ファイルチェック**
  - 本文中に「添付」「別紙」などのキーワードがあるのに添付ファイルがない場合に警告
  - チェックキーワードと除外キーワードをカスタマイズ可能

- **宛先チェック**
  - 宛先未入力チェック
  - 宛先件数の警告（設定値以上の宛先がある場合）
  - 複数の社外宛先がTo/Ccにある場合、Bcc使用を推奨

- **件名・本文チェック**
  - 件名未入力チェック
  - 本文未入力チェック
  - 本文中のTODOマーカー検知（「TODO」「あとで」など）

### 動的なダイアログサイズ

メールの内容（宛先数、添付ファイル数など）に応じて、確認ダイアログのサイズを自動調整します。

### 設定画面

リボンボタンから設定画面を開き、各種設定をカスタマイズできます。

## 対応プラットフォーム

- ✅ Outlook on the web（動作確認済み）
- ✅ Outlook on Windows（新版）
- ✅ Outlook on Mac
- ⚠️ Outlook on iOS/Android（一部機能制限あり）

## インストール方法

### 1. マニフェストファイルのダウンロード

以下のURLからマニフェストファイルをダウンロードします：

```
https://masaspc.github.io/OutlookMailAddin/manifest.xml
```

### 2. Outlook on the webでのインストール

1. Outlook on the webを開く（https://outlook.office.com）
2. 設定（⚙️アイコン）> すべてのOutlook設定を表示
3. 全般 > アドインの管理
4. 「カスタムアドインを追加」> 「ファイルから追加」
5. ダウンロードした `manifest.xml` ファイルを選択
6. 「インストール」をクリック

### 3. 動作確認

1. 新規メールを作成
2. 外部メールアドレスを宛先に入力
3. 送信ボタンをクリック
4. 確認ダイアログが表示されることを確認

## 使い方

### メール送信時の確認

1. Outlookでメールを作成
2. 送信ボタンをクリック
3. 確認ダイアログが自動的に開く
4. 表示される内容を確認：
   - 外部ドメイン宛先（ある場合）
   - 件名
   - 本文（プレビュー）
   - 添付ファイル（ある場合）
5. すべての項目にチェックを入れる
6. 「送信する」ボタンをクリック

### 設定の変更

1. メール作成画面のリボンにある「Send Guard」グループを探す
2. 「誤送信防止設定」ボタンをクリック
3. 設定画面が右側に表示される
4. 各種設定を変更：
   - 📎 添付ファイルチェック
   - 👥 宛先チェック
   - 📋 件名チェック
   - 📝 本文チェック
5. 「保存」ボタンをクリック

## 管理者向け設定

### 社内ドメインの変更

社内ドメインは設定ファイルで管理されており、ユーザーは変更できません。

**変更方法：**

1. `src/types.ts` を開く
2. `DEFAULT_SETTINGS.recipientCheck.internalDomains` を編集

```typescript
recipientCheck: {
  enabled: true,
  internalDomains: ['tokyobaynet.co.jp', 'example.com'],  // 複数設定可能
  warnExternalRecipients: true,
  warnMultipleExternal: true,
  maxRecipientCount: 50,
},
```

3. 再ビルドとデプロイ

```bash
npm run build
git add -A
git commit -m "Update internal domains"
git push origin <branch-name>
```

4. GitHub Actionsが自動的にデプロイ

### その他のデフォルト設定の変更

`src/types.ts` の `DEFAULT_SETTINGS` を編集することで、以下のデフォルト設定を変更できます：

- 添付ファイルチェックのキーワード
- 除外キーワード
- TODOキーワード
- 宛先件数の上限

## デプロイ

このアドインはGitHub Pagesで自動デプロイされます。

### 初回セットアップ

1. GitHubリポジトリの Settings > Pages
2. Source: GitHub Actions を選択
3. Permissions: Read and write permissions を設定

### デプロイ手順

1. コードを変更
2. コミット＆プッシュ

```bash
git add -A
git commit -m "Update feature"
git push origin <branch-name>
```

3. GitHub Actionsが自動的にビルド＆デプロイ
4. デプロイ完了後、ユーザーはアドインを再インストール

### デプロイ確認

デプロイが成功したかは、以下のURLで確認できます：

```
https://masaspc.github.io/OutlookMailAddin/
```

## 開発

### 前提条件

- Node.js (v18以上)
- npm

### セットアップ

```bash
# 依存パッケージのインストール
npm install

# ビルド
npm run build
```

### ファイル構成

```
OutlookMailAddin/
├── manifest.xml                    # アドインマニフェスト
├── src/
│   ├── types.ts                   # 型定義とデフォルト設定
│   ├── taskpane/
│   │   ├── taskpane.tsx          # メインコンポーネント
│   │   ├── taskpane.css          # スタイル
│   │   └── components/
│   │       ├── ConfirmationPanel.tsx  # 確認ダイアログ
│   │       └── SettingsPanel.tsx      # 設定画面
│   ├── commands/
│   │   └── commands.ts           # OnMessageSendイベント処理
│   ├── checker/
│   │   ├── AttachmentChecker.ts
│   │   ├── RecipientChecker.ts
│   │   ├── SubjectChecker.ts
│   │   ├── BodyChecker.ts
│   │   └── CheckerManager.ts
│   ├── storage/
│   │   └── SettingsStorage.ts    # 設定の保存・読み込み
│   └── utils/
│       ├── DomainChecker.ts
│       └── KeywordMatcher.ts
└── assets/
    └── icon-*.png
```

## トラブルシューティング

### アドインが表示されない

**原因:** マニフェストファイルが正しくインストールされていない

**解決方法:**
1. アドイン管理画面でアドインが表示されているか確認
2. 表示されていない場合は、再インストール

### 確認ダイアログが表示されない

**原因:** メール作成画面でアドインが有効になっていない

**解決方法:**
1. ブラウザの開発者ツール（F12）を開く
2. Console タブでエラーメッセージを確認
3. 「Error 9032」が表示される場合は、既知の問題（displayInIframe制限）

### 設定が保存されない

**原因:** Office.js RoamingSettingsのエラー

**解決方法:**
1. ブラウザのキャッシュをクリア
2. Outlookからサインアウトして再度サインイン
3. アドインを再インストール

### リボンボタンが表示されない

**原因:** manifest.xmlの設定不足またはキャッシュの問題

**解決方法:**
1. アドインを削除して再インストール
2. ブラウザのキャッシュをクリア
3. Outlookをリロード（Ctrl+F5）

### 外部ドメインが正しく判定されない

**原因:** 社内ドメイン設定が正しくない

**解決方法:**
1. `src/types.ts` の `internalDomains` を確認
2. ドメインが小文字で入力されているか確認
3. 再ビルド＆デプロイ

## 技術的な特徴

### Office.js API

- **OnMessageSend イベント**: メール送信時に処理を実行
- **SoftBlock モード**: 送信を一時保留し、ユーザー確認後に送信
- **displayDialogAsync**: 確認ダイアログを表示
- **promptBeforeOpen: false**: 許可ダイアログを非表示
- **RoamingSettings**: ユーザー設定を保存

### パフォーマンス最適化

- **動的サイズ計算**: メール内容に応じてダイアログサイズを調整
- **並列チェック**: すべてのチェックを並列実行
- **本文の切り詰め**: 長い本文はURLパラメータオーバーフロー防止のため1000文字に制限

### セキュリティ

- **HTTPS必須**: GitHub Pagesで自動的にHTTPS配信
- **Same-Origin Policy**: Office.jsのセキュリティモデルに準拠

## 既知の制限事項

1. **URLバーの表示**: `displayInIframe: true`が使えないため、ダイアログにURLバーが表示されます（Office.jsの制限）
2. **モバイル対応**: iOS/Androidでは一部機能が制限される場合があります
3. **添付ファイルサイズ**: 非常に大きな添付ファイルは表示に時間がかかる場合があります

## ライセンス

MIT

## サポート

問題が発生した場合は、GitHubのIssuesで報告してください。

## 参考リンク

- [Office Add-ins ドキュメント](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Outlook Add-ins API リファレンス](https://learn.microsoft.com/en-us/javascript/api/outlook)
- [GitHub Pages](https://pages.github.com/)
