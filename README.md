# Outlook Send Guard - 誤送信防止アドイン

Outlook Okanクロスプラットフォーム版と同等の誤送信防止機能を持つOutlookアドインです。

## 主要機能

### 送信前チェック機能

- **添付ファイルチェック**
  - 本文中に「添付」「別紙」などのキーワードがあるのに添付ファイルがない場合に警告
  - チェックキーワードのカスタマイズ可能
  - 除外キーワードの設定（「添付不要」など）

- **宛先チェック**
  - 社外ドメインへの送信時に警告
  - Bcc漏れチェック（To/Ccに複数の社外アドレスがある場合）
  - 宛先未入力チェック
  - 宛先件数の警告（設定値以上の宛先がある場合）

- **件名・本文チェック**
  - 件名未入力チェック
  - 件名に「Re:」がないのに本文が返信形式の場合に警告
  - 本文未入力チェック
  - 本文中のTODOマーカー検知（「TODO」「あとで」など）

- **時間帯チェック**
  - 深夜・早朝送信の警告（時間帯は設定可能）
  - 休日送信の警告

## 対応プラットフォーム

- Outlook on the web
- Outlook on Windows (新版)
- Outlook on Mac
- Outlook on iOS/Android（可能な範囲で）

## セットアップ

### 前提条件

- Node.js (v16以上)
- npm または yarn

### インストール

```bash
# 依存パッケージのインストール
npm install
```

### 開発用証明書の作成

Outlookアドインの開発にはHTTPSが必要です。以下のコマンドで開発用証明書を作成します。

```bash
npx office-addin-dev-certs install
```

### アイコンファイルの準備

`assets/` ディレクトリに以下のアイコンファイルを配置してください:

- `icon-16.png` (16x16px)
- `icon-32.png` (32x32px)
- `icon-64.png` (64x64px)
- `icon-80.png` (80x80px)

## 開発

### ビルド

```bash
# 開発ビルド
npm run build:dev

# 本番ビルド
npm run build
```

### 開発サーバーの起動

```bash
npm run dev-server
```

開発サーバーは `https://localhost:3000` で起動します。

### Outlookへのサイドロード

#### Outlook on the web

1. Outlookを開く
2. 設定 > すべてのOutlook設定を表示 > 全般 > アドインの管理
3. 「カスタムアドインを追加」> 「ファイルから追加」
4. `manifest.xml` ファイルを選択

#### Outlook on Windows/Mac

```bash
npm start
```

上記コマンドでOutlookが自動的に起動し、アドインがサイドロードされます。

### マニフェストの検証

```bash
npm run validate
```

## 使い方

### 基本的な使用方法

1. Outlookでメールを作成
2. 送信ボタンをクリック
3. 自動的に別ウィンドウ（ダイアログ）が開き、チェックが実行される
4. 問題がある場合は警告が表示されるので、内容を確認
5. 「送信する」ボタンをクリックして送信、または「編集に戻る」で修正

### 設定のカスタマイズ

1. ダイアログウィンドウの「⚙️ 設定」ボタンをクリック
2. 各チェック項目の有効/無効を切り替え
3. キーワード、ドメイン、時間帯などをカスタマイズ
4. 「保存」ボタンをクリック

※設定は送信チェック時に開くダイアログウィンドウから行います

## アーキテクチャ

### ファイル構成

```
outlook-send-guard/
├── manifest.xml              # アドインマニフェスト
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Task Pane UI
│   │   ├── taskpane.tsx     # Task Pane メインコンポーネント
│   │   └── components/
│   │       ├── CheckResultList.tsx   # チェック結果表示
│   │       └── SettingsPanel.tsx     # 設定画面
│   ├── commands/
│   │   └── commands.ts      # リボンコマンド処理
│   ├── checker/
│   │   ├── AttachmentChecker.ts
│   │   ├── RecipientChecker.ts
│   │   ├── SubjectChecker.ts
│   │   ├── BodyChecker.ts
│   │   ├── TimeChecker.ts
│   │   └── CheckerManager.ts
│   ├── storage/
│   │   └── SettingsStorage.ts   # 設定の保存・読み込み
│   └── utils/
│       ├── DomainChecker.ts
│       └── KeywordMatcher.ts
└── assets/
    └── icon-*.png
```

### 処理フロー

1. ユーザーが送信ボタンをクリック
2. OnMessageSendイベントが発火
3. 送信を一時的に保留
4. **別ウィンドウ（ダイアログ）を開く**
5. ダイアログ内で各種チェックを並列実行
6. チェック結果を表示
7. ユーザーが確認後、送信または編集に戻る
8. ダイアログから親ウィンドウにメッセージを送信
9. 親ウィンドウが送信を許可/キャンセル

### 技術的な特徴

- **ダイアログ方式**: Office.js の `displayDialogAsync` を使用して別ウィンドウを開く
- **5秒タイムアウト制限の回避**: ダイアログ方式を採用することで、Outlookアドインの5秒タイムアウト制限を回避
- **並列チェック**: すべてのチェックを並列実行することでパフォーマンスを最適化
- **ローカルストレージ**: Office.js RoamingSettingsを使用して設定を保存
- **React + Fluent UI**: モダンなUIフレームワークを使用
- **messageParent API**: ダイアログから親ウィンドウへのメッセージ送信で送信制御

## トラブルシューティング

### アドインが読み込まれない

- マニフェストファイルが正しいか確認
- 開発サーバーが起動しているか確認（`https://localhost:3000`）
- 証明書が正しくインストールされているか確認

### チェックが実行されない

- 設定でチェック項目が有効になっているか確認
- ブラウザの開発者ツールでエラーが出ていないか確認

### 設定が保存されない

- Office.js RoamingSettingsが正しく動作しているか確認
- ブラウザのコンソールでエラーメッセージを確認

## ライセンス

MIT

## 参考

- [Office Add-ins ドキュメント](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Outlook Okan](https://github.com/)
