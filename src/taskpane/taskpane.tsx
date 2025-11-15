import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { CheckResultList } from './components/CheckResultList';
import { SettingsPanel } from './components/SettingsPanel';
import { CheckResult, MailData, Settings } from '../types';
import { SettingsStorage } from '../storage/SettingsStorage';
import { CheckerManager } from '../checker/CheckerManager';
import './taskpane.css';

interface AppState {
  checkResults: CheckResult[];
  isLoading: boolean;
  isSettingsOpen: boolean;
  settings: Settings;
}

class App extends React.Component<{}, AppState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      checkResults: [],
      isLoading: true,
      isSettingsOpen: false,
      settings: SettingsStorage.loadSettings(),
    };
  }

  async componentDidMount() {
    // Office.jsの初期化を待つ
    await Office.onReady();

    // メールのチェックを実行
    await this.runChecks();
  }

  /**
   * メールのチェックを実行する
   */
  runChecks = async () => {
    this.setState({ isLoading: true });

    try {
      const mailData = await this.getMailData();
      const results = await CheckerManager.checkAll(mailData, this.state.settings);

      this.setState({
        checkResults: results,
        isLoading: false,
      });
    } catch (error) {
      console.error('チェック実行エラー:', error);
      this.setState({
        checkResults: [
          {
            severity: 'error',
            category: 'body',
            message: 'チェックエラー',
            details: 'メールのチェック中にエラーが発生しました。',
          },
        ],
        isLoading: false,
      });
    }
  };

  /**
   * メールデータを取得する
   */
  getMailData = async (): Promise<MailData> => {
    return new Promise((resolve, reject) => {
      const item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error('メールアイテムが見つかりません'));
        return;
      }

      // 件名を取得
      const subject = item.subject || '';

      // 本文を取得
      item.body.getAsync(Office.CoercionType.Html, async (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error('本文の取得に失敗しました'));
          return;
        }

        const body = result.value;

        // 宛先を取得
        const to = await this.getRecipients(item.to);
        const cc = await this.getRecipients(item.cc);
        const bcc = await this.getRecipients(item.bcc);

        // 添付ファイルを取得
        const attachments = item.attachments || [];

        resolve({
          subject,
          body,
          to,
          cc,
          bcc,
          attachments,
        });
      });
    });
  };

  /**
   * 受信者のメールアドレスを取得する
   */
  getRecipients = async (recipients: Office.Recipients | undefined): Promise<string[]> => {
    if (!recipients) return [];

    return new Promise((resolve) => {
      recipients.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const emails = result.value.map((r) => r.emailAddress);
          resolve(emails);
        } else {
          resolve([]);
        }
      });
    });
  };

  /**
   * 送信ボタンをクリックしたときの処理
   * 親ウィンドウ（Outlook）にメッセージを送信して、メール送信を許可する
   */
  handleSend = () => {
    try {
      Office.context.ui.messageParent(JSON.stringify({ action: 'send' }));
    } catch (error) {
      console.error('送信メッセージの送信に失敗しました:', error);
    }
  };

  /**
   * キャンセルボタン（編集に戻る）をクリックしたときの処理
   * 親ウィンドウ（Outlook）にメッセージを送信して、メール送信をキャンセルする
   */
  handleCancel = () => {
    try {
      Office.context.ui.messageParent(JSON.stringify({ action: 'cancel' }));
    } catch (error) {
      console.error('キャンセルメッセージの送信に失敗しました:', error);
    }
  };

  /**
   * 設定を開く
   */
  handleOpenSettings = () => {
    this.setState({ isSettingsOpen: true });
  };

  /**
   * 設定を閉じる
   */
  handleCloseSettings = () => {
    this.setState({ isSettingsOpen: false });
  };

  /**
   * 設定を保存する
   */
  handleSaveSettings = async (settings: Settings) => {
    try {
      await SettingsStorage.saveSettings(settings);
      this.setState({ settings, isSettingsOpen: false });
      // 設定が変更されたらチェックを再実行
      await this.runChecks();
    } catch (error) {
      console.error('設定の保存に失敗しました:', error);
      alert('設定の保存に失敗しました。もう一度お試しください。');
    }
  };

  render() {
    const { checkResults, isLoading, isSettingsOpen, settings } = this.state;

    return (
      <FluentProvider theme={webLightTheme}>
        <div className="app-container">
          {isSettingsOpen ? (
            <SettingsPanel
              settings={settings}
              onSave={this.handleSaveSettings}
              onClose={this.handleCloseSettings}
            />
          ) : (
            <>
              <div className="header">
                <h2>送信前チェック</h2>
                <button className="settings-button" onClick={this.handleOpenSettings}>
                  ⚙️ 設定
                </button>
              </div>

              <div className="content">
                {isLoading ? (
                  <div className="loading">チェック中...</div>
                ) : (
                  <CheckResultList results={checkResults} />
                )}
              </div>

              <div className="footer">
                <button className="send-button" onClick={this.handleSend}>
                  送信する
                </button>
                <button className="cancel-button" onClick={this.handleCancel}>
                  編集に戻る
                </button>
              </div>
            </>
          )}
        </div>
      </FluentProvider>
    );
  }
}

// DOMにレンダリング
Office.onReady(() => {
  const container = document.getElementById('root');
  if (container) {
    const root = createRoot(container);
    root.render(<App />);
  }
});
