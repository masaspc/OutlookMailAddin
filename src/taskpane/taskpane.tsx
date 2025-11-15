import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ConfirmationPanel } from './components/ConfirmationPanel';
import { SettingsPanel } from './components/SettingsPanel';
import { CheckResult, Settings } from '../types';
import { SettingsStorage } from '../storage/SettingsStorage';
import { CheckerManager } from '../checker/CheckerManager';
import './taskpane.css';

interface MailDataDisplay {
  subject: string;
  body: string;
  to: string[];
  cc: string[];
  bcc: string[];
  attachments: Array<{
    id: string;
    name: string;
    size: number;
    attachmentType: string;
  }>;
}

interface AppState {
  checkResults: CheckResult[];
  mailData: MailDataDisplay | null;
  isLoading: boolean;
  isSettingsOpen: boolean;
  settings: Settings;
}

class App extends React.Component<{}, AppState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      checkResults: [],
      mailData: null,
      isLoading: true,
      isSettingsOpen: false,
      settings: SettingsStorage.loadSettings(),
    };
  }

  async componentDidMount() {
    // Office.jsの初期化を待つ
    await Office.onReady();

    // URLパラメータからチェック結果を読み取る
    this.loadCheckResults();
  }

  /**
   * URLパラメータからメールデータを読み取り、チェック処理を実行する
   */
  loadCheckResults = async () => {
    this.setState({ isLoading: true });

    try {
      // URLパラメータを取得
      const urlParams = new URLSearchParams(window.location.search);
      const dataBase64 = urlParams.get('data');

      if (!dataBase64) {
        // URLパラメータがない場合は、リボンボタンから開かれたので設定画面を表示
        console.log('リボンボタンから開かれました。設定画面を表示します。');
        this.setState({
          checkResults: [],
          mailData: null,
          isLoading: false,
          isSettingsOpen: true,
        });
        return;
      }

      // Base64デコードしてJSONパース
      const dataJson = decodeURIComponent(atob(dataBase64));
      const data = JSON.parse(dataJson);

      const mailData = data.mailData;
      if (!mailData) {
        throw new Error('メールデータがありません');
      }

      // 設定を読み込み
      const settings = this.state.settings;

      // チェックを実行（ダイアログ内で実行）
      const checkResults = await CheckerManager.checkAll(
        {
          subject: mailData.subject || '',
          body: mailData.body || '',
          to: mailData.to || [],
          cc: mailData.cc || [],
          bcc: mailData.bcc || [],
          attachments: mailData.attachments || [],
        },
        settings
      );

      this.setState({
        checkResults,
        mailData,
        isLoading: false,
      });
    } catch (error) {
      console.error('データの読み込みエラー:', error);
      this.setState({
        checkResults: [
          {
            severity: 'error',
            category: 'body',
            message: 'データエラー',
            details: 'メールデータの読み込みに失敗しました。',
          },
        ],
        mailData: null,
        isLoading: false,
      });
    }
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
      this.setState({ settings });
      // 保存成功（メッセージなし）
    } catch (error) {
      console.error('設定の保存に失敗しました:', error);
      // エラーメッセージをコンソールに出力
    }
  };

  render() {
    const { checkResults, mailData, isLoading, isSettingsOpen, settings } = this.state;

    // リボンボタンから開いた場合（mailDataがnull）は、常に設定画面を表示
    if (mailData === null) {
      return (
        <FluentProvider theme={webLightTheme}>
          <div className="app-container">
            <SettingsPanel
              settings={settings}
              onSave={this.handleSaveSettings}
              onClose={this.handleCloseSettings}
            />
          </div>
        </FluentProvider>
      );
    }

    // メール送信時の確認ダイアログの場合
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
              {isLoading ? (
                <div className="loading">チェック中...</div>
              ) : mailData ? (
                <ConfirmationPanel
                  mailData={mailData}
                  checkResults={checkResults}
                  settings={settings}
                  onSend={this.handleSend}
                  onCancel={this.handleCancel}
                />
              ) : (
                <div className="error-message">
                  <h2>エラー</h2>
                  <p>メールデータの読み込みに失敗しました。</p>
                  <button className="cancel-button" onClick={this.handleCancel}>
                    閉じる
                  </button>
                </div>
              )}
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
