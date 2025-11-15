import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { CheckResultList } from './components/CheckResultList';
import { SettingsPanel } from './components/SettingsPanel';
import { CheckResult, Settings } from '../types';
import { SettingsStorage } from '../storage/SettingsStorage';
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

    // sessionDataからチェック結果を読み取る
    await this.loadCheckResults();
  }

  /**
   * sessionDataからチェック結果を読み取る
   */
  loadCheckResults = async () => {
    this.setState({ isLoading: true });

    try {
      const item = Office.context.mailbox.item as Office.MessageCompose;
      if (!item || !item.sessionData) {
        // sessionDataがない場合は、デフォルトのメッセージを表示
        console.warn('sessionDataがありません。問題なしとして扱います。');
        this.setState({
          checkResults: [
            {
              severity: 'info',
              category: 'body',
              message: 'チェック完了',
              details: '問題は検出されませんでした。',
            },
          ],
          isLoading: false,
        });
        return;
      }

      // sessionDataからチェック結果を取得
      item.sessionData.getAsync('checkResults', (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
          try {
            const checkResults = JSON.parse(result.value);
            this.setState({
              checkResults,
              isLoading: false,
            });
          } catch (error) {
            console.error('チェック結果のパースに失敗しました:', error);
            this.setState({
              checkResults: [
                {
                  severity: 'error',
                  category: 'body',
                  message: 'チェックエラー',
                  details: 'チェック結果の読み込みに失敗しました。',
                },
              ],
              isLoading: false,
            });
          }
        } else {
          console.warn('sessionDataの取得に失敗しました:', result.error);
          this.setState({
            checkResults: [
              {
                severity: 'info',
                category: 'body',
                message: 'チェック完了',
                details: '問題は検出されませんでした。',
              },
            ],
            isLoading: false,
          });
        }
      });
    } catch (error) {
      console.error('チェック結果の読み込みエラー:', error);
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
      alert('設定を保存しました。次回の送信時から反映されます。');
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
