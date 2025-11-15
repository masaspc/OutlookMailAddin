import * as React from 'react';
import { CheckResult, Settings } from '../../types';

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

interface ConfirmationPanelProps {
  mailData: MailDataDisplay;
  checkResults: CheckResult[];
  settings: Settings;
  onSend: () => void;
  onCancel: () => void;
}

interface ConfirmationState {
  externalRecipientsConfirmed: Set<string>;
  subjectConfirmed: boolean;
  bodyConfirmed: boolean;
  attachmentsConfirmed: Set<string>;
}

export class ConfirmationPanel extends React.Component<ConfirmationPanelProps, ConfirmationState> {
  constructor(props: ConfirmationPanelProps) {
    super(props);
    this.state = {
      externalRecipientsConfirmed: new Set(),
      subjectConfirmed: false,
      bodyConfirmed: false,
      attachmentsConfirmed: new Set(),
    };
  }

  /**
   * 外部ドメインかどうかを判定
   */
  isExternalDomain = (email: string): boolean => {
    const { settings } = this.props;
    const domain = email.split('@')[1]?.toLowerCase();
    if (!domain) return false;

    return !settings.recipientCheck.internalDomains.some(
      (internalDomain) => domain === internalDomain.toLowerCase()
    );
  };

  /**
   * すべての外部宛先を取得（タイプ付き）
   */
  getExternalRecipients = (): Array<{ email: string; type: 'TO' | 'CC' | 'BCC' }> => {
    const { mailData } = this.props;
    const recipients: Array<{ email: string; type: 'TO' | 'CC' | 'BCC' }> = [];

    mailData.to.forEach(email => {
      if (this.isExternalDomain(email)) {
        recipients.push({ email, type: 'TO' });
      }
    });

    mailData.cc.forEach(email => {
      if (this.isExternalDomain(email)) {
        recipients.push({ email, type: 'CC' });
      }
    });

    mailData.bcc.forEach(email => {
      if (this.isExternalDomain(email)) {
        recipients.push({ email, type: 'BCC' });
      }
    });

    return recipients;
  };

  /**
   * すべての確認項目がチェックされているかどうか
   */
  canSend = (): boolean => {
    const externalRecipients = this.getExternalRecipients();
    const { mailData } = this.props;
    const { externalRecipientsConfirmed, subjectConfirmed, bodyConfirmed, attachmentsConfirmed } = this.state;

    // 外部宛先がある場合、すべて確認されているか
    const allExternalConfirmed = externalRecipients.every((recipient) =>
      externalRecipientsConfirmed.has(recipient.email)
    );

    // 件名・本文の確認
    let needsSubjectConfirm = mailData.subject && mailData.subject.trim() !== '';
    let needsBodyConfirm = mailData.body && mailData.body.trim() !== '';

    // 添付ファイルの確認（すべての添付ファイルがチェックされているか）
    let needsAttachmentsConfirm = mailData.attachments && mailData.attachments.length > 0;
    const allAttachmentsConfirmed = !needsAttachmentsConfirm ||
      mailData.attachments.every((attachment) => attachmentsConfirmed.has(attachment.name));

    return (
      allExternalConfirmed &&
      (!needsSubjectConfirm || subjectConfirmed) &&
      (!needsBodyConfirm || bodyConfirmed) &&
      allAttachmentsConfirmed
    );
  };

  /**
   * 外部宛先のチェックボックスを切り替え
   */
  toggleExternalRecipient = (email: string) => {
    this.setState((prevState) => {
      const newSet = new Set(prevState.externalRecipientsConfirmed);
      if (newSet.has(email)) {
        newSet.delete(email);
      } else {
        newSet.add(email);
      }
      return { externalRecipientsConfirmed: newSet };
    });
  };

  /**
   * 添付ファイルのチェックボックスを切り替え
   */
  toggleAttachment = (attachmentName: string) => {
    this.setState((prevState) => {
      const newSet = new Set(prevState.attachmentsConfirmed);
      if (newSet.has(attachmentName)) {
        newSet.delete(attachmentName);
      } else {
        newSet.add(attachmentName);
      }
      return { attachmentsConfirmed: newSet };
    });
  };

  /**
   * HTMLタグを除去してテキストのみ取得
   */
  stripHtml = (html: string): string => {
    const tmp = document.createElement('DIV');
    tmp.innerHTML = html;
    return tmp.textContent || tmp.innerText || '';
  };

  /**
   * 本文のプレビュー（最初の200文字）
   */
  getBodyPreview = (): string => {
    const { mailData } = this.props;
    const plainText = this.stripHtml(mailData.body);
    return plainText.length > 200 ? plainText.substring(0, 200) + '...' : plainText;
  };

  render() {
    const { mailData, checkResults, onCancel } = this.props;
    const { externalRecipientsConfirmed, subjectConfirmed, bodyConfirmed, attachmentsConfirmed } = this.state;
    const externalRecipients = this.getExternalRecipients();

    return (
      <div className="confirmation-panel">
        <h2>送信前確認</h2>

        {/* エラーと警告の表示 */}
        {checkResults.length > 0 && (
          <div className="check-results">
            {checkResults.map((result, index) => (
              <div key={index} className={`check-result ${result.severity}`}>
                <strong>{result.message}</strong>
                {result.details && <p>{result.details}</p>}
              </div>
            ))}
          </div>
        )}

        {/* 外部宛先の確認 */}
        {externalRecipients.length > 0 && (
          <div className="section external-recipients">
            <h3>⚠️ 外部ドメイン宛先の確認</h3>
            <p>以下の外部アドレスに送信します。確認してチェックしてください：</p>
            <div className="recipient-list horizontal-list">
              {externalRecipients.map((recipient) => (
                <label key={recipient.email} className="checkbox-item external-recipient">
                  <input
                    type="checkbox"
                    checked={externalRecipientsConfirmed.has(recipient.email)}
                    onChange={() => this.toggleExternalRecipient(recipient.email)}
                  />
                  <span>({recipient.type}) {recipient.email}</span>
                </label>
              ))}
            </div>
          </div>
        )}

        {/* 件名の確認 */}
        {mailData.subject && mailData.subject.trim() !== '' && (
          <div className="section subject-section">
            <h3>📧 件名</h3>
            <div className="horizontal-layout">
              <label className="checkbox-item inline-checkbox">
                <input
                  type="checkbox"
                  checked={subjectConfirmed}
                  onChange={(e) => this.setState({ subjectConfirmed: e.target.checked })}
                />
                <span>件名を確認しました</span>
              </label>
              <div className="subject-box">{mailData.subject}</div>
            </div>
          </div>
        )}

        {/* 本文の確認 */}
        {mailData.body && mailData.body.trim() !== '' && (
          <div className="section body-section">
            <h3>📝 本文（プレビュー）</h3>
            <div className="horizontal-layout">
              <label className="checkbox-item inline-checkbox">
                <input
                  type="checkbox"
                  checked={bodyConfirmed}
                  onChange={(e) => this.setState({ bodyConfirmed: e.target.checked })}
                />
                <span>本文を確認しました</span>
              </label>
              <div className="body-preview">{this.getBodyPreview()}</div>
            </div>
          </div>
        )}

        {/* 添付ファイルの確認 */}
        {mailData.attachments && mailData.attachments.length > 0 && (
          <div className="section attachments-section">
            <h3>📎 添付ファイル</h3>
            <p>以下の添付ファイルを確認してチェックしてください：</p>
            <div className="attachment-list horizontal-list">
              {mailData.attachments.map((attachment, index) => (
                <label key={index} className="checkbox-item attachment-item">
                  <input
                    type="checkbox"
                    checked={attachmentsConfirmed.has(attachment.name)}
                    onChange={() => this.toggleAttachment(attachment.name)}
                  />
                  <span>{attachment.name} ({Math.round(attachment.size / 1024)} KB)</span>
                </label>
              ))}
            </div>
          </div>
        )}

        {/* ボタン */}
        <div className="footer">
          <button
            className="send-button"
            onClick={this.props.onSend}
            disabled={!this.canSend()}
            title={!this.canSend() ? 'すべての項目を確認してください' : ''}
          >
            送信する
          </button>
          <button className="cancel-button" onClick={onCancel}>
            編集に戻る
          </button>
        </div>
      </div>
    );
  }
}
