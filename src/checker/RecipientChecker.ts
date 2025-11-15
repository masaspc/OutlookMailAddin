import { CheckResult, MailData, RecipientCheckSettings } from '../types';
import { DomainChecker } from '../utils/DomainChecker';

/**
 * 宛先チェッカー
 */
export class RecipientChecker {
  /**
   * 宛先に関するチェックを実行する
   * @param mailData メールデータ
   * @param settings チェック設定
   * @returns チェック結果の配列
   */
  static check(mailData: MailData, settings: RecipientCheckSettings): CheckResult[] {
    if (!settings.enabled) {
      return [];
    }

    const results: CheckResult[] = [];

    // 全宛先を取得
    const allRecipients = [...mailData.to, ...mailData.cc, ...mailData.bcc];

    // 宛先未入力チェック
    if (allRecipients.length === 0) {
      results.push({
        severity: 'error',
        category: 'recipient',
        message: '宛先が入力されていません',
        details: '送信先のメールアドレスを入力してください。',
      });
      return results;
    }

    // 宛先件数チェック
    if (allRecipients.length > settings.maxRecipientCount) {
      results.push({
        severity: 'warning',
        category: 'recipient',
        message: `宛先が多すぎます（${allRecipients.length}件）`,
        details: `宛先が${settings.maxRecipientCount}件を超えています。本当にこれらすべての宛先に送信しますか?`,
      });
    }

    // 社外ドメインチェック
    if (settings.warnExternalRecipients && settings.internalDomains.length > 0) {
      const externalRecipients = DomainChecker.filterExternalEmails(allRecipients, settings.internalDomains);

      if (externalRecipients.length > 0) {
        results.push({
          severity: 'warning',
          category: 'recipient',
          message: '社外への送信が含まれています',
          details: `以下の社外アドレスへの送信が含まれています:\n${externalRecipients.join('\n')}`,
        });
      }
    }

    // 複数の社外宛先がある場合のBccチェック
    if (settings.warnMultipleExternal && settings.internalDomains.length > 0) {
      const externalToCC = DomainChecker.filterExternalEmails([...mailData.to, ...mailData.cc], settings.internalDomains);

      if (externalToCC.length > 1) {
        results.push({
          severity: 'warning',
          category: 'recipient',
          message: '複数の社外宛先がTo/Ccに含まれています',
          details: `複数の社外アドレス（${externalToCC.length}件）がTo/Ccに含まれています。個人情報保護のため、Bccの使用を検討してください。`,
        });
      }
    }

    return results;
  }
}
