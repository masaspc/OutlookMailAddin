import { CheckResult, MailData, SubjectCheckSettings } from '../types';
import { KeywordMatcher } from '../utils/KeywordMatcher';

/**
 * 件名チェッカー
 */
export class SubjectChecker {
  /**
   * 件名に関するチェックを実行する
   * @param mailData メールデータ
   * @param settings チェック設定
   * @returns チェック結果の配列
   */
  static check(mailData: MailData, settings: SubjectCheckSettings): CheckResult[] {
    if (!settings.enabled) {
      return [];
    }

    const results: CheckResult[] = [];

    // 件名未入力チェック
    if (settings.requireSubject) {
      const subject = mailData.subject?.trim() || '';

      if (subject === '') {
        results.push({
          severity: 'warning',
          category: 'subject',
          message: '件名が入力されていません',
          details: 'メールの件名を入力してください。件名があるとメールの内容が分かりやすくなります。',
        });
      }
    }

    // 件名に「Re:」がないのに本文が返信形式の場合のチェック
    const subject = mailData.subject || '';
    const bodyText = KeywordMatcher.stripHtml(mailData.body);

    // 返信を示すパターン
    const replyPatterns = [
      /^>.*$/m, // 引用行
      /^On .+ wrote:$/m, // 英語の返信パターン
      /^\d{4}年\d{1,2}月\d{1,2}日.*wrote:/m, // 日本語の返信パターン
      /^From:.*\nSent:.*\nTo:/m, // Outlookの転送パターン
    ];

    const looksLikeReply = replyPatterns.some((pattern) => pattern.test(bodyText));
    const hasRePrefix = /^(Re:|RE:|返信:)/i.test(subject);

    if (looksLikeReply && !hasRePrefix) {
      results.push({
        severity: 'info',
        category: 'subject',
        message: '返信メールのようですが、件名にRe:がありません',
        details: 'メール本文が返信形式のようですが、件名に「Re:」が含まれていません。新規メールとして送信しますか?',
      });
    }

    return results;
  }
}
