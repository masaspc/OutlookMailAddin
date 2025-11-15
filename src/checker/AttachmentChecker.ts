import { CheckResult, MailData, AttachmentCheckSettings } from '../types';
import { KeywordMatcher } from '../utils/KeywordMatcher';

/**
 * 添付ファイルチェッカー
 */
export class AttachmentChecker {
  /**
   * 添付ファイルに関するチェックを実行する
   * @param mailData メールデータ
   * @param settings チェック設定
   * @returns チェック結果の配列
   */
  static check(mailData: MailData, settings: AttachmentCheckSettings): CheckResult[] {
    if (!settings.enabled) {
      return [];
    }

    const results: CheckResult[] = [];

    // 本文から添付ファイル関連のキーワードをチェック
    const bodyText = KeywordMatcher.stripHtml(mailData.body);
    const subjectText = mailData.subject || '';
    const fullText = `${subjectText} ${bodyText}`;

    // 除外キーワードが含まれている場合はチェックをスキップ
    if (KeywordMatcher.containsAny(fullText, settings.excludeKeywords)) {
      return results;
    }

    // 添付キーワードが含まれているかチェック
    const hasAttachmentKeyword = KeywordMatcher.containsAny(fullText, settings.keywords);

    // 実際に添付ファイルがあるかチェック
    const hasAttachments = mailData.attachments && mailData.attachments.length > 0;

    // 添付キーワードがあるのに添付ファイルがない場合は警告
    if (hasAttachmentKeyword && !hasAttachments) {
      const matchedKeywords = KeywordMatcher.findMatches(fullText, settings.keywords);
      results.push({
        severity: 'warning',
        category: 'attachment',
        message: '添付ファイルが見つかりません',
        details: `メール本文に「${matchedKeywords.join('、')}」というキーワードが含まれていますが、添付ファイルがありません。添付ファイルを忘れていませんか?`,
      });
    }

    return results;
  }
}
