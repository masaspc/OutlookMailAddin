import { CheckResult, MailData, Settings } from '../types';
import { AttachmentChecker } from './AttachmentChecker';
import { RecipientChecker } from './RecipientChecker';
import { SubjectChecker } from './SubjectChecker';
import { BodyChecker } from './BodyChecker';
import { TimeChecker } from './TimeChecker';

/**
 * すべてのチェッカーを統合管理するクラス
 */
export class CheckerManager {
  /**
   * すべてのチェックを並列実行する
   * @param mailData メールデータ
   * @param settings チェック設定
   * @returns チェック結果の配列
   */
  static async checkAll(mailData: MailData, settings: Settings): Promise<CheckResult[]> {
    const results: CheckResult[] = [];

    try {
      // すべてのチェックを並列実行
      const [
        attachmentResults,
        recipientResults,
        subjectResults,
        bodyResults,
        timeResults,
      ] = await Promise.all([
        Promise.resolve(AttachmentChecker.check(mailData, settings.attachmentCheck)),
        Promise.resolve(RecipientChecker.check(mailData, settings.recipientCheck)),
        Promise.resolve(SubjectChecker.check(mailData, settings.subjectCheck)),
        Promise.resolve(BodyChecker.check(mailData, settings.bodyCheck)),
        Promise.resolve(TimeChecker.check(settings.timeCheck)),
      ]);

      // 結果を統合
      results.push(
        ...attachmentResults,
        ...recipientResults,
        ...subjectResults,
        ...bodyResults,
        ...timeResults
      );

      // 重要度順にソート（error > warning > info）
      results.sort((a, b) => {
        const severityOrder = { error: 0, warning: 1, info: 2 };
        return severityOrder[a.severity] - severityOrder[b.severity];
      });
    } catch (error) {
      console.error('チェック実行中にエラーが発生しました:', error);
      results.push({
        severity: 'error',
        category: 'body',
        message: 'チェック実行エラー',
        details: 'メールのチェック中にエラーが発生しました。もう一度お試しください。',
      });
    }

    return results;
  }

  /**
   * エラーレベルのチェック結果があるかどうかを判定する
   * @param results チェック結果の配列
   * @returns エラーレベルのチェック結果がある場合true
   */
  static hasErrors(results: CheckResult[]): boolean {
    return results.some((result) => result.severity === 'error');
  }

  /**
   * 警告レベルのチェック結果があるかどうかを判定する
   * @param results チェック結果の配列
   * @returns 警告レベルのチェック結果がある場合true
   */
  static hasWarnings(results: CheckResult[]): boolean {
    return results.some((result) => result.severity === 'warning');
  }

  /**
   * チェック結果のサマリーを取得する
   * @param results チェック結果の配列
   * @returns サマリー文字列
   */
  static getSummary(results: CheckResult[]): string {
    const errorCount = results.filter((r) => r.severity === 'error').length;
    const warningCount = results.filter((r) => r.severity === 'warning').length;
    const infoCount = results.filter((r) => r.severity === 'info').length;

    const parts: string[] = [];
    if (errorCount > 0) parts.push(`エラー: ${errorCount}件`);
    if (warningCount > 0) parts.push(`警告: ${warningCount}件`);
    if (infoCount > 0) parts.push(`情報: ${infoCount}件`);

    return parts.length > 0 ? parts.join(', ') : 'チェック完了（問題なし）';
  }
}
