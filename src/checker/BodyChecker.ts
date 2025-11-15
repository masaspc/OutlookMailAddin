import { CheckResult, MailData, BodyCheckSettings } from '../types';
import { KeywordMatcher } from '../utils/KeywordMatcher';

/**
 * 本文チェッカー
 */
export class BodyChecker {
  /**
   * 本文に関するチェックを実行する
   * @param mailData メールデータ
   * @param settings チェック設定
   * @returns チェック結果の配列
   */
  static check(mailData: MailData, settings: BodyCheckSettings): CheckResult[] {
    if (!settings.enabled) {
      return [];
    }

    const results: CheckResult[] = [];
    const bodyText = KeywordMatcher.stripHtml(mailData.body);

    // 本文未入力チェック
    if (settings.requireBody) {
      const trimmedBody = bodyText.trim();

      // 署名を除いた本文が空の場合
      // 簡易的に、署名の開始を示す一般的なパターンを除外
      const bodyWithoutSignature = this.removeSignature(trimmedBody);

      if (bodyWithoutSignature === '') {
        results.push({
          severity: 'warning',
          category: 'body',
          message: '本文が入力されていません',
          details: 'メールの本文を入力してください。',
        });
      }
    }

    // TODOマーカーチェック
    if (settings.todoKeywords.length > 0) {
      const matchedTodos = KeywordMatcher.findMatches(bodyText, settings.todoKeywords);

      if (matchedTodos.length > 0) {
        results.push({
          severity: 'warning',
          category: 'body',
          message: 'TODOマーカーが含まれています',
          details: `本文に「${matchedTodos.join('、')}」が含まれています。編集が完了していない可能性があります。`,
        });
      }
    }

    return results;
  }

  /**
   * 本文から署名を除去する（簡易版）
   * @param body 本文
   * @returns 署名を除去した本文
   */
  private static removeSignature(body: string): string {
    // 一般的な署名の開始パターン
    const signaturePatterns = [
      /^--\s*$/m, // "--" で始まる行
      /^-{2,}\s*$/m, // "-----" などの区切り線
      /^よろしくお願い/m,
      /^Best regards/im,
      /^Regards/im,
      /^Sincerely/im,
    ];

    let bodyWithoutSig = body;

    for (const pattern of signaturePatterns) {
      const match = bodyWithoutSig.match(pattern);
      if (match && match.index !== undefined) {
        bodyWithoutSig = bodyWithoutSig.substring(0, match.index);
        break;
      }
    }

    return bodyWithoutSig.trim();
  }
}
