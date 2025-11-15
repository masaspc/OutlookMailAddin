/**
 * キーワードマッチング用ユーティリティクラス
 */
export class KeywordMatcher {
  /**
   * テキストに指定されたキーワードのいずれかが含まれているかチェックする
   * @param text チェック対象のテキスト
   * @param keywords キーワードのリスト
   * @returns キーワードが見つかった場合true
   */
  static containsAny(text: string, keywords: string[]): boolean {
    if (!text || keywords.length === 0) return false;

    const lowerText = text.toLowerCase();
    return keywords.some((keyword) => {
      if (!keyword) return false;
      const lowerKeyword = keyword.toLowerCase();
      return lowerText.includes(lowerKeyword);
    });
  }

  /**
   * テキストに指定されたキーワードすべてが含まれているかチェックする
   * @param text チェック対象のテキスト
   * @param keywords キーワードのリスト
   * @returns すべてのキーワードが見つかった場合true
   */
  static containsAll(text: string, keywords: string[]): boolean {
    if (!text || keywords.length === 0) return false;

    const lowerText = text.toLowerCase();
    return keywords.every((keyword) => {
      if (!keyword) return true; // 空のキーワードは無視
      const lowerKeyword = keyword.toLowerCase();
      return lowerText.includes(lowerKeyword);
    });
  }

  /**
   * テキストから一致したキーワードを抽出する
   * @param text チェック対象のテキスト
   * @param keywords キーワードのリスト
   * @returns 一致したキーワードのリスト
   */
  static findMatches(text: string, keywords: string[]): string[] {
    if (!text || keywords.length === 0) return [];

    const lowerText = text.toLowerCase();
    return keywords.filter((keyword) => {
      if (!keyword) return false;
      const lowerKeyword = keyword.toLowerCase();
      return lowerText.includes(lowerKeyword);
    });
  }

  /**
   * 正規表現パターンにマッチするかチェックする
   * @param text チェック対象のテキスト
   * @param pattern 正規表現パターン
   * @returns マッチした場合true
   */
  static matchesPattern(text: string, pattern: RegExp): boolean {
    if (!text) return false;
    return pattern.test(text);
  }

  /**
   * HTMLタグを除去してテキストのみを抽出する
   * @param html HTMLテキスト
   * @returns プレーンテキスト
   */
  static stripHtml(html: string): string {
    if (!html) return '';

    // HTMLタグを除去
    let text = html.replace(/<[^>]*>/g, ' ');

    // HTMLエンティティをデコード
    const entities: { [key: string]: string } = {
      '&nbsp;': ' ',
      '&lt;': '<',
      '&gt;': '>',
      '&amp;': '&',
      '&quot;': '"',
      '&#39;': "'",
    };

    Object.keys(entities).forEach((entity) => {
      text = text.replace(new RegExp(entity, 'g'), entities[entity]);
    });

    // 連続する空白を1つにまとめる
    text = text.replace(/\s+/g, ' ').trim();

    return text;
  }
}
