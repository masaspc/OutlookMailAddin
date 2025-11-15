/**
 * ドメインチェック用ユーティリティクラス
 */
export class DomainChecker {
  /**
   * メールアドレスからドメインを抽出する
   * @param email メールアドレス
   * @returns ドメイン（抽出できない場合は空文字列）
   */
  static extractDomain(email: string): string {
    const match = email.match(/@(.+)$/);
    return match ? match[1].toLowerCase() : '';
  }

  /**
   * メールアドレスが社内ドメインかどうかをチェックする
   * @param email メールアドレス
   * @param internalDomains 社内ドメインのリスト
   * @returns 社内ドメインの場合true
   */
  static isInternalDomain(email: string, internalDomains: string[]): boolean {
    const domain = this.extractDomain(email);
    if (!domain) return false;

    return internalDomains.some((internalDomain) => {
      const normalizedInternal = internalDomain.toLowerCase();
      return domain === normalizedInternal || domain.endsWith(`.${normalizedInternal}`);
    });
  }

  /**
   * メールアドレスが外部ドメインかどうかをチェックする
   * @param email メールアドレス
   * @param internalDomains 社内ドメインのリスト
   * @returns 外部ドメインの場合true
   */
  static isExternalDomain(email: string, internalDomains: string[]): boolean {
    if (!email) return false;
    return !this.isInternalDomain(email, internalDomains);
  }

  /**
   * メールアドレスのリストから外部ドメインのアドレスを抽出する
   * @param emails メールアドレスのリスト
   * @param internalDomains 社内ドメインのリスト
   * @returns 外部ドメインのメールアドレスのリスト
   */
  static filterExternalEmails(emails: string[], internalDomains: string[]): string[] {
    return emails.filter((email) => this.isExternalDomain(email, internalDomains));
  }

  /**
   * メールアドレスのリストから社内ドメインのアドレスを抽出する
   * @param emails メールアドレスのリスト
   * @param internalDomains 社内ドメインのリスト
   * @returns 社内ドメインのメールアドレスのリスト
   */
  static filterInternalEmails(emails: string[], internalDomains: string[]): string[] {
    return emails.filter((email) => this.isInternalDomain(email, internalDomains));
  }
}
