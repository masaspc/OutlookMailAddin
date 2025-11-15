/**
 * チェック結果の重要度
 */
export type CheckSeverity = 'error' | 'warning' | 'info';

/**
 * チェックカテゴリ
 */
export type CheckCategory = 'attachment' | 'recipient' | 'subject' | 'body' | 'time';

/**
 * チェック結果
 */
export interface CheckResult {
  severity: CheckSeverity;
  category: CheckCategory;
  message: string;
  details?: string;
}

/**
 * 添付ファイルチェック設定
 */
export interface AttachmentCheckSettings {
  enabled: boolean;
  keywords: string[];        // 「添付」「別紙」など
  excludeKeywords: string[]; // 「添付不要」など
}

/**
 * 宛先チェック設定
 */
export interface RecipientCheckSettings {
  enabled: boolean;
  internalDomains: string[]; // 社内ドメインリスト
  warnExternalRecipients: boolean;
  warnMultipleExternal: boolean;
  maxRecipientCount: number;
}

/**
 * 件名チェック設定
 */
export interface SubjectCheckSettings {
  enabled: boolean;
  requireSubject: boolean;
}

/**
 * 本文チェック設定
 */
export interface BodyCheckSettings {
  enabled: boolean;
  requireBody: boolean;
  todoKeywords: string[];    // 「TODO」「あとで」など
}

/**
 * 時間帯チェック設定
 */
export interface TimeCheckSettings {
  enabled: boolean;
  warnLateNight: boolean;
  lateNightStart: string;    // "22:00"
  lateNightEnd: string;      // "06:00"
  warnWeekend: boolean;
}

/**
 * 全体設定
 */
export interface Settings {
  attachmentCheck: AttachmentCheckSettings;
  recipientCheck: RecipientCheckSettings;
  subjectCheck: SubjectCheckSettings;
  bodyCheck: BodyCheckSettings;
  timeCheck: TimeCheckSettings;
}

/**
 * デフォルト設定
 */
export const DEFAULT_SETTINGS: Settings = {
  attachmentCheck: {
    enabled: true,
    keywords: ['添付', '別紙', 'ファイル', '資料', '送付'],
    excludeKeywords: ['添付不要', '添付なし', '添付しません'],
  },
  recipientCheck: {
    enabled: true,
    internalDomains: ['tokyobaynet.co.jp'],
    warnExternalRecipients: true,
    warnMultipleExternal: true,
    maxRecipientCount: 50,
  },
  subjectCheck: {
    enabled: true,
    requireSubject: true,
  },
  bodyCheck: {
    enabled: true,
    requireBody: true,
    todoKeywords: ['TODO', 'あとで', 'FIXME', 'XXX', '要修正'],
  },
  timeCheck: {
    enabled: true,
    warnLateNight: true,
    lateNightStart: '22:00',
    lateNightEnd: '06:00',
    warnWeekend: true,
  },
};

/**
 * メールデータ
 */
export interface MailData {
  subject: string;
  body: string;
  to: string[];
  cc: string[];
  bcc: string[];
  attachments: Office.AttachmentDetails[];
}
