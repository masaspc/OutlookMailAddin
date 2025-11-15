import { Settings, DEFAULT_SETTINGS } from '../types';

/**
 * 設定の保存・読み込みを管理するクラス
 * Office.jsのRoamingSettingsを使用
 */
export class SettingsStorage {
  private static readonly SETTINGS_KEY = 'outlookSendGuardSettings';

  /**
   * 設定を保存する
   * @param settings 保存する設定
   */
  static async saveSettings(settings: Settings): Promise<void> {
    return new Promise((resolve, reject) => {
      try {
        Office.context.roamingSettings.set(this.SETTINGS_KEY, JSON.stringify(settings));
        Office.context.roamingSettings.saveAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(asyncResult.error?.message || '設定の保存に失敗しました'));
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * 設定を読み込む
   * @returns 設定オブジェクト（保存されていない場合はデフォルト設定）
   */
  static loadSettings(): Settings {
    try {
      const settingsJson = Office.context.roamingSettings.get(this.SETTINGS_KEY) as string;
      if (settingsJson) {
        const settings = JSON.parse(settingsJson) as Settings;
        // デフォルト設定とマージ（新しい設定項目が追加された場合に対応）
        return this.mergeWithDefaults(settings);
      }
    } catch (error) {
      console.error('設定の読み込みに失敗しました:', error);
    }
    return DEFAULT_SETTINGS;
  }

  /**
   * デフォルト設定とマージする
   * @param settings 保存された設定
   * @returns マージされた設定
   */
  private static mergeWithDefaults(settings: Settings): Settings {
    return {
      attachmentCheck: {
        ...DEFAULT_SETTINGS.attachmentCheck,
        ...settings.attachmentCheck,
      },
      recipientCheck: {
        ...DEFAULT_SETTINGS.recipientCheck,
        ...settings.recipientCheck,
      },
      subjectCheck: {
        ...DEFAULT_SETTINGS.subjectCheck,
        ...settings.subjectCheck,
      },
      bodyCheck: {
        ...DEFAULT_SETTINGS.bodyCheck,
        ...settings.bodyCheck,
      },
      timeCheck: {
        ...DEFAULT_SETTINGS.timeCheck,
        ...settings.timeCheck,
      },
    };
  }

  /**
   * 設定をリセットする（デフォルトに戻す）
   */
  static async resetSettings(): Promise<void> {
    return this.saveSettings(DEFAULT_SETTINGS);
  }

  /**
   * 特定の設定項目を更新する
   * @param key 設定のキー
   * @param value 設定の値
   */
  static async updateSetting<K extends keyof Settings>(key: K, value: Settings[K]): Promise<void> {
    const currentSettings = this.loadSettings();
    currentSettings[key] = value;
    return this.saveSettings(currentSettings);
  }
}
