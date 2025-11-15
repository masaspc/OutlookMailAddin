import { CheckResult, TimeCheckSettings } from '../types';

/**
 * 時間帯チェッカー
 */
export class TimeChecker {
  /**
   * 送信時間に関するチェックを実行する
   * @param settings チェック設定
   * @returns チェック結果の配列
   */
  static check(settings: TimeCheckSettings): CheckResult[] {
    if (!settings.enabled) {
      return [];
    }

    const results: CheckResult[] = [];
    const now = new Date();

    // 深夜・早朝送信チェック
    if (settings.warnLateNight) {
      if (this.isLateNight(now, settings.lateNightStart, settings.lateNightEnd)) {
        results.push({
          severity: 'info',
          category: 'time',
          message: '深夜・早朝の時間帯です',
          details: `現在は${this.formatTime(now)}です。深夜・早朝のメール送信は受信者に負担をかける可能性があります。本当に今送信しますか?`,
        });
      }
    }

    // 休日送信チェック
    if (settings.warnWeekend) {
      if (this.isWeekend(now)) {
        const dayName = this.getDayName(now);
        results.push({
          severity: 'info',
          category: 'time',
          message: '休日です',
          details: `今日は${dayName}です。休日のメール送信は受信者に負担をかける可能性があります。本当に今送信しますか?`,
        });
      }
    }

    return results;
  }

  /**
   * 深夜・早朝かどうかを判定する
   * @param date 判定する日時
   * @param startTime 深夜開始時刻（例: "22:00"）
   * @param endTime 早朝終了時刻（例: "06:00"）
   * @returns 深夜・早朝の場合true
   */
  private static isLateNight(date: Date, startTime: string, endTime: string): boolean {
    const currentHour = date.getHours();
    const currentMinute = date.getMinutes();
    const currentTimeInMinutes = currentHour * 60 + currentMinute;

    const [startHour, startMinute] = startTime.split(':').map(Number);
    const startTimeInMinutes = startHour * 60 + startMinute;

    const [endHour, endMinute] = endTime.split(':').map(Number);
    const endTimeInMinutes = endHour * 60 + endMinute;

    // 深夜帯が日付をまたぐ場合（例: 22:00-06:00）
    if (startTimeInMinutes > endTimeInMinutes) {
      return currentTimeInMinutes >= startTimeInMinutes || currentTimeInMinutes < endTimeInMinutes;
    } else {
      // 深夜帯が日付をまたがない場合
      return currentTimeInMinutes >= startTimeInMinutes && currentTimeInMinutes < endTimeInMinutes;
    }
  }

  /**
   * 休日（土日）かどうかを判定する
   * @param date 判定する日時
   * @returns 休日の場合true
   */
  private static isWeekend(date: Date): boolean {
    const dayOfWeek = date.getDay();
    return dayOfWeek === 0 || dayOfWeek === 6; // 0: 日曜日, 6: 土曜日
  }

  /**
   * 曜日名を取得する
   * @param date 日時
   * @returns 曜日名
   */
  private static getDayName(date: Date): string {
    const days = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];
    return days[date.getDay()];
  }

  /**
   * 時刻を "HH:MM" 形式でフォーマットする
   * @param date 日時
   * @returns フォーマットされた時刻
   */
  private static formatTime(date: Date): string {
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
  }
}
