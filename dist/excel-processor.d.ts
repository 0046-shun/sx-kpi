export declare class ExcelProcessor {
    readExcelFile(file: File): Promise<any[]>;
    private cleanAndValidateData;
    private parseDateFromKColumn;
    private parseDate;
    private parseTime;
    private parseAge;
    /**
     * 担当者名を正規化する
     * 括弧内の情報（役職、担当範囲、技術分野など）を除いて基本の担当者名を抽出
     * 例：
     * - "山田(SE)" → "山田"
     * - "山田(岡田)" → "山田"
     * - "山田(技)" → "山田"
     * - "山田（技術）" → "山田"
     */
    private normalizeStaffName;
    /**
     * 指定された日付の受注かどうかを判定
     * @param row データ行
     * @param targetDate 対象日付
     * @param isDailyReport 日報生成時かどうか（true: 日報、false: 月報）
     */
    isOrderForDate(row: any, targetDate: Date, isDailyReport?: boolean): boolean;
    /**
     * 単独契約かどうかを判定
     */
    isSingle(row: any): boolean;
    /**
     * 過量販売かどうかを判定
     */
    isExcessive(row: any): boolean;
    /**
     * 時間外対応かどうかを判定
     * 18:30以降の対応を時間外とする
     */
    isOvertime(row: any): boolean;
    /**
     * 現在の日付フィルターを取得（動的）
     */
    getCurrentDateFilter(): Date;
}
