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
     * 受注としてカウントするかどうかを判定
     * 動的日付判定に対応
     */
    private isOrderForDate;
    /**
     * 単独契約かどうかを判定
     */
    isSingle(confirmation: any): boolean;
    /**
     * 過量販売かどうかを判定
     */
    isExcessive(confirmation: any): boolean;
    /**
     * 時間外対応かどうかを判定
     * 18:30以降の対応を時間外とする
     */
    isOvertime(date: any, time: any, confirmation: any, confirmationDateTime: any, age: any, targetDate?: Date): boolean;
    /**
     * 現在の日付フィルターを取得（動的）
     */
    getCurrentDateFilter(): Date;
}
