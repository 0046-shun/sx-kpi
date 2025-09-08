import { HolidaySettings, StaffData } from './types.js';
export declare class ReportGenerator {
    private currentTargetDate;
    private excelProcessor;
    private calendarManager;
    constructor(excelProcessor: any, calendarManager: any);
    updateHolidaySettings(settings: HolidaySettings): void;
    private getRegionName;
    private isHolidayConstruction;
    private isProhibitedConstruction;
    private isPublicHoliday;
    private isProhibitedDay;
    generateDailyReport(data: any[], date: string): any;
    generateMonthlyReport(data: any[], month: string): any;
    private getOvertimeCount;
    private isOvertime;
    private calculateReportData;
    private calculateRegionStats;
    private calculateAgeStats;
    private isSameDate;
    private getTotalExcessive;
    private getTotalSingle;
    private getTotalHolidayConstruction;
    private getTotalProhibitedConstruction;
    private isOrderForDate;
    private calculateElderlyStaffRanking;
    private calculateSingleContractRanking;
    private calculateExcessiveSalesRanking;
    private calculateNormalAgeStaffRanking;
    private assignRanks;
    createDailyReportHTML(report: any): string;
    createMonthlyReportHTML(report: any): string;
    private createRegionStatsHTML;
    private createAgeStatsHTML;
    exportToPDF(report: any, type: string): Promise<void>;
    private exportMonthlyReportToPDF;
    private createPDFHTML;
    private createPDFHTMLAsync;
    exportToCSV(report: any, type: string): Promise<void>;
    private exportStaffRankingCSVs;
    private downloadCSV;
    private createRankingTableHTML;
    generateStaffData(data: any[], targetDate: Date): StaffData[];
    private isSimpleOrder;
    private getContractorAge;
    private normalizeStaffName;
    exportStaffDataToCSV(staffData: StaffData[]): Promise<void>;
    createStaffDataHTML(staffData: StaffData[]): string;
    /**
     * 受注カウント日付を計算（A列とK列を比較して遅い日付を採用）
     * @param row データ行
     * @returns 受注カウントに使用される日付
     */
    private calculateEffectiveDate;
    createDataConfirmationHTML(data: any[]): string;
    private createDailyReportPDFHTML;
    private createRegionStatsPDFHTML;
    private createRegionCardPDFHTML;
    private createAgeStatsPDFHTML;
    createMonthlyReportPDFHTML(report: any): string;
    private createRankingTablePDFHTML;
}
