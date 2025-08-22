import { HolidaySettings } from './types.js';
export declare class ReportGenerator {
    private currentTargetDate;
    private holidaySettings;
    updateHolidaySettings(settings: HolidaySettings): void;
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
    createDailyReportHTML(report: any): string;
    createMonthlyReportHTML(report: any): string;
    private createRegionStatsHTML;
    private createAgeStatsHTML;
    exportToPDF(report: any, type: string): Promise<void>;
    private createPDFHTML;
    exportToCSV(report: any, type: string): Promise<void>;
}
