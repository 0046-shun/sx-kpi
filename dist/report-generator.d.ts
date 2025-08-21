export declare class ReportGenerator {
    generateDailyReport(data: any[], date: string): any;
    generateMonthlyReport(data: any[], month: string): any;
    private calculateReportData;
    private calculateRegionStats;
    private calculateAgeStats;
    private isSameDate;
    createDailyReportHTML(report: any): string;
    createMonthlyReportHTML(report: any): string;
    private createRegionStatsHTML;
    private createAgeStatsHTML;
    exportToPDF(report: any, type: string): Promise<void>;
    exportToCSV(report: any, type: string): Promise<void>;
}
