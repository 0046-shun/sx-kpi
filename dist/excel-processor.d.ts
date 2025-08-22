export declare class ExcelProcessor {
    readExcelFile(file: File): Promise<any[]>;
    private cleanAndValidateData;
    private parseDateFromKColumn;
    private parseDate;
    private parseTime;
    private parseAge;
    private isOrder;
    private isSingle;
    private isExcessive;
    private isOvertime;
    private isElderly;
    private getCurrentDateFilter;
    private getRegionName;
}
