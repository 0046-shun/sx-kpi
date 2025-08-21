export declare class ExcelProcessor {
    readExcelFile(file: File): Promise<any[]>;
    private cleanAndValidateData;
    private parseDateFromKColumn;
    private parseDate;
    private parseTime;
    private parseAge;
    private isOrder;
    private isOvertime;
    private isElderly;
    private isExcessive;
    private isSingle;
    private getCurrentDateFilter;
    private getRegionName;
}
