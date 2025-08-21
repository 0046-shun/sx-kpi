export declare class DataManager {
    private readonly STORAGE_KEY;
    saveData(data: any[]): void;
    loadData(): any[] | null;
    clearData(): void;
    createDataTableHTML(data: any[], filterState?: {
        region?: string;
        age?: string;
        overtime?: string;
    }): string;
    private filterAndSortData;
    private createPaginationHTML;
    private formatDate;
    private formatAmount;
    private createFilterStatusHTML;
    exportDataToCSV(data: any[]): Promise<void>;
    getDataStatistics(data: any[]): any;
}
