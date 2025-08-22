import { HolidaySettings } from './types.js';
export declare class DataManager {
    private static instance;
    private data;
    private holidaySettings;
    private constructor();
    static getInstance(): DataManager;
    setData(data: any[]): void;
    getData(): any[];
    getHolidaySettings(): HolidaySettings;
    setHolidaySettings(settings: HolidaySettings): void;
    addPublicHoliday(date: Date): void;
    removePublicHoliday(date: Date): void;
    addProhibitedDay(date: Date): void;
    removeProhibitedDay(date: Date): void;
    isPublicHoliday(date: Date): boolean;
    isProhibitedDay(date: Date): boolean;
    private formatDateForStorage;
    private saveHolidaySettings;
    private loadHolidaySettings;
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
