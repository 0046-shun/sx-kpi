import { ExcelProcessor } from './excel-processor';
import { CalendarManager } from './calendar-manager';
export interface BIChartData {
    labels: string[];
    datasets: {
        label: string;
        data: number[];
        backgroundColor?: string | string[];
        borderColor?: string | string[];
        borderWidth?: number;
    }[];
}
export interface BIStats {
    totalOrders: number;
    totalAmount: number;
    averageAmount: number;
    regionOrders: {
        region: string;
        count: number;
    }[];
    regionAmounts: {
        region: string;
        amount: number;
    }[];
    dailyTrend: {
        day: string;
        count: number;
        dayType?: string;
        dayLabel?: string;
        dayColor?: string;
    }[];
    staffOrderRanking: {
        name: string;
        count: number;
        regionNo: string;
        departmentNo: string;
        regionName: string;
    }[];
    staffAmountRanking: {
        name: string;
        amount: number;
        regionNo: string;
        departmentNo: string;
        regionName: string;
    }[];
}
export declare class BITool {
    private excelProcessor;
    private calendarManager;
    private charts;
    constructor(excelProcessor: ExcelProcessor, calendarManager: CalendarManager);
    private parseAmount;
    private normalizeStaffName;
    private getRegionName;
    private formatDateWithDayOfWeek;
    private getHolidaySettings;
    private getDayType;
    createBIDashboard(data: any[], targetMonth?: string): string;
    private calculateStats;
    renderCharts(data: any[], targetMonth?: string): void;
    private renderRegionOrdersChart;
    private renderRegionAmountsChart;
    private renderDailyTrendChart;
    destroyCharts(): void;
}
