declare global {
    interface Window {
        XLSX: any;
        jspdf: any;
        saveAs: any;
    }
    const XLSX: any;
    const jsPDF: any;
}
export interface HolidaySettings {
    publicHolidays: Date[];
    prohibitedDays: Date[];
}
export interface OrderData {
    rowNumber: number;
    date: Date | null;
    time: string | null;
    regionNumber: any;
    departmentNumber: any;
    staffName: string;
    contractor: string;
    contractorAge: number | null;
    contractorRelation: string;
    contractorTel: string;
    confirmation: string;
    confirmationDateTime: string;
    confirmer: string;
    confirmerAge: number | null;
    confirmerRelation: string;
    confirmerTel: string;
    productName: string;
    quantity: any;
    amount: any;
    contractDate: Date | null;
    startDate: Date | null;
    startTime: string;
    completionDate: Date | null;
    paymentMethod: string;
    receptionist: string;
    coFlyer: any;
    designEstimateNumber: string;
    remarks: string;
    otherCompany: any;
    history: any;
    mainContract: any;
    total: any;
    isActionCompleted: boolean;
    isOrder: boolean;
    isOvertime: boolean;
    isElderly: boolean;
    isExcessive: boolean;
    isSingle: boolean;
    region: string;
    isHolidayConstruction: boolean;
    isProhibitedConstruction: boolean;
}
export interface ReportData {
    type: 'daily' | 'monthly';
    totalOrders: number;
    overtimeOrders: number;
    regionStats: RegionStats;
    ageStats: AgeStats;
    rawData: OrderData[];
}
export interface RegionStats {
    [key: string]: {
        orders: number;
        overtime: number;
        excessive: number;
        single: number;
        holidayConstruction: number;
        prohibitedConstruction: number;
    };
}
export interface AgeStats {
    elderly: {
        total: number;
        excessive: number;
        single: number;
    };
    normal: {
        total: number;
        excessive: number;
    };
}
export interface DataStatistics {
    totalRecords: number;
    dateRange: {
        start: Date | null;
        end: Date | null;
    };
    regionDistribution: {
        [key: string]: number;
    };
    ageDistribution: {
        elderly: number;
        normal: number;
    };
    overtimeDistribution: {
        [key: string]: number;
    };
    holidayConstructionDistribution: {
        [key: string]: number;
    };
    prohibitedConstructionDistribution: {
        [key: string]: number;
    };
}
export interface ReportOptions {
    includeHolidayConstruction: boolean;
    includeProhibitedConstruction: boolean;
    holidaySettings: HolidaySettings;
}
export interface CalendarSettings {
    year: number;
    month: number;
    holidays: HolidaySettings;
}
export interface FilterSettings {
    region?: string;
    age?: string;
    overtime?: string;
    dateFrom?: string;
    dateTo?: string;
}
export interface ExportSettings {
    format: 'pdf' | 'csv';
    includeRawData: boolean;
    includeStatistics: boolean;
    fileName?: string;
}
export interface StaffRanking {
    rank: number;
    regionNo: string;
    departmentNo: string;
    staffName: string;
    count: number;
}
export interface StaffData {
    regionNo: string;
    departmentNo: string;
    staffName: string;
    totalOrders: number;
    normalAgeOrders: number;
    elderlyOrders: number;
    corporateOrders: number;
    singleOrders: number;
    excessiveOrders: number;
    overtimeOrders: number;
}
export interface MonthlyReportData extends ReportData {
    selectedMonth: number;
    selectedYear: number;
    elderlyStaffRanking: StaffRanking[];
    singleContractRanking: StaffRanking[];
    excessiveSalesRanking: StaffRanking[];
    normalAgeStaffRanking: StaffRanking[];
}
