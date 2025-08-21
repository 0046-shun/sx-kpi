declare global {
    interface Window {
        XLSX: any;
        jspdf: any;
        saveAs: any;
    }
    const XLSX: any;
    const jsPDF: any;
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
    overtimeRate: string;
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
