// グローバル型定義
declare global {
    interface Window {
        XLSX: any;
        jspdf: any;
        saveAs: any;
    }
    
    // グローバル変数の型定義
    const XLSX: any;
    const jsPDF: any;
}

// 受注データの型定義
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
    
    // 計算フィールド
    isActionCompleted: boolean;
    isOrder: boolean;
    isOvertime: boolean;
    isElderly: boolean;
    isExcessive: boolean;
    isSingle: boolean;
    region: string;
}

// レポートデータの型定義
export interface ReportData {
    type: 'daily' | 'monthly';
    totalOrders: number;
    overtimeOrders: number;
    regionStats: RegionStats;
    ageStats: AgeStats;
    rawData: OrderData[];
}

// 地区統計の型定義
export interface RegionStats {
    [key: string]: {
        orders: number;
        overtime: number;
        excessive: number;
        single: number;
    };
}

// 年齢統計の型定義
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

// データ統計の型定義
export interface DataStatistics {
    totalRecords: number;
    dateRange: {
        start: Date | null;
        end: Date | null;
    };
    regionDistribution: { [key: string]: number };
    ageDistribution: {
        elderly: number;
        normal: number;
    };
    overtimeRate: string;
}

// フィルタ設定の型定義
export interface FilterSettings {
    region?: string;
    age?: string;
    overtime?: string;
    dateFrom?: string;
    dateTo?: string;
}

// エクスポート設定の型定義
export interface ExportSettings {
    format: 'pdf' | 'csv';
    includeRawData: boolean;
    includeStatistics: boolean;
    fileName?: string;
}