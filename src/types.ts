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

// 公休日・禁止日設定の型定義
export interface HolidaySettings {
    publicHolidays: Date[];  // 公休日
    prohibitedDays: Date[];  // 禁止日
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
    isHolidayConstruction: boolean;  // 公休日施工
    isProhibitedConstruction: boolean;  // 禁止日施工
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
        holidayConstruction: number;  // 公休日施工
        prohibitedConstruction: number;  // 禁止日施工
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
    overtimeDistribution: { [key: string]: number };
    holidayConstructionDistribution: { [key: string]: number };  // 公休日施工分布
    prohibitedConstructionDistribution: { [key: string]: number };  // 禁止日施工分布
}

// レポート生成オプション
export interface ReportOptions {
    includeHolidayConstruction: boolean;
    includeProhibitedConstruction: boolean;
    holidaySettings: HolidaySettings;
}

// カレンダー設定の型定義
export interface CalendarSettings {
    year: number;
    month: number;
    holidays: HolidaySettings;
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

// 担当者別ランキングの型定義
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
    singleOrders: number;
    excessiveOrders: number;
    overtimeOrders: number;
}

// 月報データの型定義（拡張）
export interface MonthlyReportData extends ReportData {
    selectedMonth: number;
    selectedYear: number;
    // 担当者別ランキング
    elderlyStaffRanking: StaffRanking[];
    singleContractRanking: StaffRanking[];
    excessiveSalesRanking: StaffRanking[];
    normalAgeStaffRanking: StaffRanking[];
}