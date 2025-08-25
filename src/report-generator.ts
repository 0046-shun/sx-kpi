import { HolidaySettings, StaffData } from './types.js';

export class ReportGenerator {
    private currentTargetDate: Date | null = null;
    private excelProcessor: any;
    private calendarManager: any;

    constructor(excelProcessor: any, calendarManager: any) {
        this.excelProcessor = excelProcessor;
        this.calendarManager = calendarManager;
    }
    
    // 公休日・禁止日設定を更新
    updateHolidaySettings(settings: HolidaySettings): void {
        if (this.calendarManager) {
            this.calendarManager.updateSettings(settings);
        }
    }

    // 地区名を取得
    private getRegionName(regionNumber: any): string {
        if (!regionNumber) return 'その他';
        
        const number = parseInt(regionNumber);
        if (isNaN(number)) return 'その他';
        
        switch (number) {
            case 511:
                return '九州地区';
            case 521:
            case 531:
                return '中四国地区';
            case 541:
                return '関西地区';
            case 561:
                return '関東地区';
            default:
                return 'その他';
        }
    }


    
    // 公休日施工の判定
    private isHolidayConstruction(row: any): boolean {
        // T列（着工日）が公休日 → カウント
        if (row.startDate && this.isPublicHoliday(row.startDate)) {
            return true;
        }
        // T列でカウントされていない場合のみ、V列（完工予定日）を照合
        if (!row.startDate || !this.isPublicHoliday(row.startDate)) {
            if (row.completionDate && this.isPublicHoliday(row.completionDate)) {
                return true;
            }
        }
        return false;
    }
    
    // 禁止日施工の判定
    private isProhibitedConstruction(row: any): boolean {
        // T列（着工日）が禁止日 → カウント
        if (row.startDate && this.isProhibitedDay(row.startDate)) {
            return true;
        }
        // T列でカウントされていない場合のみ、V列（完工予定日）を照合
        if (!row.startDate || !this.isProhibitedDay(row.startDate)) {
            if (row.completionDate && this.isProhibitedDay(row.completionDate)) {
                return true;
            }
        }
        return false;
    }
    
    // 公休日かどうかの判定
    private isPublicHoliday(date: Date): boolean {
        if (this.calendarManager) {
            return this.calendarManager.isPublicHoliday(date);
        }
        return false;
    }
    
    // 禁止日かどうかの判定
    private isProhibitedDay(date: Date): boolean {
        if (this.calendarManager) {
            return this.calendarManager.isProhibitedDay(date);
        }
        return false;
    }
    
    generateDailyReport(data: any[], date: string): any {
        const targetDate = new Date(date);
        this.currentTargetDate = targetDate; // 時間外判定で使用
        const targetMonth = targetDate.getMonth();
        const targetDay = targetDate.getDate();
        
        console.log('日報生成開始 - 対象日:', date);
        console.log('対象月:', targetMonth + 1, '対象日:', targetDay);
        console.log('総データ件数:', data.length);
        

        
        // 受注日が対象日のデータを取得（公休日・禁止日施工判定用）
        const dailyData = data.filter(row => {
            // 日付チェック
            if (!row.date) {
                return false;
            }
            
                const rowMonth = row.date.getMonth();
                const rowDay = row.date.getDate();
            const targetMonth = targetDate.getMonth();
            const targetDay = targetDate.getDate();
            
            const isDateMatch = rowMonth === targetMonth && rowDay === targetDay;
            
            if (!isDateMatch) return false;
            
            // K列の文字列マッチングチェック
            if (row.confirmationDateTime) {
                const kColumnStr = String(row.confirmationDateTime);
                const targetDateStr = `${targetMonth + 1}/${targetDay}`;
                const targetDateStrAlt = `${targetMonth + 1}/${targetDay.toString().padStart(2, '0')}`;
                
                if (kColumnStr.includes(targetDateStr) || kColumnStr.includes(targetDateStrAlt)) {
                    return true;
                }
            }
            
            // K列の日付パターンチェック
            if (row.confirmationDateTime) {
                const kColumnStr = String(row.confirmationDateTime);
                const datePattern = kColumnStr.match(/(\d{1,2})\/(\d{1,2})/);
                if (datePattern) {
                    const kMonth = parseInt(datePattern[1]);
                    const kDay = parseInt(datePattern[2]);
                    if (kMonth === targetMonth + 1 && kDay === targetDay) {
                        return true;
                    }
                }
            }
            
            // J列条件チェック（1、2、5の場合は除外）
            let isJColumnValid = true;
            const confirmationValue = row.confirmation;
            
            if (typeof confirmationValue === 'number') {
                if (confirmationValue === 1 || confirmationValue === 2 || confirmationValue === 5) {
                    isJColumnValid = false;
                }
            } else if (typeof confirmationValue === 'string') {
                const trimmedValue = confirmationValue.trim();
                if (trimmedValue !== '') {
                    const confirmationNum = parseInt(trimmedValue);
                    if (!isNaN(confirmationNum) && (confirmationNum === 1 || confirmationNum === 2 || confirmationNum === 5)) {
                        isJColumnValid = false;
                    }
                }
            }
            
            // K列除外キーワードチェック
            let isKColumnValid = true;
            if (row.confirmationDateTime) {
                const kColumnStr = String(row.confirmationDateTime);
                
                // 除外キーワードのチェック
                if (kColumnStr.includes('担当待ち') || kColumnStr.includes('直電') || 
                    kColumnStr.includes('契約時') || kColumnStr.includes('待ち')) {
                    isKColumnValid = false;
                }
                
                // 有効な受注パターンのチェック
                if (kColumnStr.includes('単独契約') || kColumnStr.includes('過量販売')) {
                    isKColumnValid = true;
                }
            }
            
            const isValid = isDateMatch && isJColumnValid && isKColumnValid;
            
            return isValid;
        });
        
        console.log('日報対象データ件数:', dailyData.length);
        

        
        const reportData = this.calculateReportData(dailyData, 'daily');
        // 選択された日付情報を追加
        reportData.selectedDate = date;
        return reportData;
    }
    
    generateMonthlyReport(data: any[], month: string): any {
        const [year, monthNum] = month.split('-').map(Number);
        console.log('月報生成開始 - 対象年月:', year, monthNum);
        console.log('総データ件数:', data.length);
        
        const monthlyData = data.filter(row => {
            // 日付チェック
            if (!row.date) {
                return false;
            }
            
            const isDateMatch = row.date.getFullYear() === year && row.date.getMonth() === monthNum - 1;
            
            if (!isDateMatch) return false;
            
            // J列条件チェック（1、2、5の場合は除外）
            let isJColumnValid = true;
            const confirmationValue = row.confirmation;
            
            if (typeof confirmationValue === 'number') {
                if (confirmationValue === 1 || confirmationValue === 2 || confirmationValue === 5) {
                    isJColumnValid = false;
                }
            } else if (typeof confirmationValue === 'string') {
                const trimmedValue = confirmationValue.trim();
                if (trimmedValue !== '') {
                    const confirmationNum = parseInt(trimmedValue);
                    if (!isNaN(confirmationNum) && (confirmationNum === 1 || confirmationNum === 2 || confirmationNum === 5)) {
                        isJColumnValid = false;
                    }
                }
            }
            
            // K列除外キーワードチェック
            let isKColumnValid = true;
            if (row.confirmationDateTime) {
                const kColumnStr = String(row.confirmationDateTime);
                
                // 除外キーワードのチェック
                if (kColumnStr.includes('担当待ち') || kColumnStr.includes('直電') || 
                    kColumnStr.includes('契約時') || kColumnStr.includes('待ち')) {
                    isKColumnValid = false;
                }
                
                // 有効な受注パターンのチェック
                if (kColumnStr.includes('単独契約') || kColumnStr.includes('過量販売')) {
                    isKColumnValid = true;
                }
            }
            
            const isValid = isJColumnValid && isKColumnValid;
            
            return isValid;
        });
        
        const reportData = this.calculateReportData(monthlyData, 'monthly');
        reportData.rawData = monthlyData; // 月報データを保存
        
        // 選択された月情報を追加
        reportData.selectedMonth = month;
        reportData.selectedYear = year;
        
        // 担当者別ランキング集計（月報の対象期間で判定）
        const targetYear = year;
        const targetMonth = monthNum - 1; // JavaScript月は0ベース
        
        reportData.elderlyStaffRanking = this.calculateElderlyStaffRanking(monthlyData, targetYear, targetMonth);
        reportData.singleContractRanking = this.calculateSingleContractRanking(monthlyData, targetYear, targetMonth);
        reportData.excessiveSalesRanking = this.calculateExcessiveSalesRanking(monthlyData, targetYear, targetMonth);
        reportData.normalAgeStaffRanking = this.calculateNormalAgeStaffRanking(monthlyData, targetYear, targetMonth);
        
        return reportData;
    }
    
    // 時間外カウント取得（AB列とK列を独立してカウント）
    private getOvertimeCount(row: any, targetDate?: Date | null): number {
        // 月報の場合は、行の日付を使用
        let effectiveTargetDate: Date;
        if (!targetDate) {
            if (row.date && row.date instanceof Date) {
                effectiveTargetDate = row.date;
            } else {
                return 0;
            }
        } else {
            effectiveTargetDate = targetDate;
        }
        
        const targetMonth = effectiveTargetDate.getMonth();
        const targetDay = effectiveTargetDate.getDate();
        
        // 8/20のデータのみ詳細ログを出力
        const isTarget820 = row.date && row.date instanceof Date && 
                           row.date.getMonth() === 7 && row.date.getDate() === 20;
        
        let contractorOvertimeCount = 0; // ①のカウント
        let confirmerOvertimeCount = 0;  // ②のカウント
        
        // ①A列が対象日付であり、B列が18:30以降であるものをカウント
        if (row.date && row.date instanceof Date) {
            const aDateMonth = row.date.getMonth();
            const aDateDay = row.date.getDate();
            
            if (aDateMonth === targetMonth && aDateDay === targetDay) {
                // B列の時間チェック
                if (row.time !== null && row.time !== undefined) {
                    let hours = 0, minutes = 0, totalMinutes = 0;
                    
                    // 数値形式（Excel時間）の場合
                    if (typeof row.time === 'number') {
                        hours = Math.floor(row.time * 24);
                        minutes = Math.floor((row.time * 24 - hours) * 60);
                        totalMinutes = hours * 60 + minutes;
                    }
                    // 文字列形式（HH:MM）の場合
                    else if (typeof row.time === 'string' && row.time.includes(':')) {
                        const [h, m] = row.time.split(':').map(Number);
                        if (!isNaN(h) && !isNaN(m)) {
                            hours = h;
                            minutes = m;
                            totalMinutes = hours * 60 + minutes;
                        }
                    }
                    
                    if (isTarget820) {
                        // デバッグ情報（必要に応じて）
                    }
                    
                    if (totalMinutes >= 18 * 60 + 30) { // 18:30以降
                        contractorOvertimeCount = 1;

                    }
                }
            }
        }
        
        // ②K列の文字列を日付と時間と個人名に分解して、日付が対象日付であり、時間が18:30以降であるものをカウント
        if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
            const kColumnStr = String(row.confirmationDateTime);
            
            if (isTarget820) {

            }
            
            // ③K列が「同時」の場合は、AB列でカウント済なのでカウント無し
            if (kColumnStr === '同時') {
                // confirmerOvertimeCount = 0; // 明示的に0のまま
            }
            // K列に日付+時間が含まれる場合
            else if (kColumnStr.includes(':')) {
                // 日付+時間のパターンを検出（例: "8/19　18：44　柚木"）
                const timeMatch = kColumnStr.match(/(\d{1,2})\/(\d{1,2})\s*(\d{1,2})[：:]\s*(\d{1,2})/);
                if (timeMatch) {
                    const [_, month, day, hour, minute] = timeMatch;
                    
                    // 日付が対象日付と一致するかチェック
                    const kMonth = parseInt(month);
                    const kDay = parseInt(day);
                    if (kMonth === targetMonth + 1 && kDay === targetDay) {
                        // 時間が18:30以降かチェック
                        const kHour = parseInt(hour);
                        const kMinute = parseInt(minute);
                        const kTotalMinutes = kHour * 60 + kMinute;
                        
                        if (kTotalMinutes >= 18 * 60 + 30) {
                            confirmerOvertimeCount = 1;

                        }
                    }
                }
                // スペース区切りの従来の方法
                else {
                    const kColumnData = kColumnStr.split(' ');
                    if (kColumnData.length >= 2) {
                        const dateStr = kColumnData[0]; // 例: 8/20
                        const timeStr = kColumnData[1]; // 例: 19:37
                        
                        if (dateStr.includes('/')) {
                            const [month, day] = dateStr.split('/').map(Number);
                            
                            // 日付が対象日付と一致するかチェック
                            if (month === targetMonth + 1 && day === targetDay) {
                                // 時間が18:30以降かチェック
                                if (timeStr.includes(':')) {
                                    const [hours, minutes] = timeStr.split(':').map(Number);
                                    const totalMinutes = hours * 60 + minutes;
                                    
                                    if (totalMinutes >= 18 * 60 + 30) {
                                        confirmerOvertimeCount = 1;

                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        
        // ①＋②が時間外件数となる（独立してカウント）
        const totalCount = contractorOvertimeCount + confirmerOvertimeCount;
        
        if (isTarget820) {
            // デバッグ情報（必要に応じて）
        }
        
        if (totalCount > 0) {
            // 時間外判定結果（必要に応じて）
        }
        
        return totalCount;
    }
    
    // 後方互換性のための isOvertime メソッド（boolean返却）
    private isOvertime(row: any, targetDate?: Date | null): boolean {
        return this.getOvertimeCount(row, targetDate) > 0;
    }
    
    private calculateReportData(data: any[], type: string): any {
        // 基本統計
        const totalOrders = data.length;
        const overtimeOrders = data.reduce((total, row) => {
            return total + this.getOvertimeCount(row, type === 'daily' ? this.currentTargetDate : null);
        }, 0);
        
        // 地区別集計
        const regionStats = this.calculateRegionStats(data);
        
        // 高齢者・通常年齢の集計
        const ageStats = this.calculateAgeStats(data);
        
        return {
            type,
            totalOrders,
            overtimeOrders,
            regionStats,
            ageStats,
            rawData: data
        };
    }
    
    private calculateRegionStats(data: any[]): any {
        const regions: { [key: string]: { 
            orders: number; 
            overtime: number;
            excessive: number;
            single: number;
            holidayConstruction: number;
            prohibitedConstruction: number;
        } } = {
            '九州地区': { orders: 0, overtime: 0, excessive: 0, single: 0, holidayConstruction: 0, prohibitedConstruction: 0 },
            '中四国地区': { orders: 0, overtime: 0, excessive: 0, single: 0, holidayConstruction: 0, prohibitedConstruction: 0 },
            '関西地区': { orders: 0, overtime: 0, excessive: 0, single: 0, holidayConstruction: 0, prohibitedConstruction: 0 },
            '関東地区': { orders: 0, overtime: 0, excessive: 0, single: 0, holidayConstruction: 0, prohibitedConstruction: 0 },
            'その他': { orders: 0, overtime: 0, excessive: 0, single: 0, holidayConstruction: 0, prohibitedConstruction: 0 }
        };
        
        data.forEach(row => {
            const region = this.getRegionName(row.regionNumber);
            if (regions[region]) {
                regions[region].orders++;
                regions[region].overtime += this.getOvertimeCount(row, null);
                
                // 動的に判定
                if (this.excelProcessor.isExcessive(row)) {
                    regions[region].excessive++;
                }
                if (this.excelProcessor.isSingle(row)) {
                    regions[region].single++;
                }
                
                // 公休日・禁止日施工の判定と集計
                // T列（着工日）が公休日・禁止日 → カウント
                // T列でカウントされていない場合のみ、V列（完工予定日）を照合
                if (this.isHolidayConstruction(row)) {
                    regions[region].holidayConstruction++;
                }
                if (this.isProhibitedConstruction(row)) {
                    regions[region].prohibitedConstruction++;
                }
            }
        });
        
        return regions;
    }
    
    private calculateAgeStats(data: any[]): any {
        const stats = {
            elderly: {
                total: 0,
                excessive: 0,
                single: 0
            },
            normal: {
                total: 0,
                excessive: 0
            }
        };
        
        data.forEach(row => {
            const age = row.contractorAge || row.age;
            const isElderly = age && age >= 70;
            
            if (isElderly) {
                stats.elderly.total++;
                if (this.excelProcessor.isExcessive(row)) {
                    stats.elderly.excessive++;
                }
                if (this.excelProcessor.isSingle(row)) {
                    stats.elderly.single++;
                }
            } else {
                stats.normal.total++;
                if (this.excelProcessor.isExcessive(row)) {
                    stats.normal.excessive++;
                }
            }
        });
        
        return stats;
    }
    
    private isSameDate(date1: Date, date2: Date): boolean {
        return date1.getFullYear() === date2.getFullYear() &&
               date1.getMonth() === date2.getMonth() &&
               date1.getDate() === date2.getDate();
    }
    
    // 総過量販売件数を取得
    private getTotalExcessive(regionStats: any): number {
        return Object.values(regionStats).reduce((total: number, stats: any) => {
            return total + (stats.excessive || 0);
        }, 0);
    }
    
    // 総単独契約件数を取得
    private getTotalSingle(regionStats: any): number {
        return Object.values(regionStats).reduce((total: number, stats: any) => {
            return total + (stats.single || 0);
        }, 0);
    }
    
    // 総公休日施工件数を取得
    private getTotalHolidayConstruction(regionStats: any): number {
        return Object.values(regionStats).reduce((total: number, stats: any) => {
            return total + (stats.holidayConstruction || 0);
        }, 0);
    }
    
    // 総禁止日施工件数を取得
    private getTotalProhibitedConstruction(regionStats: any): number {
        return Object.values(regionStats).reduce((total: number, stats: any) => {
            return total + (stats.prohibitedConstruction || 0);
        }, 0);
    }

    // 受注判定メソッド（日付指定版）
    private isOrderForDate(row: any, targetDate: Date): boolean {
        return this.excelProcessor.isOrderForDate(row, targetDate);
    }

    // 担当者別ランキング集計メソッド

    // 契約者70歳以上の受注件数トップ10ランキング
    private calculateElderlyStaffRanking(data: any[], targetYear?: number, targetMonth?: number): any[] {

        const staffCounts = new Map<string, { regionNo: string; departmentNo: string; staffName: string; count: number }>();
        
        data.forEach((row, index) => {
            // 動的にisOrderを計算（行の日付を使用）
            const isOrder = this.isOrderForDate(row, row.date);
            const age = row.contractorAge || row.age;
            
            if (index < 10) {
                // デバッグ情報（必要に応じて）
            }
            
            // 条件: 70歳以上 AND 受注 AND 担当者名が存在
            if (age && age >= 70 && isOrder && row.staffName && row.staffName.trim() !== '') {
                const key = `${row.departmentNumber}_${row.staffName}`;
                const existing = staffCounts.get(key);
                if (existing) {
                    existing.count++;
                } else {
                    staffCounts.set(key, {
                        regionNo: row.regionNumber || '',
                        departmentNo: row.departmentNumber || '',
                        staffName: row.staffName,
                        count: 1
                    });
                }
            }
        });
        
        // 件数降順でソート
        const sorted = Array.from(staffCounts.values()).sort((a, b) => b.count - a.count);
        
        // トップ10を取得（同件数は同順位、次の順位は飛ばす）
        return this.assignRanks(sorted).slice(0, 10);
    }

    // 単独契約ランキング
    private calculateSingleContractRanking(data: any[], targetYear?: number, targetMonth?: number): any[] {

        const staffCounts = new Map<string, { regionNo: string; departmentNo: string; staffName: string; count: number }>();
        
        data.forEach((row, index) => {
            // 動的にisOrderを計算（行の日付を使用）
            const isOrder = this.isOrderForDate(row, row.date);
            
            if (index < 10) {

            }
            
            // 条件: 単独契約 AND 受注 AND 担当者名が存在
            if (this.excelProcessor.isSingle(row) && isOrder && row.staffName && row.staffName.trim() !== '') {
                const key = `${row.departmentNumber}_${row.staffName}`;
                const existing = staffCounts.get(key);
                if (existing) {
                    existing.count++;
                } else {
                    staffCounts.set(key, {
                        regionNo: row.regionNumber || '',
                        departmentNo: row.departmentNumber || '',
                        staffName: row.staffName,
                        count: 1
                    });
                }
            }
        });
        
        // 件数降順でソート
        const sorted = Array.from(staffCounts.values()).sort((a, b) => b.count - a.count);
        
        // 1件以上のみ
        const filtered = sorted.filter(item => item.count >= 1);
        return this.assignRanks(filtered);
    }

    // 過量販売ランキング
    private calculateExcessiveSalesRanking(data: any[], targetYear?: number, targetMonth?: number): any[] {

        const staffCounts = new Map<string, { regionNo: string; departmentNo: string; staffName: string; count: number }>();
        
        data.forEach((row, index) => {
            // 動的にisOrderを計算（行の日付を使用）
            const isOrder = this.isOrderForDate(row, row.date);
            
            if (index < 10) {

            }
            
            // 条件: 過量販売 AND 受注 AND 担当者名が存在
            if (this.excelProcessor.isExcessive(row) && isOrder && row.staffName && row.staffName.trim() !== '') {
                const key = `${row.departmentNumber}_${row.staffName}`;
                const existing = staffCounts.get(key);
                if (existing) {
                    existing.count++;
                } else {
                    staffCounts.set(key, {
                        regionNo: row.regionNumber || '',
                        departmentNo: row.departmentNumber || '',
                        staffName: row.staffName,
                        count: 1
                    });
                }
            }
        });
        
        // 件数降順でソート
        const sorted = Array.from(staffCounts.values()).sort((a, b) => b.count - a.count);
        
        // 1件以上のみ
        const filtered = sorted.filter(item => item.count >= 1);
        return this.assignRanks(filtered);
    }

    // 69歳以下契約件数の担当別件数
    private calculateNormalAgeStaffRanking(data: any[], targetYear?: number, targetMonth?: number): any[] {

        const staffCounts = new Map<string, { regionNo: string; departmentNo: string; staffName: string; count: number }>();
        
        data.forEach((row, index) => {
            // 動的にisOrderを計算（行の日付を使用）
            const isOrder = this.isOrderForDate(row, row.date);
            const age = row.contractorAge || row.age;
            const isNormalAge = !age || age < 70;
            
            if (index < 10) {
                // デバッグ情報（必要に応じて）
            }
            
            // 条件: 69歳以下 AND 受注 AND 担当者名が存在
            if (isNormalAge && isOrder && row.staffName && row.staffName.trim() !== '') {
                const key = `${row.departmentNumber}_${row.staffName}`;
                const existing = staffCounts.get(key);
                if (existing) {
                    existing.count++;
                } else {
                    staffCounts.set(key, {
                        regionNo: row.regionNumber || '',
                        departmentNo: row.departmentNumber || '',
                        staffName: row.staffName,
                        count: 1
                    });
                }
            }
        });
        
        // 件数降順でソート
        const sorted = Array.from(staffCounts.values()).sort((a, b) => b.count - a.count);
        
        // 全担当者（0件含む）
        return this.assignRanks(sorted);
    }

    // ランキング順位を割り当て（同件数は同順位、次の順位は飛ばす）
    private assignRanks(sorted: any[]): any[] {
        if (sorted.length === 0) return [];
        
        const result = [];
        let currentRank = 1;
        let currentCount = sorted[0].count;
        
        for (let i = 0; i < sorted.length; i++) {
            if (sorted[i].count < currentCount) {
                currentRank = i + 1;
                currentCount = sorted[i].count;
            }
            
            result.push({
                rank: currentRank,
                regionNo: sorted[i].regionNo,
                departmentNo: sorted[i].departmentNo,
                staffName: sorted[i].staffName,
                count: sorted[i].count
            });
        }
        
        return result;
    }
    
    createDailyReportHTML(report: any): string {
        // 選択された日付から日付テキストを取得
        const selectedDate = report.selectedDate ? new Date(report.selectedDate) : new Date();
        const dateText = selectedDate.toLocaleDateString('ja-JP');
        
        return `
            <div class="report-section fade-in">
                <h3 class="report-title">
                    <i class="fas fa-calendar-day me-2"></i>日報 - ${dateText}
                </h3>
                
                <!-- 基本統計 -->
                <div class="mb-4">
                    <div class="total-stats-container">
                        <h5 class="total-stats-title"><i class="fas fa-chart-bar me-2"></i>総件数</h5>
                        <div class="total-stats-grid">
                            <div class="total-stat-item">
                                <div class="total-stat-number">${report.totalOrders}</div>
                                <div class="total-stat-label">受注件数</div>
                        </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${report.overtimeOrders}</div>
                                <div class="total-stat-label">時間外対応</div>
                    </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalExcessive(report.regionStats)}</div>
                                <div class="total-stat-label">過量販売</div>
                        </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalSingle(report.regionStats)}</div>
                                <div class="total-stat-label">単独契約</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalHolidayConstruction(report.regionStats)}</div>
                                <div class="total-stat-label">公休日施工</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalProhibitedConstruction(report.regionStats)}</div>
                                <div class="total-stat-label">禁止日施工</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- 地区別受注件数 -->
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-map-marker-alt me-2"></i>地区別受注件数</h5>
                    ${this.createRegionStatsHTML(report.regionStats)}
                </div>
                
                <!-- 年齢別集計 -->
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-users me-2"></i>年齢別集計</h5>
                    ${this.createAgeStatsHTML(report.ageStats)}
                </div>
                
                <!-- エクスポートボタン -->
                <div class="text-center">
                    <button class="btn btn-success btn-export" data-format="pdf">
                        <i class="fas fa-file-pdf me-2"></i>PDF出力
                    </button>
                    <button class="btn btn-info btn-export" data-format="csv">
                        <i class="fas fa-file-csv me-2"></i>CSV出力
                    </button>
                </div>
            </div>
        `;
    }
    
    createMonthlyReportHTML(report: any): string {
        // 選択された月から年月を取得
        const selectedDate = report.selectedMonth ? new Date(report.selectedMonth + '-01') : new Date();
        const monthText = selectedDate.toLocaleDateString('ja-JP', { year: 'numeric', month: 'long' });
        
        return `
            <div class="report-section fade-in">
                <h3 class="report-title">
                    <i class="fas fa-calendar-alt me-2"></i>月報 - ${monthText}
                </h3>
                
                <!-- 基本統計 -->
                <div class="mb-4">
                    <div class="total-stats-container">
                        <h5 class="total-stats-title"><i class="fas fa-chart-bar me-2"></i>総件数</h5>
                        <div class="total-stats-grid">
                            <div class="total-stat-item">
                                <div class="total-stat-number">${report.totalOrders}</div>
                                <div class="total-stat-label">受注件数</div>
                        </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${report.overtimeOrders}</div>
                                <div class="total-stat-label">時間外対応</div>
                    </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalExcessive(report.regionStats)}</div>
                                <div class="total-stat-label">過量販売</div>
                        </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalSingle(report.regionStats)}</div>
                                <div class="total-stat-label">単独契約</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalHolidayConstruction(report.regionStats)}</div>
                                <div class="total-stat-label">公休日施工</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalProhibitedConstruction(report.regionStats)}</div>
                                <div class="total-stat-label">禁止日施工</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- 地区別受注件数 -->
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-map-marker-alt me-2"></i>地区別受注件数</h5>
                    ${this.createRegionStatsHTML(report.regionStats)}
                </div>
                
                <!-- 年齢別集計 -->
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-users me-2"></i>年齢別集計</h5>
                    ${this.createAgeStatsHTML(report.ageStats)}
                </div>

                <!-- 担当者別ランキング -->
                ${report.elderlyStaffRanking ? `
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-trophy me-2"></i>担当者別ランキング</h5>
                    
                    <!-- ①契約者70歳以上の受注件数トップ10ランキング -->
                    <div class="ranking-section mb-4">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <h6 class="ranking-title mb-0">①契約者70歳以上の受注件数トップ10ランキング</h6>
                            <button class="btn btn-outline-primary btn-sm" type="button" data-bs-toggle="collapse" data-bs-target="#elderlyRankingCollapse" aria-expanded="false" aria-controls="elderlyRankingCollapse">
                                <i class="fas fa-eye me-1"></i>表示/非表示
                            </button>
                        </div>
                        <div class="collapse" id="elderlyRankingCollapse">
                            ${this.createRankingTableHTML(report.elderlyStaffRanking)}
                        </div>
                    </div>

                    <!-- ②単独契約を持っている担当者一覧 -->
                    <div class="ranking-section mb-4">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <h6 class="ranking-title mb-0">②単独契約を持っている担当者一覧</h6>
                            <button class="btn btn-outline-primary btn-sm" type="button" data-bs-toggle="collapse" data-bs-target="#singleContractRankingCollapse" aria-expanded="false" aria-controls="singleContractRankingCollapse">
                                <i class="fas fa-eye me-1"></i>表示/非表示
                            </button>
                        </div>
                        <div class="collapse" id="singleContractRankingCollapse">
                            ${this.createRankingTableHTML(report.singleContractRanking)}
                        </div>
                    </div>

                    <!-- ③過量契約を持っている担当者一覧 -->
                    <div class="ranking-section mb-4">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <h6 class="ranking-title mb-0">③過量契約を持っている担当者一覧</h6>
                            <button class="btn btn-outline-primary btn-sm" type="button" data-bs-toggle="collapse" data-bs-target="#excessiveSalesRankingCollapse" aria-expanded="false" aria-controls="excessiveSalesRankingCollapse">
                                <i class="fas fa-eye me-1"></i>表示/非表示
                            </button>
                        </div>
                        <div class="collapse" id="excessiveSalesRankingCollapse">
                            ${this.createRankingTableHTML(report.excessiveSalesRanking)}
                        </div>
                    </div>

                    <!-- ④69歳以下契約件数の担当別件数 -->
                    <div class="ranking-section mb-4">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <h6 class="ranking-title mb-0">④69歳以下契約件数の担当別件数</h6>
                            <button class="btn btn-outline-primary btn-sm" type="button" data-bs-toggle="collapse" data-bs-target="#normalAgeRankingCollapse" aria-expanded="false" aria-controls="normalAgeRankingCollapse">
                                <i class="fas fa-eye me-1"></i>表示/非表示
                            </button>
                        </div>
                        <div class="collapse" id="normalAgeRankingCollapse">
                            ${this.createRankingTableHTML(report.normalAgeStaffRanking)}
                        </div>
                    </div>
                </div>
                ` : ''}
                
                <!-- エクスポートボタン -->
                <div class="text-center">
                    <button class="btn btn-success btn-export" data-format="pdf">
                        <i class="fas fa-file-pdf me-2"></i>PDF出力
                    </button>
                    <button class="btn btn-info btn-export" data-format="csv">
                        <i class="fas fa-file-csv me-2"></i>CSV出力
                    </button>
                </div>
            </div>
        `;
    }
    
    private createRegionStatsHTML(regionStats: any): string {
        let html = '<div class="row">';
        
        Object.entries(regionStats).forEach(([region, stats]: [string, any]) => {
            if (stats.orders > 0) {
                html += `
                    <div class="col-md-6 mb-3">
                        <div class="region-card">
                            <div class="region-title">${region}</div>
                            <div class="region-stats">
                                <div class="region-stat">
                                    <div class="region-stat-number">${stats.orders}</div>
                                    <div class="region-stat-label">受注件数</div>
                                </div>
                                <div class="region-stat">
                                    <div class="region-stat-number">${stats.overtime}</div>
                                    <div class="region-stat-label">時間外対応</div>
                                </div>
                                <div class="region-stat">
                                    <div class="region-stat-number">${stats.excessive}</div>
                                    <div class="region-stat-label">過量販売</div>
                                </div>
                                <div class="region-stat">
                                    <div class="region-stat-number">${stats.single}</div>
                                    <div class="region-stat-label">単独契約</div>
                                </div>
                                <div class="region-stat">
                                    <div class="region-stat-number">${stats.holidayConstruction}</div>
                                    <div class="region-stat-label">公休日施工</div>
                                </div>
                                <div class="region-stat">
                                    <div class="region-stat-number">${stats.prohibitedConstruction}</div>
                                    <div class="region-stat-label">禁止日施工</div>
                                </div>
                            </div>
                        </div>
                    </div>
                `;
            }
        });
        
        html += '</div>';
        return html;
    }
    
    private createAgeStatsHTML(ageStats: any): string {
        return `
            <div class="row">
                <div class="col-md-6 mb-3">
                    <div class="region-card">
                        <div class="region-title">契約者高齢者（70歳以上）</div>
                        <div class="region-stats">
                            <div class="region-stat">
                                <div class="region-stat-number">${ageStats.elderly.total}</div>
                                <div class="region-stat-label">総件数</div>
                            </div>
                            <div class="region-stat">
                                <div class="region-stat-number">${ageStats.elderly.excessive}</div>
                                <div class="region-stat-label">過量販売</div>
                            </div>
                            <div class="region-stat">
                                <div class="region-stat-number">${ageStats.elderly.single}</div>
                                <div class="region-stat-label">単独契約</div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-md-6 mb-3">
                    <div class="region-card">
                        <div class="region-title">契約者通常年齢（69歳以下）</div>
                        <div class="region-stats">
                            <div class="region-stat">
                                <div class="region-stat-number">${ageStats.normal.total}</div>
                                <div class="region-stat-label">総件数</div>
                            </div>
                            <div class="region-stat">
                                <div class="region-stat-number">${ageStats.normal.excessive}</div>
                                <div class="region-stat-label">過量販売</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }
    
    async exportToPDF(report: any, type: string): Promise<void> {
        try {
            // PDF用のHTML要素を動的に作成
            const pdfContainer = this.createPDFHTML(report, type);
            document.body.appendChild(pdfContainer);

            // jsPDFインスタンス（余白なし）
            const { jsPDF } = (window as any).jspdf;
            const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();
            const marginX = 14; // 左右余白（画像イメージに合わせて広め）
            const marginY = 12; // 上下余白（視覚的に落ち着く程度）

            // 対象セクションを取得（なければコンテナ全体）
            const sections = Array.from(pdfContainer.querySelectorAll('.report-section')) as HTMLElement[];
            const targets: HTMLElement[] = sections.length > 0 ? sections : [pdfContainer as HTMLElement];

            for (let i = 0; i < targets.length; i++) {
                const section = targets[i];
                // セクション単位でレンダリング（高解像度）
                const canvas = await (window as any).html2canvas(section, {
                    scale: 3,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff',
                    logging: false
                });

                const imgData = canvas.toDataURL('image/png');
                const imgWpx = canvas.width;
                const imgHpx = canvas.height;
                const imgAspect = imgWpx / imgHpx;

                // ページの利用可能領域（上下左右に適度な余白）
                const availW = pageWidth - marginX * 2;
                const availH = pageHeight - marginY * 2;

                // 高さ優先で拡大（1枚を使い切る）
                let drawHeight = availH;
                let drawWidth = drawHeight * imgAspect;
                // はみ出す場合のみ幅優先に切替
                if (drawWidth > availW) {
                    drawWidth = availW;
                    drawHeight = drawWidth / imgAspect;
                }
                // 左右は中央寄せ、縦は上寄せ（上marginY固定）
                const x = (pageWidth - drawWidth) / 2;
                const y = marginY;

                if (i > 0) {
                    pdf.addPage();
                }
                pdf.addImage(imgData, 'PNG', x, y, drawWidth, drawHeight, undefined, 'FAST');
            }

            // 一時的なHTML要素を削除
            document.body.removeChild(pdfContainer);

            // PDFをダウンロード
            const fileName = `${type === 'daily' ? '日報' : '月報'}_${new Date().toISOString().split('T')[0]}.pdf`;
            pdf.save(fileName);
        } catch (error) {
            console.error('PDF出力エラー:', error);
            alert('PDFの出力に失敗しました。');
        }
    }
    
    private createPDFHTML(report: any, type: string): HTMLElement {
        const container = document.createElement('div');
        container.className = 'pdf-container';
        
        // 新しく作成したPDF用メソッドを使用
        if (type === 'daily') {
            container.innerHTML = this.createDailyReportPDFHTML(report);
        } else if (type === 'monthly') {
            container.innerHTML = this.createMonthlyReportPDFHTML(report);
        } else {
            // フォールバック
            container.innerHTML = this.createDailyReportPDFHTML(report);
        }
        
        return container;
    }
    
    async exportToCSV(report: any, type: string): Promise<void> {
        try {
            if (type === 'monthly' && report.elderlyStaffRanking) {
                // 月報の場合は担当者別ランキングを個別CSV化
                await this.exportStaffRankingCSVs(report);
                return;
            }

            // 日報の場合は従来通りのCSV出力
            const csvData = [
                // ヘッダー行
                ['項目', '受注件数', '時間外対応', '過量販売', '単独契約', '公休日施工', '禁止日施工'],
                // 総件数行
                ['総件数', report.totalOrders.toString(), report.overtimeOrders.toString(), 
                 this.getTotalExcessive(report.regionStats).toString(), 
                 this.getTotalSingle(report.regionStats).toString(),
                 this.getTotalHolidayConstruction(report.regionStats).toString(),
                 this.getTotalProhibitedConstruction(report.regionStats).toString()],
                [''],
                // 地区別受注件数ヘッダー
                ['地区別受注件数', '受注件数', '時間外対応', '過量販売', '単独契約', '公休日施工', '禁止日施工'],
                // 各地区のデータ
                ...Object.entries(report.regionStats)
                    .filter(([_, stats]: [string, any]) => stats.orders > 0)
                    .map(([region, stats]: [string, any]) => [
                        region, 
                        stats.orders.toString(), 
                        stats.overtime.toString(), 
                        stats.excessive.toString(), 
                        stats.single.toString(),
                        stats.holidayConstruction.toString(),
                        stats.prohibitedConstruction.toString()
                    ]),
                [''],
                // 年齢別集計ヘッダー
                ['年齢別集計', '総件数', '過量販売', '単独契約', '', '', ''],
                // 高齢者データ
                ['高齢者（70歳以上）', report.ageStats.elderly.total.toString(), 
                 report.ageStats.elderly.excessive.toString(), 
                 report.ageStats.elderly.single.toString(), '', '', ''],
                // 通常年齢データ
                ['通常年齢（69歳以下）', report.ageStats.normal.total.toString(), 
                 report.ageStats.normal.excessive.toString(), '', '', '', '']
            ];
            
            // CSV文字列の作成（強化されたUTF-8エンコーディング）
            const csvContent = csvData.map(row => 
                row.map(cell => {
                    // セル内容をエスケープして文字化けを防止
                    const escapedCell = String(cell).replace(/"/g, '""');
                    return `"${escapedCell}"`;
                }).join(',')
            ).join('\r\n'); // Windows互換の改行コード
            
            // BOM付きUTF-8でBlobを作成（Excel対応）
            const bom = '\uFEFF';
            const blob = new Blob([bom + csvContent], { 
                type: 'text/csv;charset=utf-8;' 
            });
            
            const fileName = `${type === 'daily' ? '日報' : '月報'}_${new Date().toISOString().split('T')[0]}.csv`;
            
            if ((window as any).saveAs) {
                (window as any).saveAs(blob, fileName);
            } else {
                // フォールバック
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = fileName;
                link.click();
            }
            
        } catch (error) {
            console.error('CSV出力エラー:', error);
            alert('CSVの出力に失敗しました。');
        }
    }

    // 担当者別ランキングCSV出力（個別ファイル）
    private async exportStaffRankingCSVs(report: any): Promise<void> {
        try {
            const monthText = `${report.selectedYear}年${report.selectedMonth}月`;
            
            // ①契約者70歳以上の受注件数トップ10ランキング
            if (report.elderlyStaffRanking && report.elderlyStaffRanking.length > 0) {
                const elderlyData = [
                    ['①契約者70歳以上の受注件数トップ10ランキング'],
                    [''],
                    ['ランキング', '地区No.', '所属No.', '担当名', '件数'],
                    ...report.elderlyStaffRanking.map((staff: any) => [
                        staff.rank.toString(),
                        staff.regionNo,
                        staff.departmentNo,
                        staff.staffName,
                        staff.count.toString()
                    ])
                ];
                await this.downloadCSV(elderlyData, `①70歳以上受注件数ランキング_${monthText}.csv`);
            }

            // ②単独契約ランキング
            if (report.singleContractRanking && report.singleContractRanking.length > 0) {
                const singleData = [
                    ['②単独契約ランキング'],
                    [''],
                    ['ランキング', '地区No.', '所属No.', '担当名', '件数'],
                    ...report.singleContractRanking.map((staff: any) => [
                        staff.rank.toString(),
                        staff.regionNo,
                        staff.departmentNo,
                        staff.staffName,
                        staff.count.toString()
                    ])
                ];
                await this.downloadCSV(singleData, `②単独契約ランキング_${monthText}.csv`);
            }

            // ③過量販売ランキング
            if (report.excessiveSalesRanking && report.excessiveSalesRanking.length > 0) {
                const excessiveData = [
                    ['③過量販売ランキング'],
                    [''],
                    ['ランキング', '地区No.', '所属No.', '担当名', '件数'],
                    ...report.excessiveSalesRanking.map((staff: any) => [
                        staff.rank.toString(),
                        staff.regionNo,
                        staff.departmentNo,
                        staff.staffName,
                        staff.count.toString()
                    ])
                ];
                await this.downloadCSV(excessiveData, `③過量販売ランキング_${monthText}.csv`);
            }

            // ④69歳以下契約件数の担当別件数
            if (report.normalAgeStaffRanking && report.normalAgeStaffRanking.length > 0) {
                const normalData = [
                    ['④69歳以下契約件数の担当別件数'],
                    [''],
                    ['ランキング', '地区No.', '所属No.', '担当名', '件数'],
                    ...report.normalAgeStaffRanking.map((staff: any) => [
                        staff.rank.toString(),
                        staff.regionNo,
                        staff.departmentNo,
                        staff.staffName,
                        staff.count.toString()
                    ])
                ];
                await this.downloadCSV(normalData, `④69歳以下契約件数担当別_${monthText}.csv`);
            }

        } catch (error) {
            console.error('担当者別ランキングCSV出力エラー:', error);
        }
    }

    // CSVダウンロード用ヘルパーメソッド
    private async downloadCSV(data: any[][], fileName: string): Promise<void> {
        const csvContent = data.map(row => 
            row.map(cell => {
                const escapedCell = String(cell).replace(/"/g, '""');
                return `"${escapedCell}"`;
            }).join(',')
        ).join('\r\n');

        const bom = '\uFEFF';
        const blob = new Blob([bom + csvContent], { 
            type: 'text/csv;charset=utf-8;' 
        });

        if ((window as any).saveAs) {
            (window as any).saveAs(blob, fileName);
        } else {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = fileName;
            link.click();
        }
    }

    private createRankingTableHTML(ranking: any[]): string {
        if (!ranking || ranking.length === 0) {
            return '<div class="text-center text-muted py-3">データがありません</div>';
        }

        return `
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead class="table-dark">
                        <tr>
                            <th>ランキング</th>
                            <th>地区No.</th>
                            <th>所属No.</th>
                            <th>担当名</th>
                            <th>件数</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${ranking.map((staff: any) => `
                            <tr>
                                <td><strong>${staff.rank}</strong></td>
                                <td>${staff.regionNo}</td>
                                <td>${staff.departmentNo}</td>
                                <td>${staff.staffName}</td>
                                <td><span class="badge bg-primary">${staff.count}</span></td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        `;
    }

    // 担当別データを生成
    public generateStaffData(data: any[], targetDate: Date): StaffData[] {
        const staffData: StaffData[] = [];
        
        console.log(`generateStaffData開始: データ件数=${data.length}, targetDate=${targetDate.toISOString()}`);
        
        // 対象月のデータのみを抽出
        const targetYear = targetDate.getFullYear();
        const targetMonth = targetDate.getMonth();
        
        const monthlyData = data.filter(row => {
            if (!row.date) return false;
            return row.date.getFullYear() === targetYear && row.date.getMonth() === targetMonth;
        });
        
        console.log(`対象月(${targetYear}年${targetMonth + 1}月)のデータ件数: ${monthlyData.length}`);
        
        monthlyData.forEach((row, index) => {
            // 担当者名の正規化
            const normalizedStaffName = this.normalizeStaffName(row.staffName);
            
            // 受注条件の判定（シンプル化）
            const isOrder = this.isSimpleOrder(row);
            
            // 年齢の取得
            const ageNum = this.getContractorAge(row);
            
            // 最初の10件のみデバッグログ
            if (index < 10) {
                console.log(`行${index}: ${normalizedStaffName}, 年齢:${ageNum}, isOrder:${isOrder}, 確認:${row.confirmation}, K列:${row.confirmationDateTime}`);
            }
            
            if (normalizedStaffName && row.regionNumber && row.departmentNumber) {
                const existingStaff = staffData.find(s => 
                    s.regionNo === row.regionNumber && 
                    s.departmentNo === row.departmentNumber && 
                    s.staffName === normalizedStaffName
                );
                
                if (existingStaff) {
                    // 受注件数のカウント
                    if (isOrder) {
                        existingStaff.totalOrders++;
                        
                        // 年齢カウント
                        if (typeof ageNum === 'number') {
                            if (ageNum <= 69) {
                                existingStaff.normalAgeOrders++;
                            } else if (ageNum >= 70) {
                                existingStaff.elderlyOrders++;
                            }
                        }
                        
                        // その他のカウント
                        if (this.excelProcessor.isSingle(row)) {
                            existingStaff.singleOrders++;
                        }
                        if (this.excelProcessor.isExcessive(row)) {
                            existingStaff.excessiveOrders++;
                        }
                        if (this.excelProcessor.isOvertime(row)) {
                            existingStaff.overtimeOrders++;
                        }
                    }
                } else {
                    const newStaff: StaffData = {
                        regionNo: row.regionNumber,
                        departmentNo: row.departmentNumber,
                        staffName: normalizedStaffName,
                        totalOrders: isOrder ? 1 : 0,
                        normalAgeOrders: (isOrder && typeof ageNum === 'number' && ageNum <= 69) ? 1 : 0,
                        elderlyOrders: (isOrder && typeof ageNum === 'number' && ageNum >= 70) ? 1 : 0,
                        singleOrders: (isOrder && this.excelProcessor.isSingle(row)) ? 1 : 0,
                        excessiveOrders: (isOrder && this.excelProcessor.isExcessive(row)) ? 1 : 0,
                        overtimeOrders: (isOrder && this.excelProcessor.isOvertime(row)) ? 1 : 0
                    };
                    
                    if (index < 10) {
                        console.log(`新規担当者: ${normalizedStaffName}, 受注:${newStaff.totalOrders}, 69歳以下:${newStaff.normalAgeOrders}, 70歳以上:${newStaff.elderlyOrders}`);
                    }
                    
                    staffData.push(newStaff);
                }
            }
        });
        
        // 最終結果のサマリー
        const totalNormalAge = staffData.reduce((sum, staff) => sum + staff.normalAgeOrders, 0);
        const totalElderly = staffData.reduce((sum, staff) => sum + staff.elderlyOrders, 0);
        const totalOrders = staffData.reduce((sum, staff) => sum + staff.totalOrders, 0);
        console.log(`最終集計: 受注総数=${totalOrders}, 69歳以下=${totalNormalAge}, 70歳以上=${totalElderly}`);
        
        // デバッグ: 最初の10件の詳細
        console.log('=== 最初の10件の詳細 ===');
        staffData.slice(0, 10).forEach((staff, index) => {
            console.log(`${index}: ${staff.staffName} - 受注:${staff.totalOrders}, 69歳以下:${staff.normalAgeOrders}, 70歳以上:${staff.elderlyOrders}`);
        });
        
        return staffData;
    }
    
    // シンプルな受注判定
    private isSimpleOrder(row: any): boolean {
        // J列の確認区分チェック（1、2、5の場合は除外）
        const confirmation = row.confirmation;
        if (typeof confirmation === 'number') {
            if (confirmation === 1 || confirmation === 2 || confirmation === 5) {
                return false;
            }
        } else if (typeof confirmation === 'string') {
            const trimmedValue = confirmation.trim();
            if (trimmedValue !== '') {
                const confirmationNum = parseInt(trimmedValue);
                if (!isNaN(confirmationNum) && (confirmationNum === 1 || confirmationNum === 2 || confirmationNum === 5)) {
                    return false;
                }
            }
        }
        
        // K列の除外キーワードチェック
        if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
            const confirmationStr = row.confirmationDateTime;
            if (confirmationStr.includes('担当待ち') || confirmationStr.includes('直電')) {
                return false;
            }
        }
        
        return true;
    }

    // 文字列/数値いずれでも年齢を数値で取得する
    private getContractorAge(row: any): number | undefined {
        let age: any = row?.contractorAge ?? row?.age;
        
        if (age == null) {
            return undefined;
        }
        
        if (typeof age === 'number') {
            const result = Number.isFinite(age) ? age : undefined;
            return result;
        }
        
        if (typeof age === 'string') {
            const trimmedAge = age.trim();
            const matched = trimmedAge.match(/(\d+)/);
            if (matched) {
                const n = parseInt(matched[1], 10);
                if (n > 0 && n < 150) {
                    return n;
                }
            }
            if (/^\d+$/.test(trimmedAge)) {
                const n = parseInt(trimmedAge, 10);
                if (n > 0 && n < 150) {
                    return n;
                }
            }
        }
        
        return undefined;
    }
    
    // 担当者名を正規化するヘルパーメソッド
    private normalizeStaffName(staffName: any): string {
        if (!staffName || typeof staffName !== 'string') {
            return '';
        }
        
        const trimmedName = staffName.trim();
        if (trimmedName === '') {
            return '';
        }
        
        // 括弧（半角・全角）で囲まれた部分を除去
        // 例: "山下(弓弦)" → "山下"
        const normalizedName = trimmedName.replace(/[\(（].*?[\)）]/g, '');
        
        // 前後の空白を除去して返す
        return normalizedName.trim();
    }

    // 担当別データのCSV出力
    public async exportStaffDataToCSV(staffData: StaffData[]): Promise<void> {
        try {
            const reportMonthInput = document.getElementById('reportMonth') as HTMLInputElement;
            let fileName = '担当別データ_全期間.csv';
            
            if (reportMonthInput && reportMonthInput.value) {
                const [year, month] = reportMonthInput.value.split('-').map(Number);
                const monthText = `${year}年${month}月`;
                fileName = `担当別データ_${monthText}.csv`;
            }

            const csvData = [
                ['担当別データ'],
                [''],
                ['地区', '所属', '担当名', '受注件数', '契約者69歳以下件数', '70歳以上件数', '単独件数', '過量件数', '時間外件数'],
                ...staffData.map(staff => [
                    staff.regionNo || '',
                    staff.departmentNo || '',
                    staff.staffName,
                    staff.totalOrders.toString(),
                    staff.normalAgeOrders.toString(),
                    staff.elderlyOrders.toString(),
                    staff.singleOrders.toString(),
                    staff.excessiveOrders.toString(),
                    staff.overtimeOrders.toString()
                ])
            ];

            await this.downloadCSV(csvData, fileName);
        } catch (error) {
            console.error('担当別データCSV出力エラー:', error);
            alert('CSVの出力に失敗しました。');
        }
    }

    // 担当別データのHTMLを生成
    public createStaffDataHTML(staffData: StaffData[]): string {
        if (!staffData || staffData.length === 0) {
            return '<div class="alert alert-info">データがありません</div>';
        }

        const tableRows = staffData.map(staff => `
            <tr data-region="${staff.regionNo || ''}" data-department="${staff.departmentNo || ''}" data-staff="${staff.staffName}" data-orders="${staff.totalOrders}" data-normal-age="${staff.normalAgeOrders}" data-elderly="${staff.elderlyOrders}" data-single="${staff.singleOrders}" data-excessive="${staff.excessiveOrders}" data-overtime="${staff.overtimeOrders}">
                <td>${staff.regionNo || ''}</td>
                <td>${staff.departmentNo || ''}</td>
                <td>${staff.staffName}</td>
                <td>${staff.totalOrders}</td>
                <td>${staff.normalAgeOrders}</td>
                <td>${staff.elderlyOrders}</td>
                <td>${staff.singleOrders}</td>
                <td>${staff.excessiveOrders}</td>
                <td>${staff.overtimeOrders}</td>
            </tr>
        `).join('');

        return `
            <!-- 検索フォーム -->
            <div class="search-form mb-3">
                <div class="row g-3">
                    <div class="col-md-3">
                        <label for="staffNameSearch" class="form-label">担当名</label>
                        <input type="text" class="form-control" id="staffNameSearch" placeholder="担当名を入力">
                    </div>
                    <div class="col-md-2">
                        <label for="regionSearch" class="form-label">地区</label>
                        <input type="text" class="form-control" id="regionSearch" placeholder="地区番号">
                    </div>
                    <div class="col-md-2">
                        <label for="departmentSearch" class="form-label">所属</label>
                        <input type="text" class="form-control" id="departmentSearch" placeholder="所属番号">
                    </div>
                    <div class="col-md-2">
                        <label for="ordersSearch" class="form-label">受注件数</label>
                        <select class="form-select" id="ordersSearch">
                            <option value="">すべて</option>
                            <option value="0">0件</option>
                            <option value="5">5件以上</option>
                            <option value="10">10件以上</option>
                            <option value="15">15件以上</option>
                            <option value="20">20件以上</option>
                            <option value="30">30件以上</option>
                        </select>
                    </div>
                    <div class="col-md-2">
                        <label for="elderlySearch" class="form-label">70歳以上</label>
                        <select class="form-select" id="elderlySearch">
                            <option value="">すべて</option>
                            <option value="0">0件</option>
                            <option value="1">1件以上</option>
                            <option value="3">3件以上</option>
                        </select>
                    </div>
                </div>
                <div class="row mt-2">
                    <div class="col-12">
                        <button type="button" class="btn btn-primary btn-sm me-2" id="searchExecute">
                            <i class="fas fa-search me-1"></i>検索実行
                        </button>
                        <button type="button" class="btn btn-secondary btn-sm me-2" id="searchClear">
                            <i class="fas fa-times me-1"></i>条件クリア
                        </button>
                        <span class="text-muted" id="searchResultInfo">検索結果: ${staffData.length}件 / 総件数: ${staffData.length}件</span>
                    </div>
                </div>
            </div>

            <!-- データテーブル -->
            <div class="table-responsive">
                <table class="table table-striped table-hover" id="staffDataTable">
                    <thead class="table-dark">
                        <tr>
                            <th class="sortable" data-sort="region">
                                地区
                            </th>
                            <th class="sortable" data-sort="department">
                                所属
                            </th>
                            <th class="sortable" data-sort="staff">
                                担当名
                            </th>
                            <th class="sortable" data-sort="orders">
                                受注件数
                            </th>
                            <th class="sortable" data-sort="normalAge">
                                契約者69歳以下件数
                            </th>
                            <th class="sortable" data-sort="elderly">
                                70歳以上件数
                            </th>
                            <th class="sortable" data-sort="single">
                                単独件数
                            </th>
                            <th class="sortable" data-sort="excessive">
                                過量件数
                            </th>
                            <th class="sortable" data-sort="overtime">
                                時間外件数
                            </th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRows}
                    </tbody>
                </table>
            </div>
        `;
    }

    // 確認データのHTMLを生成
    public createDataConfirmationHTML(data: any[]): string {
        if (!data || data.length === 0) {
            return '<div class="alert alert-info">データがありません</div>';
        }

        const tableRows = data.map((row, index) => `
            <tr>
                <td>${index + 1}</td>
                <td>${row.date ? row.date.toLocaleDateString() : 'N/A'}</td>
                <td>${row.staffName || ''}</td>
                <td>${row.regionNumber || ''}</td>
                <td>${row.departmentNumber || ''}</td>
                <td>${row.contractor || ''}</td>
                <td>${row.contractorAge || ''}</td>
                <td>${row.confirmation || ''}</td>
                <td>${row.confirmationDateTime || ''}</td>
            </tr>
        `).join('');

        return `
            <div class="alert alert-success">
                <h5>データ確認完了</h5>
                <p>総データ件数: <strong>${data.length}</strong>件</p>
            </div>
            
            <!-- フィルター機能 -->
            <div class="row mb-3">
                <div class="col-md-4">
                    <label for="staffFilter" class="form-label">担当者名でフィルター</label>
                    <input type="text" class="form-control" id="staffFilter" placeholder="担当者名を入力...">
                </div>
                <div class="col-md-4">
                    <label for="regionFilter" class="form-label">地区№でフィルター</label>
                    <input type="text" class="form-control" id="regionFilter" placeholder="地区№を入力...">
                </div>
                <div class="col-md-4">
                    <label for="departmentFilter" class="form-label">所属№でフィルター</label>
                    <input type="text" class="form-control" id="departmentFilter" placeholder="所属№を入力...">
                </div>
            </div>
            
            <div class="table-responsive">
                <table class="table table-striped table-hover" id="dataConfirmationTable">
                    <thead class="table-dark">
                        <tr>
                            <th>#</th>
                            <th>日付</th>
                            <th>担当者名</th>
                            <th>地区№</th>
                            <th>所属№</th>
                            <th>契約者</th>
                            <th>年齢</th>
                            <th>確認</th>
                            <th>確認者日時</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRows}
                    </tbody>
                </table>
            </div>
            
            <div class="alert alert-info">
                <small>※ 全件表示中（${data.length}件）</small>
            </div>
        `;
    }

    // 日報PDF用HTML（ランキングなし）
    createDailyReportPDFHTML(report: any): string {
        // 選択された日付から日付テキストを取得
        const selectedDate = report.selectedDate ? new Date(report.selectedDate) : new Date();
        const dateText = selectedDate.toLocaleDateString('ja-JP');
        
        return `
            <div class="report-section" style="padding-top: 0;">
                <h3 class="report-title" style="margin-top: 0;">
                    <i class="fas fa-calendar-day me-2"></i>日報 - ${dateText}
                </h3>
                
                <!-- 基本統計 -->
                <div class="mb-4">
                    <div class="total-stats-container">
                        <h5 class="total-stats-title"><i class="fas fa-chart-bar me-2"></i>総件数</h5>
                        <div class="total-stats-grid">
                            <div class="total-stat-item">
                                <div class="total-stat-number">${report.totalOrders}</div>
                                <div class="total-stat-label">受注件数</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${report.overtimeOrders}</div>
                                <div class="total-stat-label">時間外対応</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalExcessive(report.regionStats)}</div>
                                <div class="total-stat-label">過量販売</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalSingle(report.regionStats)}</div>
                                <div class="total-stat-label">単独契約</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalHolidayConstruction(report.regionStats)}</div>
                                <div class="total-stat-label">公休日施工</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalProhibitedConstruction(report.regionStats)}</div>
                                <div class="total-stat-label">禁止日施工</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- 地区別受注件数 -->
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-map-marker-alt me-2"></i>地区別受注件数</h5>
                    ${this.createRegionStatsPDFHTML(report.regionStats)}
                </div>
                
                <!-- 年齢別集計 -->
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-users me-2"></i>年齢別集計</h5>
                    ${this.createAgeStatsHTML(report.ageStats)}
                </div>
            </div>
        `;
    }

    // PDF用地区統計HTML（全地区横並び）
    private createRegionStatsPDFHTML(regionStats: any): string {
        let html = '';
        
        // 地区の順序を固定
        const regionOrder = ['九州地区', '中四国地区', '関西地区', '関東地区'];
        const availableRegions = regionOrder.filter(region => 
            regionStats[region] && regionStats[region].orders > 0
        );
        
        if (availableRegions.length === 0) {
            return '<p>地区データがありません</p>';
        }
        
        // 2×2のグリッドレイアウトを強制
        html += '<div class="row" style="margin: 0; padding: 0;">';
        
        // 1段目: 九州｜中四国
        html += '<div class="col-md-6" style="padding: 0 5px 10px 0;">';
        if (availableRegions.includes('九州地区')) {
            html += this.createRegionCardHTML('九州地区', regionStats['九州地区']);
        }
        html += '</div>';
        
        html += '<div class="col-md-6" style="padding: 0 0 10px 5px;">';
        if (availableRegions.includes('中四国地区')) {
            html += this.createRegionCardHTML('中四国地区', regionStats['中四国地区']);
        }
        html += '</div>';
        
        html += '</div>';
        
        // 2段目: 関西｜関東
        html += '<div class="row" style="margin: 0; padding: 0;">';
        
        html += '<div class="col-md-6" style="padding: 5px 5px 0 0;">';
        if (availableRegions.includes('関西地区')) {
            html += this.createRegionCardHTML('関西地区', regionStats['関西地区']);
        }
        html += '</div>';
        
        html += '<div class="col-md-6" style="padding: 5px 0 0 5px;">';
        if (availableRegions.includes('関東地区')) {
            html += this.createRegionCardHTML('関東地区', regionStats['関東地区']);
        }
        html += '</div>';
        
        html += '</div>';
        
        return html;
    }
    
    // 地区カードのHTML生成（ヘルパーメソッド）
    private createRegionCardHTML(regionName: string, stats: any): string {
        return `
            <div class="region-card" style="height: 100%; margin: 0;">
                <div class="region-title">${regionName}</div>
                <div class="region-stats">
                    <div class="region-stat">
                        <div class="region-stat-number">${stats.orders}</div>
                        <div class="region-stat-label">受注件数</div>
                    </div>
                    <div class="region-stat">
                        <div class="region-stat-number">${stats.overtime}</div>
                        <div class="region-stat-label">時間外対応</div>
                    </div>
                    <div class="region-stat">
                        <div class="region-stat-number">${stats.excessive}</div>
                        <div class="region-stat-label">過量販売</div>
                    </div>
                    <div class="region-stat">
                        <div class="region-stat-number">${stats.single}</div>
                        <div class="region-stat-label">単独契約</div>
                    </div>
                    <div class="region-stat">
                        <div class="region-stat-number">${stats.holidayConstruction}</div>
                        <div class="region-stat-label">公休日施工</div>
                    </div>
                    <div class="region-stat">
                        <div class="region-stat-number">${stats.prohibitedConstruction}</div>
                        <div class="region-stat-label">禁止日施工</div>
                    </div>
                </div>
            </div>
        `;
    }

    // 月報PDF用HTML（2枚構成）
    createMonthlyReportPDFHTML(report: any): string {
        // 選択された月から年月を取得
        const selectedDate = report.selectedMonth ? new Date(report.selectedMonth + '-01') : new Date();
        const monthText = selectedDate.toLocaleDateString('ja-JP', { year: 'numeric', month: 'long' });
        
        return `
            <!-- 1枚目: 基本統計（ランキングなし） -->
            <div class="report-section" style="padding-top: 0; page-break-after: always; page-break-inside: avoid;">
                <h3 class="report-title" style="margin-top: 0;">
                    <i class="fas fa-calendar-alt me-2"></i>月報 - ${monthText}
                </h3>
                
                <!-- 基本統計 -->
                <div class="mb-4">
                    <div class="total-stats-container">
                        <h5 class="total-stats-title"><i class="fas fa-chart-bar me-2"></i>総件数</h5>
                        <div class="total-stats-grid">
                            <div class="total-stat-item">
                                <div class="total-stat-number">${report.totalOrders}</div>
                                <div class="total-stat-label">受注件数</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${report.overtimeOrders}</div>
                                <div class="total-stat-label">時間外対応</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalExcessive(report.regionStats)}</div>
                                <div class="total-stat-label">過量販売</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalSingle(report.regionStats)}</div>
                                <div class="total-stat-label">単独契約</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalHolidayConstruction(report.regionStats)}</div>
                                <div class="total-stat-label">公休日施工</div>
                            </div>
                            <div class="total-stat-item">
                                <div class="total-stat-number">${this.getTotalProhibitedConstruction(report.regionStats)}</div>
                                <div class="total-stat-label">禁止日施工</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- 地区別受注件数 -->
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-map-marker-alt me-2"></i>地区別受注件数</h5>
                    ${this.createRegionStatsPDFHTML(report.regionStats)}
                </div>
                
                <!-- 年齢別集計 -->
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-users me-2"></i>年齢別集計</h5>
                    ${this.createAgeStatsHTML(report.ageStats)}
                </div>
            </div>
            
            <!-- 強制ページ区切り -->
            <div style="page-break-before: always; height: 0; margin: 0; padding: 0;"></div>
            
            <!-- 2枚目: ランキングのみ -->
            <div class="report-section" style="padding-top: 0; page-break-before: always; page-break-after: avoid; page-break-inside: avoid;">
                <h3 class="report-title" style="margin-top: 0;">
                    <i class="fas fa-trophy me-2"></i>担当者別ランキング - ${monthText}
                </h3>
                
                ${report.elderlyStaffRanking ? `
                <!-- ①契約者70歳以上の受注件数トップ10ランキング -->
                <div class="ranking-section mb-4">
                    <h5 class="ranking-title">①契約者70歳以上の受注件数トップ10ランキング</h5>
                    ${this.createRankingTableHTML(report.elderlyStaffRanking)}
                </div>

                <!-- ②単独契約を持っている担当者一覧 -->
                <div class="ranking-section mb-4">
                    <h5 class="ranking-title">②単独契約を持っている担当者一覧</h5>
                    ${this.createRankingTableHTML(report.singleContractRanking)}
                </div>

                <!-- ③過量契約を持っている担当者一覧 -->
                <div class="ranking-section mb-4">
                    <h5 class="ranking-title">③過量契約を持っている担当者一覧</h5>
                    ${this.createRankingTableHTML(report.excessiveSalesRanking)}
                </div>

                <!-- ④69歳以下契約件数の担当別件数 -->
                <div class="ranking-section mb-4">
                    <h5 class="ranking-title">④69歳以下契約件数の担当別件数</h5>
                    ${this.createRankingTableHTML(report.normalAgeStaffRanking)}
                </div>
                ` : '<p>ランキングデータがありません</p>'}
            </div>
        `;
    }
}