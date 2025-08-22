import { HolidaySettings } from './types.js';

export class ReportGenerator {
    private currentTargetDate: Date | null = null;
    private holidaySettings: HolidaySettings = {
        publicHolidays: [],
        prohibitedDays: []
    };
    
    // 公休日・禁止日設定を更新
    updateHolidaySettings(settings: HolidaySettings): void {
        this.holidaySettings = settings;
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
        return this.holidaySettings.publicHolidays.some(holiday => 
            this.isSameDate(date, holiday)
        );
    }
    
    // 禁止日かどうかの判定
    private isProhibitedDay(date: Date): boolean {
        return this.holidaySettings.prohibitedDays.some(prohibited => 
            this.isSameDate(date, prohibited)
        );
    }
    
    generateDailyReport(data: any[], date: string): any {
        const targetDate = new Date(date);
        this.currentTargetDate = targetDate; // 時間外判定で使用
        const targetMonth = targetDate.getMonth();
        const targetDay = targetDate.getDate();
        
        console.log('日報生成開始:', {
            targetDate: targetDate.toLocaleDateString(),
            targetMonth: targetMonth,
            targetDay: targetDay,
            totalData: data.length
        });
        
        const dailyData = data.filter(row => {
            // A列の日付チェック
            let isDateMatch = false;
            
            if (row.date && row.date instanceof Date) {
                const rowMonth = row.date.getMonth();
                const rowDay = row.date.getDate();
                if (rowMonth === targetMonth && rowDay === targetDay) {
                    isDateMatch = true;
                }
            }
            
            // K列の日付チェック（A列でマッチしない場合）
            if (!isDateMatch && row.confirmationDateTime) {
                const kColumnStr = String(row.confirmationDateTime);
                const targetDateStr = `${targetMonth + 1}/${targetDay}`;
                const targetDateStrAlt = `${targetMonth + 1}月${targetDay}日`;
                
                if (kColumnStr.includes(targetDateStr) || kColumnStr.includes(targetDateStrAlt)) {
                    isDateMatch = true;
                }
                
                // K列に日付+時間の形式で該当日付が含まれるかチェック
                const kDateMatch = kColumnStr.match(/(\d{1,2})\/(\d{1,2})/);
                if (kDateMatch) {
                    const kMonth = parseInt(kDateMatch[1]);
                    const kDay = parseInt(kDateMatch[2]);
                    if (kMonth === targetMonth + 1 && kDay === targetDay) {
                        isDateMatch = true;
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
                if (kColumnStr.includes('担当待ち') || kColumnStr.includes('直電') || 
                    kColumnStr.includes('契約時') || kColumnStr.includes('契約') || 
                    kColumnStr.includes('待ち')) {
                    isKColumnValid = false;
                }
            }
            
            const isValid = isDateMatch && isJColumnValid && isKColumnValid;
            
            return isValid;
        });
        
        console.log('日報生成結果:', {
            targetDate: targetDate.toLocaleDateString(),
            totalData: data.length,
            filteredData: dailyData.length,
            filteredOut: data.length - dailyData.length
        });
        
        const reportData = this.calculateReportData(dailyData, 'daily');
        // 選択された日付情報を追加
        reportData.selectedDate = date;
        return reportData;
    }
    
    generateMonthlyReport(data: any[], month: string): any {
        const [year, monthNum] = month.split('-').map(Number);
        const monthlyData = data.filter(row => {
            // 日付チェック
            if (!row.date) return false;
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
                if (kColumnStr.includes('担当待ち') || kColumnStr.includes('直電') || 
                    kColumnStr.includes('契約時') || kColumnStr.includes('契約') || 
                    kColumnStr.includes('待ち')) {
                    isKColumnValid = false;
                }
            }
            
            return isJColumnValid && isKColumnValid;
        });
        
        const reportData = this.calculateReportData(monthlyData, 'monthly');
        // 選択された月情報を追加
        reportData.selectedMonth = month;
        return reportData;
    }
    
    // 時間外カウント取得（AB列とK列を独立してカウント）
    private getOvertimeCount(row: any, targetDate?: Date | null): number {
        if (!targetDate) return 0;
        
        const targetMonth = targetDate.getMonth();
        const targetDay = targetDate.getDate();
        
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
                        console.log('8/20データの時間チェック:', {
                            contractor: row.contractor,
                            time: row.time,
                            timeType: typeof row.time,
                            hours, minutes, totalMinutes,
                            cutoff: 18 * 60 + 30,
                            isOvertime: totalMinutes >= 18 * 60 + 30
                        });
                    }
                    
                    if (totalMinutes >= 18 * 60 + 30) { // 18:30以降
                        contractorOvertimeCount = 1;
                        console.log('時間外検出（A列+B列）:', {
                            date: row.date.toLocaleDateString(),
                            time: row.time,
                            timeType: typeof row.time,
                            hours, minutes, totalMinutes,
                            cutoff: 18 * 60 + 30
                        });
                    }
                }
            }
        }
        
        // ②K列の文字列を日付と時間と個人名に分解して、日付が対象日付であり、時間が18:30以降であるものをカウント
        if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
            const kColumnStr = String(row.confirmationDateTime);
            
            if (isTarget820) {
                console.log('8/20データのK列チェック:', {
                    contractor: row.contractor,
                    kColumnStr: kColumnStr,
                    includes同時: kColumnStr === '同時',
                    includesColon: kColumnStr.includes(':')
                });
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
                            console.log('時間外検出（K列）:', {
                                kColumnStr,
                                kMonth, kDay, kHour, kMinute,
                                kTotalMinutes,
                                cutoff: 18 * 60 + 30
                            });
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
                                        console.log('時間外検出（K列スペース区切り）:', {
                                            kColumnStr,
                                            dateStr, timeStr,
                                            month, day, hours, minutes,
                                            totalMinutes,
                                            cutoff: 18 * 60 + 30
                                        });
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
            console.log('8/20データの最終判定:', {
                contractor: row.contractor,
                time: row.time,
                confirmationDateTime: row.confirmationDateTime,
                contractorCount: contractorOvertimeCount,
                confirmerCount: confirmerOvertimeCount,
                totalCount: totalCount
            });
        }
        
        if (totalCount > 0) {
            console.log('時間外として判定:', {
                contractor: row.contractor,
                date: row.date?.toLocaleDateString(),
                time: row.time,
                confirmationDateTime: row.confirmationDateTime,
                contractorCount: contractorOvertimeCount,
                confirmerCount: confirmerOvertimeCount,
                totalCount: totalCount
            });
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
            const region = row.region;
            if (regions[region]) {
                regions[region].orders++;
                regions[region].overtime += this.getOvertimeCount(row, this.currentTargetDate);
                if (row.isExcessive) {
                    regions[region].excessive++;
                }
                if (row.isSingle) {
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
            if (row.isElderly) {
                stats.elderly.total++;
                if (row.isExcessive) stats.elderly.excessive++;
                if (row.isSingle) stats.elderly.single++;
            } else {
                stats.normal.total++;
                if (row.isExcessive) stats.normal.excessive++;
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
                                <div class="stat-number">${report.overtimeOrders}</div>
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
            
            // html2canvasでHTML要素を画像に変換
            const canvas = await (window as any).html2canvas(pdfContainer, {
                scale: 3, // より高画質化
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                logging: false
            });
            
            // jsPDFでPDF作成
            const { jsPDF } = (window as any).jspdf;
            const pdf = new jsPDF({
                orientation: 'portrait',
                unit: 'mm',
                format: 'a4'
            });
            
            // A4サイズの寸法（mm）
            const pageWidth = pdf.internal.pageSize.getWidth(); // 210mm
            const pageHeight = pdf.internal.pageSize.getHeight(); // 297mm
            
            // 画像のアスペクト比を計算
            const imgWidth = canvas.width;
            const imgHeight = canvas.height;
            const imgAspectRatio = imgWidth / imgHeight;
            
            // 余白を設定（上下左右10mm）
            const margin = 10;
            const availableWidth = pageWidth - (margin * 2);
            const availableHeight = pageHeight - (margin * 2);
            
            // 利用可能エリアに合わせてサイズを計算
            let finalWidth, finalHeight;
            if (availableWidth / availableHeight > imgAspectRatio) {
                // 高さ基準でスケール
                finalHeight = availableHeight;
                finalWidth = finalHeight * imgAspectRatio;
            } else {
                // 幅基準でスケール
                finalWidth = availableWidth;
                finalHeight = finalWidth / imgAspectRatio;
            }
            
            // 中央配置のための座標計算
            const x = (pageWidth - finalWidth) / 2;
            const y = (pageHeight - finalHeight) / 2;
            
            // 画像をPDFに追加（中央配置・最大サイズ）
            const imgData = canvas.toDataURL('image/png');
            pdf.addImage(imgData, 'PNG', x, y, finalWidth, finalHeight);
            
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
        
        // 日付の取得
        const selectedDate = report.selectedDate ? new Date(report.selectedDate) : new Date();
        const dateText = selectedDate.toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
        
        container.innerHTML = `
            <div class="pdf-header">
                <div class="pdf-title">${type === 'daily' ? '日報' : '月報'}</div>
                <div class="pdf-date">${dateText}</div>
            </div>
            
                            <!-- 基本統計 -->
                <div class="pdf-section">
                    <div class="pdf-section-title">総件数</div>
                    <div class="pdf-total-stats-grid">
                        <div class="pdf-total-stat-item">
                            <div class="pdf-total-stat-number">${report.totalOrders}</div>
                            <div class="pdf-total-stat-label">受注件数</div>
                        </div>
                        <div class="pdf-total-stat-item">
                            <div class="pdf-total-stat-number">${report.overtimeOrders}</div>
                            <div class="pdf-total-stat-label">時間外対応</div>
                        </div>
                        <div class="pdf-total-stat-item">
                            <div class="pdf-total-stat-number">${this.getTotalExcessive(report.regionStats)}</div>
                            <div class="pdf-total-stat-label">過量販売</div>
                        </div>
                        <div class="pdf-total-stat-item">
                            <div class="pdf-total-stat-number">${this.getTotalSingle(report.regionStats)}</div>
                            <div class="pdf-total-stat-label">単独契約</div>
                        </div>
                        <div class="pdf-total-stat-item">
                            <div class="pdf-total-stat-number">${this.getTotalHolidayConstruction(report.regionStats)}</div>
                            <div class="pdf-total-stat-label">公休日施工</div>
                        </div>
                        <div class="pdf-total-stat-item">
                            <div class="pdf-total-stat-number">${this.getTotalProhibitedConstruction(report.regionStats)}</div>
                            <div class="pdf-total-stat-label">禁止日施工</div>
                        </div>
                    </div>
                </div>
            
            <!-- 地区別受注件数 -->
            <div class="pdf-section">
                <div class="pdf-section-title">地区別受注件数</div>
                <div class="pdf-region-grid">
                    ${Object.entries(report.regionStats)
                        .filter(([_, stats]: [string, any]) => stats.orders > 0)
                        .map(([region, stats]: [string, any]) => `
                            <div class="pdf-region-card">
                                <div class="pdf-region-name">${region}</div>
                                <div class="pdf-region-stats">
                                    受注件数: ${stats.orders}件<br>
                                    時間外対応: ${stats.overtime}件<br>
                                    過量販売: ${stats.excessive}件<br>
                                    単独契約: ${stats.single}件<br>
                                    公休日施工: ${stats.holidayConstruction}件<br>
                                    禁止日施工: ${stats.prohibitedConstruction}件
                                </div>
                            </div>
                        `).join('')}
                </div>
            </div>
            
            <!-- 年齢別集計 -->
            <div class="pdf-section">
                <div class="pdf-section-title">年齢別集計</div>
                <div class="pdf-age-grid">
                    <div class="pdf-age-card">
                        <div class="pdf-age-title">高齢者（70歳以上）</div>
                        <div class="pdf-age-stats">
                            総件数: ${report.ageStats.elderly.total}件<br>
                            過量販売: ${report.ageStats.elderly.excessive}件<br>
                            単独契約: ${report.ageStats.elderly.single}件
                        </div>
                    </div>
                    <div class="pdf-age-card">
                        <div class="pdf-age-title">通常年齢（69歳以下）</div>
                        <div class="pdf-age-stats">
                            総件数: ${report.ageStats.normal.total}件<br>
                            過量販売: ${report.ageStats.normal.excessive}件
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        return container;
    }
    
    async exportToCSV(report: any, type: string): Promise<void> {
        try {
            // CSVデータの作成（横項目形式）
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
}