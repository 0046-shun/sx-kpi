export class ReportGenerator {
    constructor(excelProcessor, calendarManager) {
        this.currentTargetDate = null;
        this.excelProcessor = excelProcessor;
        this.calendarManager = calendarManager;
    }
    // 公休日・禁止日設定を更新
    updateHolidaySettings(settings) {
        if (this.calendarManager) {
            this.calendarManager.updateSettings(settings);
        }
    }
    // 地区名を取得
    getRegionName(regionNumber) {
        if (!regionNumber)
            return 'その他';
        const number = parseInt(regionNumber);
        if (isNaN(number))
            return 'その他';
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
    isHolidayConstruction(row) {
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
    isProhibitedConstruction(row) {
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
    isPublicHoliday(date) {
        if (this.calendarManager) {
            return this.calendarManager.isPublicHoliday(date);
        }
        return false;
    }
    // 禁止日かどうかの判定
    isProhibitedDay(date) {
        if (this.calendarManager) {
            return this.calendarManager.isProhibitedDay(date);
        }
        return false;
    }
    generateDailyReport(data, date) {
        const targetDate = new Date(date);
        this.currentTargetDate = targetDate; // 時間外判定で使用
        // excel-processor.tsのisOrderForDateメソッドを使用して日付フィルタリング
        const dailyData = data.filter(row => {
            return this.excelProcessor.isOrderForDate(row, targetDate, true); // 日報生成時はtrue
        });
        const reportData = this.calculateReportData(dailyData, 'daily');
        // 選択された日付情報を追加
        reportData.selectedDate = date;
        return reportData;
    }
    generateMonthlyReport(data, month) {
        const [year, monthNum] = month.split('-').map(Number);
        // 月報の対象月の最初の日をtargetDateとして設定
        const targetDate = new Date(year, monthNum - 1, 1);
        // excel-processor.tsのisOrderForDateメソッドを使用して月報フィルタリング
        const monthlyData = data.filter(row => this.excelProcessor.isOrderForDate(row, targetDate, false) // 月報生成時はfalse
        );
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
    getOvertimeCount(row, targetDate) {
        // 月報の場合は、行の日付を使用
        let effectiveTargetDate;
        if (!targetDate) {
            if (row.date && row.date instanceof Date) {
                effectiveTargetDate = row.date;
            }
            else {
                return 0;
            }
        }
        else {
            effectiveTargetDate = targetDate;
        }
        const targetMonth = effectiveTargetDate.getMonth();
        const targetDay = effectiveTargetDate.getDate();
        // 8/20のデータのみ詳細ログを出力
        const isTarget820 = row.date && row.date instanceof Date &&
            row.date.getMonth() === 7 && row.date.getDate() === 20;
        let contractorOvertimeCount = 0; // ①のカウント
        let confirmerOvertimeCount = 0; // ②のカウント
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
    isOvertime(row, targetDate) {
        return this.getOvertimeCount(row, targetDate) > 0;
    }
    calculateReportData(data, type) {
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
    calculateRegionStats(data) {
        const regions = {
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
    calculateAgeStats(data) {
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
            }
            else {
                stats.normal.total++;
                if (this.excelProcessor.isExcessive(row)) {
                    stats.normal.excessive++;
                }
            }
        });
        return stats;
    }
    isSameDate(date1, date2) {
        return date1.getFullYear() === date2.getFullYear() &&
            date1.getMonth() === date2.getMonth() &&
            date1.getDate() === date2.getDate();
    }
    // 総過量販売件数を取得
    getTotalExcessive(regionStats) {
        return Object.values(regionStats).reduce((total, stats) => {
            return total + (stats.excessive || 0);
        }, 0);
    }
    // 総単独契約件数を取得
    getTotalSingle(regionStats) {
        return Object.values(regionStats).reduce((total, stats) => {
            return total + (stats.single || 0);
        }, 0);
    }
    // 総公休日施工件数を取得
    getTotalHolidayConstruction(regionStats) {
        return Object.values(regionStats).reduce((total, stats) => {
            return total + (stats.holidayConstruction || 0);
        }, 0);
    }
    // 総禁止日施工件数を取得
    getTotalProhibitedConstruction(regionStats) {
        return Object.values(regionStats).reduce((total, stats) => {
            return total + (stats.prohibitedConstruction || 0);
        }, 0);
    }
    // 受注判定メソッド（日付指定版）
    isOrderForDate(row, targetDate) {
        return this.excelProcessor.isOrderForDate(row, targetDate, true); // 日報判定として扱う
    }
    // 担当者別ランキング集計メソッド
    // 契約者70歳以上の受注件数トップ10ランキング
    calculateElderlyStaffRanking(data, targetYear, targetMonth) {
        const staffCounts = new Map();
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
                }
                else {
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
    calculateSingleContractRanking(data, targetYear, targetMonth) {
        const staffCounts = new Map();
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
                }
                else {
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
    calculateExcessiveSalesRanking(data, targetYear, targetMonth) {
        const staffCounts = new Map();
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
                }
                else {
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
    calculateNormalAgeStaffRanking(data, targetYear, targetMonth) {
        const staffCounts = new Map();
        // 月報の対象月の最初の日をtargetDateとして設定
        const targetDate = targetYear && targetMonth !== undefined ? new Date(targetYear, targetMonth, 1) : new Date();
        data.forEach((row, index) => {
            // 動的にisOrderを計算（月報の対象月で判定）
            const isOrder = this.isOrderForDate(row, targetDate);
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
                }
                else {
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
    assignRanks(sorted) {
        if (sorted.length === 0)
            return [];
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
    createDailyReportHTML(report) {
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
    createMonthlyReportHTML(report) {
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
    createRegionStatsHTML(regionStats) {
        let html = '<div class="row">';
        Object.entries(regionStats).forEach(([region, stats]) => {
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
    createAgeStatsHTML(ageStats) {
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
    async exportToPDF(report, type) {
        try {
            // 処理開始メッセージ
            const messageDiv = document.createElement('div');
            messageDiv.id = 'pdf-processing-message';
            messageDiv.style.cssText = `
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: #007bff;
                color: white;
                padding: 20px 40px;
                border-radius: 10px;
                font-size: 18px;
                font-weight: bold;
                z-index: 10000;
                box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            `;
            messageDiv.innerHTML = `
                <div style="text-align: center;">
                    <div style="margin-bottom: 15px;">🔄 PDF生成中...</div>
                    <div style="font-size: 14px; opacity: 0.9;">しばらくお待ちください</div>
                </div>
            `;
            document.body.appendChild(messageDiv);
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 50));
            // PDF用のHTML要素を動的に作成（非同期で処理）
            const pdfContainer = await this.createPDFHTMLAsync(report, type);
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 50));
            // コンテナを画面外に配置（レンダリング用）
            // A4サイズ（210mm）から左右余白（20mm）を引いた幅に設定
            const containerWidth = 794 - (10 * 794 / 210); // 約715px
            pdfContainer.style.cssText = `
                position: absolute;
                left: -9999px;
                top: -9999px;
                visibility: visible;
                width: ${containerWidth}px;
                height: auto;
                background: white;
                z-index: -1;
            `;
            document.body.appendChild(pdfContainer);
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 50));
            // コンテンツが確実にレンダリングされるまで待機
            await new Promise(resolve => setTimeout(resolve, 200));
            // jsPDFインスタンス（A4サイズ、余白10mm）
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF({
                orientation: 'portrait',
                unit: 'mm',
                format: 'a4',
                compress: true
            });
            const pageWidth = pdf.internal.pageSize.getWidth(); // 210mm
            const pageHeight = pdf.internal.pageSize.getHeight(); // 297mm
            const marginX = 10; // 左右余白10mm
            const marginY = 10; // 上下余白10mm
            // メッセージ更新
            messageDiv.innerHTML = `
                <div style="text-align: center;">
                    <div style="margin-bottom: 15px;">📄 PDF生成中...</div>
                    <div style="font-size: 14px; opacity: 0.9;">高品質レンダリング中</div>
                </div>
            `;
            // 月報の場合はページ分割処理
            if (type === 'monthly' && report.elderlyStaffRanking) {
                await this.exportMonthlyReportToPDF(pdfContainer, pdf, pageWidth, pageHeight, marginX, marginY, messageDiv);
            }
            else {
                // 日報の場合も個別ページ処理に変更（月報と同じ方式）
                await this.exportMonthlyReportToPDF(pdfContainer, pdf, pageWidth, pageHeight, marginX, marginY, messageDiv);
            }
            // 一時的なHTML要素を完全削除
            if (pdfContainer && pdfContainer.parentNode) {
                pdfContainer.parentNode.removeChild(pdfContainer);
            }
            // PDFをダウンロード（対象日付・月を使用）
            let fileName = '';
            if (type === 'daily') {
                // 日報の場合：対象日付を使用
                if (report.selectedDate) {
                    const selectedDate = new Date(report.selectedDate);
                    const year = selectedDate.getFullYear();
                    const month = String(selectedDate.getMonth() + 1).padStart(2, '0');
                    const day = String(selectedDate.getDate()).padStart(2, '0');
                    fileName = `${year}-${month}-${day}.pdf`;
                }
                else {
                    // フォールバック：現在日付
                    fileName = `${new Date().toISOString().split('T')[0]}.pdf`;
                }
            }
            else {
                // 月報の場合：対象月を使用
                if (report.selectedMonth) {
                    // selectedMonthが "2025-08" 形式の場合
                    if (typeof report.selectedMonth === 'string' && report.selectedMonth.includes('-')) {
                        fileName = `${report.selectedMonth}.pdf`;
                    }
                    else if (report.selectedYear && report.selectedMonth) {
                        // selectedMonthが数値の場合
                        const year = report.selectedYear;
                        const month = String(report.selectedMonth).padStart(2, '0');
                        fileName = `${year}-${month}.pdf`;
                    }
                    else {
                        // フォールバック：現在年月
                        const now = new Date();
                        const year = now.getFullYear();
                        const month = String(now.getMonth() + 1).padStart(2, '0');
                        fileName = `${year}-${month}.pdf`;
                    }
                }
                else {
                    // フォールバック：現在年月
                    const now = new Date();
                    const year = now.getFullYear();
                    const month = String(now.getMonth() + 1).padStart(2, '0');
                    fileName = `${year}-${month}.pdf`;
                }
            }
            pdf.save(fileName);
            // 完了メッセージ
            messageDiv.innerHTML = `
                <div style="text-align: center;">
                    <div style="margin-bottom: 15px;">✅ PDF生成完了！</div>
                    <div style="font-size: 14px; opacity: 0.9;">${fileName}</div>
                </div>
            `;
            // 3秒後にメッセージを自動削除
            setTimeout(() => {
                if (messageDiv && messageDiv.parentNode) {
                    messageDiv.parentNode.removeChild(messageDiv);
                }
            }, 3000);
        }
        catch (error) {
            console.error('PDF出力エラー:', error);
            // エラーメッセージを表示
            const messageDiv = document.getElementById('pdf-processing-message');
            if (messageDiv) {
                messageDiv.innerHTML = `
                    <div style="text-align: center;">
                        <div style="margin-bottom: 15px;">❌ PDF生成エラー</div>
                        <div style="font-size: 14px; opacity: 0.9;">エラーが発生しました</div>
                    </div>
                `;
                messageDiv.style.background = '#dc3545';
                // 5秒後に削除
                setTimeout(() => {
                    if (messageDiv && messageDiv.parentNode) {
                        messageDiv.parentNode.removeChild(messageDiv);
                    }
                }, 5000);
            }
            alert('PDFの出力に失敗しました。');
        }
    }
    // 月報・日報用PDF出力（個別ページ処理）
    async exportMonthlyReportToPDF(pdfContainer, pdf, pageWidth, pageHeight, marginX, marginY, messageDiv) {
        // 各ページを個別に処理
        const pages = pdfContainer.querySelectorAll('.pdf-page');
        for (let i = 0; i < pages.length; i++) {
            const page = pages[i];
            const pageNumber = i + 1;
            // メッセージ更新
            messageDiv.innerHTML = `
                <div style="text-align: center;">
                    <div style="margin-bottom: 15px;">📄 PDFページ生成中...</div>
                    <div style="font-size: 14px; opacity: 0.9;">${pageNumber}枚目 / ${pages.length}枚</div>
                </div>
            `;
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 50));
            // 各ページを個別にhtml2canvasで処理
            // HTMLコンテナと同じ幅を使用
            const canvasWidth = 794 - (19 * 794 / 210); // 約711px（微調整）
            const canvas = await window.html2canvas(page, {
                scale: 2,
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                logging: false,
                width: canvasWidth, // 794 → 約711px
                height: 1123,
                scrollX: 0,
                scrollY: 0,
                letterRendering: true,
                imageTimeout: 10000
            });
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 50));
            // 2ページ目以降は新しいページを追加
            if (i > 0) {
                pdf.addPage();
            }
            // 画像をPDFに追加（A4サイズに合わせて調整）
            const imgData = canvas.toDataURL('image/png', 1.0);
            const imgWpx = canvas.width;
            const imgHpx = canvas.height;
            const imgAspect = imgWpx / imgHpx;
            // ページの利用可能領域
            const availW = pageWidth - marginX * 2; // 190mm
            const availH = pageHeight - marginY * 2; // 277mm
            // 高さ優先で拡大
            let drawHeight = availH;
            let drawWidth = drawHeight * imgAspect;
            // 幅がはみ出す場合は幅優先に切り替え
            if (drawWidth > availW) {
                drawWidth = availW;
                drawHeight = drawWidth / imgAspect;
            }
            // 中央配置（marginXを考慮）
            const x = marginX + (availW - drawWidth) / 2;
            const y = marginY;
            pdf.addImage(imgData, 'PNG', x, y, drawWidth, drawHeight, undefined, 'FAST');
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 50));
        }
        // メッセージ更新
        messageDiv.innerHTML = `
            <div style="text-align: center;">
                <div style="margin-bottom: 15px;">💾 PDF保存中...</div>
                <div style="font-size: 14px; opacity: 0.9;">完了までしばらくお待ちください</div>
            </div>
        `;
    }
    createPDFHTML(report, type) {
        const container = document.createElement('div');
        container.className = 'pdf-container';
        let htmlContent = '';
        // 新しく作成したPDF用メソッドを使用
        if (type === 'daily') {
            // 日付テキストを生成
            let dateText = '';
            if (report.selectedDate) {
                const selectedDate = new Date(report.selectedDate);
                dateText = selectedDate.toLocaleDateString('ja-JP');
            }
            else {
                const now = new Date();
                dateText = now.toLocaleDateString('ja-JP');
            }
            htmlContent = this.createDailyReportPDFHTML(report, dateText);
        }
        else if (type === 'monthly') {
            htmlContent = this.createMonthlyReportPDFHTML(report);
        }
        else {
            // フォールバック
            let dateText = '';
            if (report.selectedDate) {
                const selectedDate = new Date(report.selectedDate);
                dateText = selectedDate.toLocaleDateString('ja-JP');
            }
            else {
                const now = new Date();
                dateText = now.toLocaleDateString('ja-JP');
            }
            htmlContent = this.createDailyReportPDFHTML(report, dateText);
        }
        // innerHTMLではなく、insertAdjacentHTMLを使用
        container.insertAdjacentHTML('beforeend', htmlContent);
        return container;
    }
    async createPDFHTMLAsync(report, type) {
        const container = document.createElement('div');
        container.className = 'pdf-container';
        // メインスレッドを解放
        await new Promise(resolve => setTimeout(resolve, 10));
        let htmlContent = '';
        // 新しく作成したPDF用メソッドを使用
        if (type === 'daily') {
            // 日付テキストを生成
            let dateText = '';
            if (report.selectedDate) {
                const selectedDate = new Date(report.selectedDate);
                dateText = selectedDate.toLocaleDateString('ja-JP');
            }
            else {
                const now = new Date();
                dateText = now.toLocaleDateString('ja-JP');
            }
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 10));
            htmlContent = this.createDailyReportPDFHTML(report, dateText);
        }
        else if (type === 'monthly') {
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 10));
            htmlContent = this.createMonthlyReportPDFHTML(report);
        }
        else {
            // フォールバック
            let dateText = '';
            if (report.selectedDate) {
                const selectedDate = new Date(report.selectedDate);
                dateText = selectedDate.toLocaleDateString('ja-JP');
            }
            else {
                const now = new Date();
                dateText = now.toLocaleDateString('ja-JP');
            }
            // メインスレッドを解放
            await new Promise(resolve => setTimeout(resolve, 10));
            htmlContent = this.createDailyReportPDFHTML(report, dateText);
        }
        // メインスレッドを解放
        await new Promise(resolve => setTimeout(resolve, 10));
        // innerHTMLではなく、insertAdjacentHTMLを使用
        container.insertAdjacentHTML('beforeend', htmlContent);
        return container;
    }
    async exportToCSV(report, type) {
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
                    .filter(([_, stats]) => stats.orders > 0)
                    .map(([region, stats]) => [
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
            const csvContent = csvData.map(row => row.map(cell => {
                // セル内容をエスケープして文字化けを防止
                const escapedCell = String(cell).replace(/"/g, '""');
                return `"${escapedCell}"`;
            }).join(',')).join('\r\n'); // Windows互換の改行コード
            // BOM付きUTF-8でBlobを作成（Excel対応）
            const bom = '\uFEFF';
            const blob = new Blob([bom + csvContent], {
                type: 'text/csv;charset=utf-8;'
            });
            const fileName = `${type === 'daily' ? '日報' : '月報'}_${new Date().toISOString().split('T')[0]}.csv`;
            if (window.saveAs) {
                window.saveAs(blob, fileName);
            }
            else {
                // フォールバック
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = fileName;
                link.click();
            }
        }
        catch (error) {
            console.error('CSV出力エラー:', error);
            alert('CSVの出力に失敗しました。');
        }
    }
    // 担当者別ランキングCSV出力（個別ファイル）
    async exportStaffRankingCSVs(report) {
        try {
            const monthText = `${report.selectedYear}年${report.selectedMonth}月`;
            // ①契約者70歳以上の受注件数トップ10ランキング
            if (report.elderlyStaffRanking && report.elderlyStaffRanking.length > 0) {
                const elderlyData = [
                    ['①契約者70歳以上の受注件数トップ10ランキング'],
                    [''],
                    ['ランキング', '地区No.', '所属No.', '担当名', '件数'],
                    ...report.elderlyStaffRanking.map((staff) => [
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
                    ...report.singleContractRanking.map((staff) => [
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
                    ...report.excessiveSalesRanking.map((staff) => [
                        staff.rank.toString(),
                        staff.regionNo,
                        staff.departmentNo,
                        staff.staffName,
                        staff.count.toString()
                    ])
                ];
                await this.downloadCSV(excessiveData, `③過量販売ランキング_${monthText}.csv`);
            }
            // ④69歳以下契約件数の担当別件数は除外（PDF出力と同様）
        }
        catch (error) {
            console.error('担当者別ランキングCSV出力エラー:', error);
        }
    }
    // CSVダウンロード用ヘルパーメソッド
    async downloadCSV(data, fileName) {
        const csvContent = data.map(row => row.map(cell => {
            const escapedCell = String(cell).replace(/"/g, '""');
            return `"${escapedCell}"`;
        }).join(',')).join('\r\n');
        const bom = '\uFEFF';
        const blob = new Blob([bom + csvContent], {
            type: 'text/csv;charset=utf-8;'
        });
        if (window.saveAs) {
            window.saveAs(blob, fileName);
        }
        else {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = fileName;
            link.click();
        }
    }
    createRankingTableHTML(ranking) {
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
                        ${ranking.map((staff) => `
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
    generateStaffData(data, targetDate) {
        const staffData = [];
        // 対象月のデータのみを抽出（受注カウント日付を基準に判定）
        const targetYear = targetDate.getFullYear();
        const targetMonth = targetDate.getMonth();
        const monthlyData = data.filter(row => {
            if (!row.date)
                return false;
            // 受注カウント日付を計算（A列とK列を比較して遅い日付を採用）
            const effectiveDate = this.calculateEffectiveDate(row);
            if (!effectiveDate)
                return false;
            // 受注カウント日付が対象月と一致するかチェック
            const matches = effectiveDate.getFullYear() === targetYear && effectiveDate.getMonth() === targetMonth;
            return matches;
        });
        monthlyData.forEach((row, index) => {
            // 担当者名の正規化
            const normalizedStaffName = this.normalizeStaffName(row.staffName);
            // 受注条件の判定（シンプル化）
            const isOrder = this.isSimpleOrder(row);
            // 年齢の取得
            const ageNum = this.getContractorAge(row);
            if (normalizedStaffName && row.regionNumber && row.departmentNumber) {
                const existingStaff = staffData.find(s => s.regionNo === row.regionNumber &&
                    s.departmentNo === row.departmentNumber &&
                    s.staffName === normalizedStaffName);
                if (existingStaff) {
                    // 受注件数のカウント
                    if (isOrder) {
                        existingStaff.totalOrders++;
                        // 年齢カウント
                        if (typeof ageNum === 'number') {
                            if (ageNum <= 69) {
                                existingStaff.normalAgeOrders++;
                            }
                            else if (ageNum >= 70) {
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
                }
                else {
                    const newStaff = {
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
                    staffData.push(newStaff);
                }
            }
        });
        return staffData;
    }
    // シンプルな受注判定
    isSimpleOrder(row) {
        // J列条件チェック（空欄または4の場合のみ受注としてカウント）
        const confirmation = row.confirmation;
        if (typeof confirmation === 'number') {
            // 数値の場合：4のみ受注としてカウント
            return confirmation === 4;
        }
        else if (typeof confirmation === 'string') {
            const trimmedValue = confirmation.trim();
            if (trimmedValue === '') {
                // 空欄の場合：受注としてカウント
                return true;
            }
            else {
                const confirmationNum = parseInt(trimmedValue);
                if (!isNaN(confirmationNum)) {
                    // 数値に変換できる場合：4のみ受注としてカウント
                    return confirmationNum === 4;
                }
            }
        }
        else {
            // null, undefined等の場合：空欄として扱い、受注としてカウント
            return true;
        }
        return false;
    }
    // 文字列/数値いずれでも年齢を数値で取得する
    getContractorAge(row) {
        let age = row?.contractorAge ?? row?.age;
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
    normalizeStaffName(staffName) {
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
    async exportStaffDataToCSV(staffData) {
        try {
            const reportMonthInput = document.getElementById('reportMonth');
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
        }
        catch (error) {
            console.error('担当別データCSV出力エラー:', error);
            alert('CSVの出力に失敗しました。');
        }
    }
    // 担当別データのHTMLを生成
    createStaffDataHTML(staffData) {
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
    /**
     * 受注カウント日付を計算（A列とK列を比較して遅い日付を採用）
     * @param row データ行
     * @returns 受注カウントに使用される日付
     */
    calculateEffectiveDate(row) {
        if (!row.date) {
            return null;
        }
        // 基本は受注日（A列）
        let effectiveDate = row.date;
        // K列から日付を抽出（8/30 10:03 大城 のような形式）
        if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
            const confirmationStr = row.confirmationDateTime;
            const dateTimePattern = confirmationStr.match(/(\d{1,2})\/(\d{1,2})/);
            if (dateTimePattern) {
                const kColumnMonth = parseInt(dateTimePattern[1]);
                const kColumnDay = parseInt(dateTimePattern[2]);
                const kColumnDate = new Date(row.date.getFullYear(), kColumnMonth - 1, kColumnDay);
                // 受注日と確認日を比較して、遅い日付を採用
                if (kColumnDate > row.date) {
                    effectiveDate = kColumnDate;
                }
            }
        }
        return effectiveDate;
    }
    // 確認データのHTMLを生成
    createDataConfirmationHTML(data) {
        if (!data || data.length === 0) {
            return '<div class="alert alert-info">データがありません</div>';
        }
        const tableRows = data.map((row, index) => {
            // 受注カウント日付を計算（A列とK列を比較して遅い日付を採用）
            const effectiveDate = this.calculateEffectiveDate(row);
            // 受注判定結果を計算（日報と同じロジック）
            const isOrderForDate = this.excelProcessor.isOrderForDate(row, effectiveDate || row.date, true);
            return `
                <tr data-order-status="${isOrderForDate ? 'order' : 'non-order'}" data-effective-date="${effectiveDate ? effectiveDate.toISOString() : ''}">
                    <td>${index + 1}</td>
                    <td>${row.date ? row.date.toLocaleDateString() : 'N/A'}</td>
                    <td>${effectiveDate ? effectiveDate.toLocaleDateString() : 'N/A'}</td>
                    <td>${row.staffName || ''}</td>
                    <td>${row.regionNumber || ''}</td>
                    <td>${row.departmentNumber || ''}</td>
                    <td>${row.contractor || ''}</td>
                    <td>${row.contractorAge || ''}</td>
                    <td>${row.confirmation || ''}</td>
                    <td>${row.confirmationDateTime || ''}</td>
                </tr>
            `;
        }).join('');
        return `
            <div class="alert alert-success">
                <h5>データ確認完了</h5>
                <p>総データ件数: <strong>${data.length}</strong>件</p>
            </div>
            
            <!-- フィルター機能 -->
            <div class="row mb-3">
                <div class="col-md-2">
                    <label for="staffFilter" class="form-label">担当者名</label>
                    <input type="text" class="form-control" id="staffFilter" placeholder="担当者名...">
                </div>
                <div class="col-md-2">
                    <label for="regionFilter" class="form-label">地区№</label>
                    <input type="text" class="form-control" id="regionFilter" placeholder="地区№...">
                </div>
                <div class="col-md-2">
                    <label for="departmentFilter" class="form-label">所属№</label>
                    <input type="text" class="form-control" id="departmentFilter" placeholder="所属№...">
                </div>
                <div class="col-md-2">
                    <label for="dateFilter" class="form-label">受注カウント日付</label>
                    <input type="date" class="form-control" id="dateFilter">
                </div>
                <div class="col-md-2">
                    <label for="monthFilter" class="form-label">受注カウント月</label>
                    <input type="month" class="form-control" id="monthFilter">
                </div>
                <div class="col-md-2">
                    <label class="form-label">&nbsp;</label>
                    <button type="button" class="btn btn-secondary w-100" id="clearFilters">フィルタークリア</button>
                </div>
            </div>
            
            <div class="table-responsive">
                <table class="table table-striped table-hover" id="dataConfirmationTable">
                    <thead class="table-dark">
                        <tr>
                            <th>#</th>
                            <th>日付（A列）</th>
                            <th>受注カウント日付</th>
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
    // 日報PDF用HTML（月報1枚目と同じレイアウト）
    createDailyReportPDFHTML(report, dateText) {
        // 月報の1枚目と完全に同じHTMLを生成（タイトルのみ「日報」に変更）
        const html = `
            <div class="pdf-page" data-page="1枚目: 基本統計">
                <div class="pdf-header">
                    <div class="pdf-title">日報</div>
                    <div class="pdf-date">${dateText}</div>
                </div>
                
                <!-- 基本統計 -->
                <div class="pdf-section">
                    <div class="pdf-section-title">総件数</div>
                    <div class="pdf-total-stats-container">
                        <div class="pdf-total-stats-grid">
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${report.totalOrders || 0}</div>
                                <div class="pdf-total-stat-label">受注件数</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${report.overtimeOrders || 0}</div>
                                <div class="pdf-total-stat-label">時間外対応</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${this.getTotalExcessive(report.regionStats || {}) || 0}</div>
                                <div class="pdf-total-stat-label">過量販売</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${this.getTotalSingle(report.regionStats || {}) || 0}</div>
                                <div class="pdf-total-stat-label">単独契約</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${this.getTotalHolidayConstruction(report.regionStats || {}) || 0}</div>
                                <div class="pdf-total-stat-label">公休日施工</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${this.getTotalProhibitedConstruction(report.regionStats || {}) || 0}</div>
                                <div class="pdf-total-stat-label">禁止日施工</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- 地区別受注件数 -->
                <div class="pdf-section">
                    <div class="pdf-section-title">地区別受注件数</div>
                    ${this.createRegionStatsPDFHTML(report.regionStats || {})}
                </div>
                
                <!-- 年齢別集計 -->
                <div class="pdf-section">
                    <div class="pdf-section-title">年齢別集計</div>
                    ${this.createAgeStatsPDFHTML(report.ageStats || {})}
                </div>
            </div>
        `;
        return html;
    }
    // PDF用地区統計HTML（2×2グリッド）
    createRegionStatsPDFHTML(regionStats) {
        let html = '';
        // データの存在確認
        if (!regionStats || Object.keys(regionStats).length === 0) {
            return '<div class="text-center text-muted py-3">地区データがありません</div>';
        }
        // 地区の順序を固定
        const regionOrder = ['九州地区', '中四国地区', '関西地区', '関東地区'];
        const availableRegions = regionOrder.filter(region => regionStats[region] && regionStats[region].orders > 0);
        if (availableRegions.length === 0) {
            return '<div class="text-center text-muted py-3">地区データがありません</div>';
        }
        // 1つの2×2グリッドレイアウトで4つの地区を表示
        html += '<div class="pdf-region-grid">';
        html += this.createRegionCardPDFHTML('九州地区', regionStats['九州地区']);
        html += this.createRegionCardPDFHTML('中四国地区', regionStats['中四国地区']);
        html += this.createRegionCardPDFHTML('関西地区', regionStats['関西地区']);
        html += this.createRegionCardPDFHTML('関東地区', regionStats['関東地区']);
        html += '</div>';
        return html;
    }
    // 地区カードのPDF用HTML生成（ヘルパーメソッド）
    createRegionCardPDFHTML(regionName, stats) {
        if (!stats) {
            return `<div class="pdf-region-card"><div class="pdf-region-name">${regionName}</div><div class="text-center text-muted">データなし</div></div>`;
        }
        return `
            <div class="pdf-region-card">
                <div class="pdf-region-name">${regionName}</div>
                <div class="pdf-region-stats">
                    <div class="pdf-region-stat-item">
                        <div class="pdf-region-stat-value">${stats.orders || 0}</div>
                        <div class="pdf-region-stat-label">受注件数</div>
                    </div>
                    <div class="pdf-region-stat-item">
                        <div class="pdf-region-stat-value">${stats.overtime || 0}</div>
                        <div class="pdf-region-stat-label">時間外対応</div>
                    </div>
                    <div class="pdf-region-stat-item">
                        <div class="pdf-region-stat-value">${stats.excessive || 0}</div>
                        <div class="pdf-region-stat-label">過量販売</div>
                    </div>
                    <div class="pdf-region-stat-item">
                        <div class="pdf-region-stat-value">${stats.single || 0}</div>
                        <div class="pdf-region-stat-label">単独契約</div>
                    </div>
                    <div class="pdf-region-stat-item">
                        <div class="pdf-region-stat-value">${stats.holidayConstruction || 0}</div>
                        <div class="pdf-region-stat-label">公休日施工</div>
                    </div>
                    <div class="pdf-region-stat-item">
                        <div class="pdf-region-stat-value">${stats.prohibitedConstruction || 0}</div>
                        <div class="pdf-region-stat-label">禁止日施工</div>
                    </div>
                </div>
            </div>
        `;
    }
    // PDF用年齢統計HTML
    createAgeStatsPDFHTML(ageStats) {
        if (!ageStats) {
            return '<div class="text-center text-muted py-3">年齢データがありません</div>';
        }
        const html = `
            <div class="pdf-age-grid">
                <div class="pdf-age-card">
                    <div class="pdf-age-title">高齢者（70歳以上）</div>
                    <div class="pdf-age-stats">
                        <div class="pdf-age-stat-item">
                            <div class="pdf-age-stat-value">${ageStats.elderly?.total || 0}</div>
                            <div class="pdf-age-stat-label">総件数</div>
                        </div>
                        <div class="pdf-age-stat-item">
                            <div class="pdf-age-stat-value">${ageStats.elderly?.excessive || 0}</div>
                            <div class="pdf-age-stat-label">過量販売</div>
                        </div>
                        <div class="pdf-age-stat-item">
                            <div class="pdf-age-stat-value">${ageStats.elderly?.single || 0}</div>
                            <div class="pdf-age-stat-label">単独契約</div>
                        </div>
                    </div>
                </div>
                
                <div class="pdf-age-card">
                    <div class="pdf-age-title">通常年齢（69歳以下）</div>
                    <div class="pdf-age-stats">
                        <div class="pdf-age-stat-item">
                            <div class="pdf-age-stat-value">${ageStats.normal?.total || 0}</div>
                            <div class="pdf-age-stat-label">総件数</div>
                        </div>
                        <div class="pdf-age-stat-item">
                            <div class="pdf-age-stat-value">${ageStats.normal?.excessive || 0}</div>
                            <div class="pdf-age-stat-label">過量販売</div>
                        </div>
                        <div class="pdf-age-stat-item">
                            <div class="pdf-age-stat-value">${ageStats.normal?.single || 0}</div>
                            <div class="pdf-age-stat-label">単独契約</div>
                        </div>
                    </div>
                </div>
            </div>
        `;
        return html;
    }
    // 月報PDF用HTML（4ページ構成）
    createMonthlyReportPDFHTML(report) {
        // 選択された月から年月を取得
        let monthText = '';
        if (report.selectedYear && report.selectedMonth) {
            monthText = `${report.selectedYear}年${report.selectedMonth}月`;
        }
        else {
            const now = new Date();
            monthText = `${now.getFullYear()}年${now.getMonth() + 1}月`;
        }
        let html = '';
        // 1枚目: 基本統計セクション
        const firstPage = `
            <div class="pdf-page" data-page="1枚目: 基本統計">
                <div class="pdf-header">
                    <div class="pdf-title">月報</div>
                    <div class="pdf-date">${monthText}</div>
                </div>
                
                <!-- 基本統計 -->
                <div class="pdf-section">
                    <div class="pdf-section-title">総件数</div>
                    <div class="pdf-total-stats-container">
                        <div class="pdf-total-stats-grid">
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${report.totalOrders || 0}</div>
                                <div class="pdf-total-stat-label">受注件数</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${report.overtimeOrders || 0}</div>
                                <div class="pdf-total-stat-label">時間外対応</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${this.getTotalExcessive(report.regionStats || {}) || 0}</div>
                                <div class="pdf-total-stat-label">過量販売</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${this.getTotalSingle(report.regionStats || {}) || 0}</div>
                                <div class="pdf-total-stat-label">単独契約</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${this.getTotalHolidayConstruction(report.regionStats || {}) || 0}</div>
                                <div class="pdf-total-stat-label">公休日施工</div>
                            </div>
                            <div class="pdf-total-stat-item">
                                <div class="pdf-total-stat-number">${this.getTotalProhibitedConstruction(report.regionStats || {}) || 0}</div>
                                <div class="pdf-total-stat-label">禁止日施工</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- 地区別受注件数 -->
                <div class="pdf-section">
                    <div class="pdf-section-title">地区別受注件数</div>
                    ${this.createRegionStatsPDFHTML(report.regionStats || {})}
                </div>
                
                <!-- 年齢別集計 -->
                <div class="pdf-section">
                    <div class="pdf-section-title">年齢別集計</div>
                    ${this.createAgeStatsPDFHTML(report.ageStats || {})}
                </div>
            </div>
        `;
        html += firstPage;
        // 2枚目: ①高齢者契約ランキング
        if (report.elderlyStaffRanking && report.elderlyStaffRanking.length > 0) {
            html += `
                <!-- ページ区切り -->
                <div class="pdf-page-break"></div>
                
                <!-- 2枚目: ①高齢者契約ランキング -->
                <div class="pdf-page" data-page="2枚目: ①高齢者契約ランキング">
                    <div class="pdf-header">
                        <div class="pdf-title">担当者別ランキング</div>
                        <div class="pdf-date">${monthText}</div>
                    </div>
                    
                    <div class="pdf-ranking-section">
                        <div class="pdf-ranking-title">①契約者70歳以上の受注件数トップ10ランキング</div>
                        ${this.createRankingTablePDFHTML(report.elderlyStaffRanking)}
                </div>
                </div>
            `;
        }
        else {
            html += `
                <!-- ページ区切り -->
                <div class="pdf-page-break"></div>
                
                <!-- 2枚目: ①高齢者契約ランキング -->
                <div class="pdf-page" data-page="2枚目: ①高齢者契約ランキング">
                    <div class="pdf-header">
                        <div class="pdf-title">担当者別ランキング</div>
                        <div class="pdf-date">${monthText}</div>
                </div>

                    <div class="pdf-ranking-section">
                        <div class="pdf-ranking-title">①契約者70歳以上の受注件数トップ10ランキング</div>
                        <div class="text-center text-muted py-3">データがありません</div>
                </div>
                </div>
            `;
        }
        // 3枚目: ②単独契約ランキング
        if (report.singleContractRanking && report.singleContractRanking.length > 0) {
            html += `
                <!-- ページ区切り -->
                <div class="pdf-page-break"></div>
                
                <!-- 3枚目: ②単独契約ランキング -->
                <div class="pdf-page" data-page="3枚目: ②単独契約ランキング">
                    <div class="pdf-header">
                        <div class="pdf-title">担当者別ランキング</div>
                        <div class="pdf-date">${monthText}</div>
                </div>
                    
                    <div class="pdf-ranking-section">
                        <div class="pdf-ranking-title">②単独契約を持っている担当者一覧</div>
                        ${this.createRankingTablePDFHTML(report.singleContractRanking)}
                    </div>
                </div>
            `;
        }
        else {
            html += `
                <!-- ページ区切り -->
                <div class="pdf-page-break"></div>
                
                <!-- 3枚目: ②単独契約ランキング -->
                <div class="pdf-page" data-page="3枚目: ②単独契約ランキング">
                    <div class="pdf-header">
                        <div class="pdf-title">担当者別ランキング</div>
                        <div class="pdf-date">${monthText}</div>
                    </div>
                    
                    <div class="pdf-ranking-section">
                        <div class="pdf-ranking-title">②単独契約を持っている担当者一覧</div>
                        <div class="text-center text-muted py-3">データがありません</div>
                    </div>
                </div>
            `;
        }
        // 4枚目: ③過量販売ランキング
        if (report.excessiveSalesRanking && report.excessiveSalesRanking.length > 0) {
            html += `
                <!-- ページ区切り -->
                <div class="pdf-page-break"></div>
                
                <!-- 4枚目: ③過量販売ランキング -->
                <div class="pdf-page" data-page="4枚目: ③過量販売ランキング">
                    <div class="pdf-header">
                        <div class="pdf-title">担当者別ランキング</div>
                        <div class="pdf-date">${monthText}</div>
                    </div>
                    
                    <div class="pdf-ranking-section">
                        <div class="pdf-ranking-title">③過量契約を持っている担当者一覧</div>
                        ${this.createRankingTablePDFHTML(report.excessiveSalesRanking)}
                    </div>
                </div>
            `;
        }
        else {
            html += `
                <!-- ページ区切り -->
                <div class="pdf-page-break"></div>
                
                <!-- 4枚目: ③過量販売ランキング -->
                <div class="pdf-page" data-page="4枚目: ③過量販売ランキング">
                    <div class="pdf-header">
                        <div class="pdf-title">担当者別ランキング</div>
                        <div class="pdf-date">${monthText}</div>
                    </div>
                    
                    <div class="pdf-ranking-section">
                        <div class="pdf-ranking-title">③過量契約を持っている担当者一覧</div>
                        <div class="text-center text-muted py-3">データがありません</div>
                    </div>
                </div>
            `;
        }
        return html;
    }
    // PDF用ランキングテーブルHTML
    createRankingTablePDFHTML(ranking) {
        if (!ranking || ranking.length === 0) {
            return '<div class="text-center text-muted py-3">データがありません</div>';
        }
        return `
            <div class="pdf-ranking-table">
                <div class="pdf-ranking-header">
                    <div class="pdf-ranking-cell">ランキング</div>
                    <div class="pdf-ranking-cell">地区No.</div>
                    <div class="pdf-ranking-cell">所属No.</div>
                    <div class="pdf-ranking-cell">担当名</div>
                    <div class="pdf-ranking-cell">件数</div>
                </div>
                ${ranking.map((staff) => `
                    <div class="pdf-ranking-row">
                        <div class="pdf-ranking-cell">${staff.rank || '-'}</div>
                        <div class="pdf-ranking-cell">${staff.regionNo || '-'}</div>
                        <div class="pdf-ranking-cell">${staff.departmentNo || '-'}</div>
                        <div class="pdf-ranking-cell">${staff.staffName || '-'}</div>
                        <div class="pdf-ranking-cell">${staff.count || 0}</div>
                    </div>
                `).join('')}
            </div>
        `;
    }
}
//# sourceMappingURL=report-generator.js.map