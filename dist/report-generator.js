export class ReportGenerator {
    constructor() {
        this.currentTargetDate = null;
        this.holidaySettings = {
            publicHolidays: [],
            prohibitedDays: []
        };
    }
    // 公休日・禁止日設定を更新
    updateHolidaySettings(settings) {
        this.holidaySettings = settings;
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
        return this.holidaySettings.publicHolidays.some(holiday => this.isSameDate(date, holiday));
    }
    // 禁止日かどうかの判定
    isProhibitedDay(date) {
        return this.holidaySettings.prohibitedDays.some(prohibited => this.isSameDate(date, prohibited));
    }
    generateDailyReport(data, date) {
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
            }
            else if (typeof confirmationValue === 'string') {
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
    generateMonthlyReport(data, month) {
        const [year, monthNum] = month.split('-').map(Number);
        const monthlyData = data.filter(row => {
            // 日付チェック
            if (!row.date)
                return false;
            const isDateMatch = row.date.getFullYear() === year && row.date.getMonth() === monthNum - 1;
            if (!isDateMatch)
                return false;
            // J列条件チェック（1、2、5の場合は除外）
            let isJColumnValid = true;
            const confirmationValue = row.confirmation;
            if (typeof confirmationValue === 'number') {
                if (confirmationValue === 1 || confirmationValue === 2 || confirmationValue === 5) {
                    isJColumnValid = false;
                }
            }
            else if (typeof confirmationValue === 'string') {
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
        reportData.selectedYear = year;
        // 担当者別ランキング集計
        reportData.elderlyStaffRanking = this.calculateElderlyStaffRanking(monthlyData);
        reportData.singleContractRanking = this.calculateSingleContractRanking(monthlyData);
        reportData.excessiveSalesRanking = this.calculateExcessiveSalesRanking(monthlyData);
        reportData.normalAgeStaffRanking = this.calculateNormalAgeStaffRanking(monthlyData);
        return reportData;
    }
    // 時間外カウント取得（AB列とK列を独立してカウント）
    getOvertimeCount(row, targetDate) {
        if (!targetDate)
            return 0;
        const targetMonth = targetDate.getMonth();
        const targetDay = targetDate.getDate();
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
            if (row.isElderly) {
                stats.elderly.total++;
                if (row.isExcessive)
                    stats.elderly.excessive++;
                if (row.isSingle)
                    stats.elderly.single++;
            }
            else {
                stats.normal.total++;
                if (row.isExcessive)
                    stats.normal.excessive++;
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
        // 基本的な条件チェック
        if (!row.staffName || row.staffName.trim() === '') {
            return false;
        }
        // 日付が存在するかチェック
        if (!row.date) {
            return false;
        }
        // 簡素化された受注判定（一時的）
        return true;
    }
    // 担当者別ランキング集計メソッド
    // 契約者70歳以上の受注件数トップ10ランキング
    calculateElderlyStaffRanking(data) {
        console.log('calculateElderlyStaffRanking 開始 - データ件数:', data.length);
        const staffCounts = new Map();
        data.forEach((row, index) => {
            // 動的にisOrderを計算
            const isOrder = this.isOrderForDate(row, new Date());
            const age = row.contractorAge || row.age;
            if (index < 10) {
                console.log(`高齢者 行${index}:`, {
                    staffName: row.staffName,
                    age: age,
                    isOrder: isOrder,
                    isElderly: age && age >= 70
                });
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
    calculateSingleContractRanking(data) {
        console.log('calculateSingleContractRanking 開始 - データ件数:', data.length);
        const staffCounts = new Map();
        data.forEach((row, index) => {
            // 動的にisOrderを計算
            const isOrder = this.isOrderForDate(row, new Date());
            if (index < 10) {
                console.log(`単独契約 行${index}:`, {
                    staffName: row.staffName,
                    isSingle: row.isSingle,
                    isOrder: isOrder
                });
            }
            // 条件: 単独契約 AND 受注 AND 担当者名が存在
            if (row.isSingle && isOrder && row.staffName && row.staffName.trim() !== '') {
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
    calculateExcessiveSalesRanking(data) {
        console.log('calculateExcessiveSalesRanking 開始 - データ件数:', data.length);
        const staffCounts = new Map();
        data.forEach((row, index) => {
            // 動的にisOrderを計算
            const isOrder = this.isOrderForDate(row, new Date());
            if (index < 10) {
                console.log(`過量販売 行${index}:`, {
                    staffName: row.staffName,
                    isExcessive: row.isExcessive,
                    isOrder: isOrder
                });
            }
            // 条件: 過量販売 AND 受注 AND 担当者名が存在
            if (row.isExcessive && isOrder && row.staffName && row.staffName.trim() !== '') {
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
    calculateNormalAgeStaffRanking(data) {
        console.log('calculateNormalAgeStaffRanking 開始 - データ件数:', data.length);
        const staffCounts = new Map();
        data.forEach((row, index) => {
            // 動的にisOrderを計算
            const isOrder = this.isOrderForDate(row, new Date());
            const age = row.contractorAge || row.age;
            const isNormalAge = !age || age < 70;
            if (index < 10) {
                console.log(`69歳以下 行${index}:`, {
                    staffName: row.staffName,
                    age: age,
                    isNormalAge: isNormalAge,
                    isOrder: isOrder
                });
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

                <!-- 担当者別ランキング -->
                ${report.elderlyStaffRanking ? `
                <div class="mb-4">
                    <h5 class="mb-3"><i class="fas fa-trophy me-2"></i>担当者別ランキング</h5>
                    
                    <!-- ①契約者70歳以上の受注件数トップ10ランキング -->
                    <div class="ranking-section mb-4">
                        <h6 class="ranking-title">①契約者70歳以上の受注件数トップ10ランキング</h6>
                        ${this.createRankingTableHTML(report.elderlyStaffRanking)}
                    </div>

                    <!-- ②単独契約ランキング -->
                    <div class="ranking-section mb-4">
                        <h6 class="ranking-title">②単独契約ランキング</h6>
                        ${this.createRankingTableHTML(report.singleContractRanking)}
                    </div>

                    <!-- ③過量販売ランキング -->
                    <div class="ranking-section mb-4">
                        <h6 class="ranking-title">③過量販売ランキング</h6>
                        ${this.createRankingTableHTML(report.excessiveSalesRanking)}
                    </div>

                    <!-- ④69歳以下契約件数の担当別件数 -->
                    <div class="ranking-section mb-4">
                        <h6 class="ranking-title">④69歳以下契約件数の担当別件数</h6>
                        ${this.createRankingTableHTML(report.normalAgeStaffRanking)}
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
                        <h6 class="ranking-title">①契約者70歳以上の受注件数トップ10ランキング</h6>
                        ${this.createRankingTableHTML(report.elderlyStaffRanking)}
                    </div>

                    <!-- ②単独契約ランキング -->
                    <div class="ranking-section mb-4">
                        <h6 class="ranking-title">②単独契約ランキング</h6>
                        ${this.createRankingTableHTML(report.singleContractRanking)}
                    </div>

                    <!-- ③過量販売ランキング -->
                    <div class="ranking-section mb-4">
                        <h6 class="ranking-title">③過量販売ランキング</h6>
                        ${this.createRankingTableHTML(report.excessiveSalesRanking)}
                    </div>

                    <!-- ④69歳以下契約件数の担当別件数 -->
                    <div class="ranking-section mb-4">
                        <h6 class="ranking-title">④69歳以下契約件数の担当別件数</h6>
                        ${this.createRankingTableHTML(report.normalAgeStaffRanking)}
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
            // PDF用のHTML要素を動的に作成
            const pdfContainer = this.createPDFHTML(report, type);
            document.body.appendChild(pdfContainer);
            // html2canvasでHTML要素を画像に変換
            const canvas = await window.html2canvas(pdfContainer, {
                scale: 3, // より高画質化
                useCORS: true,
                allowTaint: true,
                backgroundColor: '#ffffff',
                logging: false
            });
            // jsPDFでPDF作成
            const { jsPDF } = window.jspdf;
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
            }
            else {
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
        }
        catch (error) {
            console.error('PDF出力エラー:', error);
            alert('PDFの出力に失敗しました。');
        }
    }
    createPDFHTML(report, type) {
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
            .filter(([_, stats]) => stats.orders > 0)
            .map(([region, stats]) => `
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

            <!-- 担当者別ランキング（PDF対象：①②③のみ） -->
            <div class="pdf-section">
                <div class="pdf-section-title">担当者別ランキング</div>
                
                <!-- ①契約者70歳以上の受注件数トップ10ランキング -->
                <div class="pdf-ranking-section">
                    <div class="pdf-ranking-title">①契約者70歳以上の受注件数トップ10ランキング</div>
                    <div class="pdf-ranking-table">
                        <div class="pdf-ranking-header">
                            <div class="pdf-ranking-cell">ランキング</div>
                            <div class="pdf-ranking-cell">地区No.</div>
                            <div class="pdf-ranking-cell">所属No.</div>
                            <div class="pdf-ranking-cell">担当名</div>
                            <div class="pdf-ranking-cell">件数</div>
                        </div>
                        ${report.elderlyStaffRanking && report.elderlyStaffRanking.length > 0 ?
            report.elderlyStaffRanking.map((staff) => `
                                <div class="pdf-ranking-row">
                                    <div class="pdf-ranking-cell">${staff.rank}</div>
                                    <div class="pdf-ranking-cell">${staff.regionNo}</div>
                                    <div class="pdf-ranking-cell">${staff.departmentNo}</div>
                                    <div class="pdf-ranking-cell">${staff.staffName}</div>
                                    <div class="pdf-ranking-cell">${staff.count}</div>
                                </div>
                            `).join('') :
            '<div class="pdf-ranking-row"><div class="pdf-ranking-cell" colspan="5">データがありません</div></div>'}
                    </div>
                </div>

                <!-- ②単独契約ランキング -->
                <div class="pdf-ranking-section">
                    <div class="pdf-ranking-title">②単独契約ランキング</div>
                    <div class="pdf-ranking-table">
                        <div class="pdf-ranking-header">
                            <div class="pdf-ranking-cell">ランキング</div>
                            <div class="pdf-ranking-cell">地区No.</div>
                            <div class="pdf-ranking-cell">所属No.</div>
                            <div class="pdf-ranking-cell">担当名</div>
                            <div class="pdf-ranking-cell">件数</div>
                        </div>
                        ${report.singleContractRanking && report.singleContractRanking.length > 0 ?
            report.singleContractRanking.map((staff) => `
                                <div class="pdf-ranking-row">
                                    <div class="pdf-ranking-cell">${staff.rank}</div>
                                    <div class="pdf-ranking-cell">${staff.regionNo}</div>
                                    <div class="pdf-ranking-cell">${staff.departmentNo}</div>
                                    <div class="pdf-ranking-cell">${staff.staffName}</div>
                                    <div class="pdf-ranking-cell">${staff.count}</div>
                                </div>
                            `).join('') :
            '<div class="pdf-ranking-row"><div class="pdf-ranking-cell" colspan="5">データがありません</div></div>'}
                    </div>
                </div>

                <!-- ③過量販売ランキング -->
                <div class="pdf-ranking-section">
                    <div class="pdf-ranking-title">③過量販売ランキング</div>
                    <div class="pdf-ranking-table">
                        <div class="pdf-ranking-header">
                            <div class="pdf-ranking-cell">ランキング</div>
                            <div class="pdf-ranking-cell">地区No.</div>
                            <div class="pdf-ranking-cell">所属No.</div>
                            <div class="pdf-ranking-cell">担当名</div>
                            <div class="pdf-ranking-cell">件数</div>
                        </div>
                        ${report.excessiveSalesRanking && report.excessiveSalesRanking.length > 0 ?
            report.excessiveSalesRanking.map((staff) => `
                                <div class="pdf-ranking-row">
                                    <div class="pdf-ranking-cell">${staff.rank}</div>
                                    <div class="pdf-ranking-cell">${staff.regionNo}</div>
                                    <div class="pdf-ranking-cell">${staff.departmentNo}</div>
                                    <div class="pdf-ranking-cell">${staff.staffName}</div>
                                    <div class="pdf-ranking-cell">${staff.count}</div>
                                </div>
                            `).join('') :
            '<div class="pdf-ranking-row"><div class="pdf-ranking-cell" colspan="5">データがありません</div></div>'}
                    </div>
                </div>
            </div>
        `;
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
            // ④69歳以下契約件数の担当別件数
            if (report.normalAgeStaffRanking && report.normalAgeStaffRanking.length > 0) {
                const normalData = [
                    ['④69歳以下契約件数の担当別件数'],
                    [''],
                    ['ランキング', '地区No.', '所属No.', '担当名', '件数'],
                    ...report.normalAgeStaffRanking.map((staff) => [
                        staff.rank.toString(),
                        staff.regionNo,
                        staff.departmentNo,
                        staff.staffName,
                        staff.count.toString()
                    ])
                ];
                await this.downloadCSV(normalData, `④69歳以下契約件数担当別_${monthText}.csv`);
            }
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
    generateStaffData(data) {
        console.log('generateStaffData開始, データ件数:', data.length);
        const staffMap = new Map();
        // データの詳細を確認
        let orderCount = 0;
        let staffNameCount = 0;
        let validStaffCount = 0;
        data.forEach((row, index) => {
            // 動的にisOrderを計算
            const isOrder = this.isOrderForDate(row, new Date());
            row.isOrder = isOrder; // 結果を保存
            // 受注件数のカウント
            if (isOrder) {
                orderCount++;
            }
            // 担当者名がある件数のカウント
            if (row.staffName && row.staffName.trim() !== '') {
                staffNameCount++;
            }
            // 受注かつ担当者名がある件数のカウント
            if (isOrder && row.staffName && row.staffName.trim() !== '') {
                validStaffCount++;
                if (index < 5) { // 最初の5件のデータをログ出力
                    console.log(`有効な行${index}:`, {
                        isOrder: row.isOrder,
                        staffName: row.staffName,
                        departmentNumber: row.departmentNumber,
                        regionNumber: row.regionNumber,
                        age: row.age,
                        isSingle: row.isSingle,
                        isExcessive: row.isExcessive,
                        isOvertime: row.isOvertime
                    });
                }
                const key = `${row.regionNumber || ''}-${row.departmentNumber || ''}-${row.staffName}`;
                if (!staffMap.has(key)) {
                    staffMap.set(key, {
                        regionNo: row.regionNumber || '',
                        departmentNo: row.departmentNumber || '',
                        staffName: row.staffName,
                        totalOrders: 0,
                        normalAgeOrders: 0,
                        elderlyOrders: 0,
                        singleOrders: 0,
                        excessiveOrders: 0,
                        overtimeOrders: 0
                    });
                }
                const staff = staffMap.get(key);
                staff.totalOrders++;
                // 年齢による分類（contractorAgeを使用）
                const age = row.contractorAge || row.age;
                if (age && age >= 70) {
                    staff.elderlyOrders++;
                }
                else {
                    staff.normalAgeOrders++;
                }
                // その他の分類
                if (row.isSingle)
                    staff.singleOrders++;
                if (row.isExcessive)
                    staff.excessiveOrders++;
                if (row.isOvertime)
                    staff.overtimeOrders++;
            }
        });
        console.log('データ分析結果:', {
            totalRows: data.length,
            orderRows: orderCount,
            staffNameRows: staffNameCount,
            validStaffRows: validStaffCount
        });
        console.log('担当者別集計完了, 担当者数:', staffMap.size);
        // 件数降順でソート
        const result = Array.from(staffMap.values()).sort((a, b) => b.totalOrders - a.totalOrders);
        console.log('担当別データ結果(最初の3件):', result.slice(0, 3));
        return result;
    }
    // 担当別データのHTMLを生成
    createStaffDataHTML(staffData) {
        if (!staffData || staffData.length === 0) {
            return '<div class="text-center text-muted py-3">データがありません</div>';
        }
        return `
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead class="table-dark">
                        <tr>
                            <th>地区</th>
                            <th>所属</th>
                            <th>担当名</th>
                            <th>受注件数</th>
                            <th>69歳以下件数</th>
                            <th>70歳以上件数</th>
                            <th>単独件数</th>
                            <th>過量件数</th>
                            <th>時間外件数</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${staffData.map(staff => `
                            <tr>
                                <td>${staff.regionNo}</td>
                                <td>${staff.departmentNo}</td>
                                <td><strong>${staff.staffName}</strong></td>
                                <td><span class="badge bg-primary">${staff.totalOrders}</span></td>
                                <td><span class="badge bg-success">${staff.normalAgeOrders}</span></td>
                                <td><span class="badge bg-warning">${staff.elderlyOrders}</span></td>
                                <td><span class="badge bg-info">${staff.singleOrders}</span></td>
                                <td><span class="badge bg-danger">${staff.excessiveOrders}</span></td>
                                <td><span class="badge bg-secondary">${staff.overtimeOrders}</span></td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        `;
    }
}
//# sourceMappingURL=report-generator.js.map