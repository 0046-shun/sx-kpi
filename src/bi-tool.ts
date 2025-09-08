import { ExcelProcessor } from './excel-processor';
import { CalendarManager } from './calendar-manager';

// Chart.jsの型定義
declare const Chart: any;

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
    regionOrders: { region: string; count: number }[];
    regionAmounts: { region: string; amount: number }[];
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

export class BITool {
    private excelProcessor: ExcelProcessor;
    private calendarManager: CalendarManager;
    private charts: Map<string, any> = new Map();

    constructor(excelProcessor: ExcelProcessor, calendarManager: CalendarManager) {
        this.excelProcessor = excelProcessor;
        this.calendarManager = calendarManager;
    }

    // 金額から「→」記号を処理して実際の金額を取得
    private parseAmount(value: any): number {
        if (!value) return 0;
        
        if (typeof value === 'number') {
            return value;
        }
        
        if (typeof value === 'string') {
            const trimmedValue = value.trim();
            
            // 「→」記号がある場合は後の金額を使用
            if (trimmedValue.includes('→')) {
                const parts = trimmedValue.split('→');
                if (parts.length >= 2) {
                    const afterAmount = parts[1].trim();
                    const parsed = parseFloat(afterAmount.replace(/[^\d.-]/g, ''));
                    return isNaN(parsed) ? 0 : parsed;
                }
            }
            
            // 通常の数値文字列の場合
            const parsed = parseFloat(trimmedValue.replace(/[^\d.-]/g, ''));
            return isNaN(parsed) ? 0 : parsed;
        }
        
        return 0;
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
        // 例: "山田(SE)" → "山田"
        // 例: "山田（技術）" → "山田"
        const normalizedName = trimmedName.replace(/[\(（].*?[\)）]/g, '');
        
        // 前後の空白を除去して返す
        return normalizedName.trim();
    }

    // 地区番号を地区名に変換
    private getRegionName(regionNumber: any): string {
        if (!regionNumber) return '未分類';
        
        const regionStr = String(regionNumber).trim();
        
        // 地区分類マッピング
        if (regionStr === '511') {
            return '九州地区';
        } else if (regionStr === '521' || regionStr === '531') {
            return '中四国地区';
        } else if (regionStr === '541') {
            return '関西地区';
        } else if (regionStr === '561') {
            return '関東地区';
        } else {
            return `地区${regionStr}`;
        }
    }

    // 日付に曜日を追加
    private formatDateWithDayOfWeek(date: Date): string {
        const month = date.getMonth() + 1;
        const day = date.getDate();
        const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
        return `${month}/${day}(${dayOfWeek})`;
    }

    // CalendarManagerから公休日設定を取得（重複除去）
    private getHolidaySettings(): any {
        try {
            const settings = this.calendarManager.getSettings();
            if (settings) {
                // 重複を除去
                const uniquePublicHolidays = settings.publicHolidays ? 
                    Array.from(new Set(settings.publicHolidays.map((d: Date) => d.toISOString().split('T')[0])))
                        .map(dateStr => new Date(dateStr)) : [];
                
                const uniqueProhibitedDays = settings.prohibitedDays ? 
                    Array.from(new Set(settings.prohibitedDays.map((d: Date) => d.toISOString().split('T')[0])))
                        .map(dateStr => new Date(dateStr)) : [];
                
                return {
                    publicHolidays: uniquePublicHolidays,
                    prohibitedDays: uniqueProhibitedDays
                };
            }
            return null;
        } catch (error) {
            console.error('公休日設定の取得に失敗しました:', error);
            return null;
        }
    }

    // 日付タイプの判定（日曜日・月曜日 + モーダル設定の公休日・禁止日を全て灰色で扱う）
    private getDayType(date: Date): { type: 'normal' | 'holiday', label: string, color: string } {
        const dayOfWeek = date.getDay(); // 0=日曜日, 1=月曜日, 2=火曜日, ...
        const targetDateStr = date.toISOString().split('T')[0];
        
        // 日曜日・月曜日は公休日
        if (dayOfWeek === 0 || dayOfWeek === 1) {
            return {
                type: 'holiday',
                label: '公休日',
                color: '#6c757d'  // 灰色
            };
        }
        
        // モーダルで設定された公休日・禁止日をチェック（どちらも灰色で扱う）
        const holidaySettings = this.getHolidaySettings();
        if (holidaySettings) {
            // 公休日をチェック
            if (holidaySettings.publicHolidays && Array.isArray(holidaySettings.publicHolidays)) {
                const isHoliday = holidaySettings.publicHolidays.some((holidayDate: Date) => {
                    const holidayDateStr = holidayDate.toISOString().split('T')[0];
                    return holidayDateStr === targetDateStr;
                });
                if (isHoliday) {
                    return {
                        type: 'holiday',
                        label: '公休日',
                        color: '#6c757d'  // 灰色
                    };
                }
            }
            
            // 禁止日をチェック（公休日と同じ灰色で扱う）
            if (holidaySettings.prohibitedDays && Array.isArray(holidaySettings.prohibitedDays)) {
                const isProhibited = holidaySettings.prohibitedDays.some((prohibitedDate: Date) => {
                    const prohibitedDateStr = prohibitedDate.toISOString().split('T')[0];
                    return prohibitedDateStr === targetDateStr;
                });
                if (isProhibited) {
                    return {
                        type: 'holiday',
                        label: '公休日',
                        color: '#6c757d'  // 灰色
                    };
                }
            }
        }
        
        return { type: 'normal', label: '通常日', color: '#4ecdc4' };
    }

    // BIダッシュボードのHTMLを生成
    public createBIDashboard(data: any[], targetMonth?: string): string {
        if (!data || data.length === 0) {
            return `
                <div class="alert alert-warning">
                    <h5><i class="fas fa-exclamation-triangle me-2"></i>データがありません</h5>
                    <p>BI分析を行うには、まずExcelファイルをアップロードしてデータを読み込んでください。</p>
                </div>
            `;
        }

        const stats = this.calculateStats(data, targetMonth);
        const monthLabel = targetMonth ? `${targetMonth}月` : '当月';
        
        return `
            <div class="bi-dashboard">
                <!-- 月選択 -->
                <div class="row mb-3">
                    <div class="col-md-12">
                        <div class="card">
                            <div class="card-body">
                                <div class="d-flex align-items-center">
                                    <label for="monthSelector" class="form-label me-3 mb-0">分析対象月:</label>
                                    <input type="month" class="form-control me-3" id="monthSelector" style="width: 200px;">
                                    <button class="btn btn-primary" id="updateBI">
                                        <i class="fas fa-sync me-1"></i>更新
                                    </button>
                                </div>
                                <div class="mt-2">
                                    <small class="text-muted">
                                        総データ件数: ${data.length}件 | 
                                        フィルタ後: ${stats.totalOrders}件 | 
                                        対象月: ${monthLabel}
                                    </small>
                                    ${stats.totalOrders === 0 ? `
                                        <button class="btn btn-sm btn-warning ms-2" id="showAllData">
                                            <i class="fas fa-eye me-1"></i>全データ表示
                                        </button>
                                    ` : ''}
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- 統計サマリー -->
                <div class="row mb-4">
                    <div class="col-md-4">
                        <div class="card bg-primary text-white">
                            <div class="card-body text-center">
                                <h3 class="card-title">${stats.totalOrders}</h3>
                                <p class="card-text">受注件数</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card bg-success text-white">
                            <div class="card-body text-center">
                                <h3 class="card-title">¥${stats.totalAmount.toLocaleString()}</h3>
                                <p class="card-text">受注金額</p>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card bg-info text-white">
                            <div class="card-body text-center">
                                <h3 class="card-title">¥${stats.averageAmount.toLocaleString()}</h3>
                                <p class="card-text">平均金額</p>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- グラフエリア -->
                <div class="row mb-4">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <h5 class="mb-0"><i class="fas fa-chart-pie me-2"></i>地区別受注件数</h5>
                            </div>
                            <div class="card-body">
                                <canvas id="regionOrdersChart" width="400" height="300"></canvas>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <h5 class="mb-0"><i class="fas fa-chart-bar me-2"></i>地区別受注金額</h5>
                            </div>
                            <div class="card-body">
                                <canvas id="regionAmountsChart" width="400" height="300"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- 推移表 -->
                <div class="row mb-4">
                    <div class="col-md-12">
                        <div class="card">
                            <div class="card-header">
                                <h5 class="mb-0"><i class="fas fa-chart-line me-2"></i>${monthLabel}受注件数推移</h5>
                            </div>
                            <div class="card-body">
                                <canvas id="dailyTrendChart" width="800" height="300"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- 統計ランキング表 -->
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <h5 class="mb-0"><i class="fas fa-trophy me-2"></i>担当別受注件数ランキング</h5>
                            </div>
                            <div class="card-body">
                                <table class="table table-sm">
                                    <thead>
                                        <tr>
                                            <th>順位</th>
                                            <th>地区</th>
                                            <th>所属No.</th>
                                            <th>担当名</th>
                                            <th>受注件数</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${stats.staffOrderRanking.map((staff, index) => `
                                            <tr>
                                                <td>${index + 1}</td>
                                                <td>${staff.regionName}</td>
                                                <td>${staff.departmentNo}</td>
                                                <td>${staff.name}</td>
                                                <td>${staff.count}件</td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <h5 class="mb-0"><i class="fas fa-trophy me-2"></i>担当別受注金額ランキング</h5>
                            </div>
                            <div class="card-body">
                                <table class="table table-sm">
                                    <thead>
                                        <tr>
                                            <th>順位</th>
                                            <th>地区</th>
                                            <th>所属No.</th>
                                            <th>担当名</th>
                                            <th>受注金額</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${stats.staffAmountRanking.map((staff, index) => `
                                            <tr>
                                                <td>${index + 1}</td>
                                                <td>${staff.regionName}</td>
                                                <td>${staff.departmentNo}</td>
                                                <td>${staff.name}</td>
                                                <td>¥${staff.amount.toLocaleString()}</td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }

    // 統計データを計算
    private calculateStats(data: any[], targetMonth?: string): BIStats {
        const stats: BIStats = {
            totalOrders: 0,
            totalAmount: 0,
            averageAmount: 0,
            regionOrders: [],
            regionAmounts: [],
            dailyTrend: [],
            staffOrderRanking: [],
            staffAmountRanking: []
        };

        // 対象月のデータをフィルタリング
        let filteredData = data;
        if (targetMonth) {
            const [year, month] = targetMonth.split('-');
            filteredData = data.filter(row => {
                // 受注日（A列）または確認日（J列）から日付を取得
                let dateToCheck = null;
                
                // A列の受注日をチェック
                if (row.date) {
                    dateToCheck = new Date(row.date);
                }
                // K列の確認者日時をチェック
                else if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
                    const dateMatch = row.confirmationDateTime.match(/(\d{1,2})\/(\d{1,2})/);
                    if (dateMatch) {
                        const monthNum = parseInt(dateMatch[1]);
                        const dayNum = parseInt(dateMatch[2]);
                        const yearNum = parseInt(year); // 対象年を使用
                        dateToCheck = new Date(yearNum, monthNum - 1, dayNum);
                    }
                }
                // J列の確認から日付を抽出（フォールバック）
                else if (row.confirmation && typeof row.confirmation === 'string') {
                    const dateMatch = row.confirmation.match(/(\d{1,2})\/(\d{1,2})/);
                    if (dateMatch) {
                        const monthNum = parseInt(dateMatch[1]);
                        const dayNum = parseInt(dateMatch[2]);
                        const yearNum = parseInt(year); // 対象年を使用
                        dateToCheck = new Date(yearNum, monthNum - 1, dayNum);
                    }
                }
                
                if (!dateToCheck) return false;
                
                return dateToCheck.getFullYear() === parseInt(year) && 
                       dateToCheck.getMonth() === parseInt(month) - 1;
            });
        }

        // 基本統計の計算
        const staffCount: Map<string, number> = new Map();
        const staffAmount: Map<string, number> = new Map();
        const regionCount: Map<string, number> = new Map();
        const regionAmount: Map<string, number> = new Map();
        const dailyCount: Map<string, number> = new Map();

        filteredData.forEach(row => {
            // 金額の集計（「→」記号を処理）
            const parsedAmount = this.parseAmount(row.amount);
            stats.totalAmount += parsedAmount;

            // 担当者別集計（複合キー：地区№+所属№+担当者名を使用）
            if (row.staffName && row.regionNumber && row.departmentNumber) {
                const normalizedStaffName = this.normalizeStaffName(row.staffName);
                if (normalizedStaffName) {
                    // 複合キーを作成（地区№|所属№|担当者名）
                    const compositeKey = `${row.regionNumber}|${row.departmentNumber}|${normalizedStaffName}`;
                    
                    const count = staffCount.get(compositeKey) || 0;
                    staffCount.set(compositeKey, count + 1);
                    
                    const amount = staffAmount.get(compositeKey) || 0;
                    staffAmount.set(compositeKey, amount + parsedAmount);
                }
            }

            // 地区別集計
            if (row.regionNumber) {
                const regionName = this.getRegionName(row.regionNumber);
                const count = regionCount.get(regionName) || 0;
                regionCount.set(regionName, count + 1);
                
                const amount = regionAmount.get(regionName) || 0;
                regionAmount.set(regionName, amount + parsedAmount);
            }

            // 日別集計
            let dateForDaily = null;
            
            // A列の受注日をチェック
            if (row.date) {
                dateForDaily = new Date(row.date);
            }
            // K列の確認者日時をチェック
            else if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
                const dateMatch = row.confirmationDateTime.match(/(\d{1,2})\/(\d{1,2})/);
                if (dateMatch) {
                    const monthNum = parseInt(dateMatch[1]);
                    const dayNum = parseInt(dateMatch[2]);
                    const yearNum = targetMonth ? parseInt(targetMonth.split('-')[0]) : new Date().getFullYear();
                    dateForDaily = new Date(yearNum, monthNum - 1, dayNum);
                }
            }
            // J列の確認から日付を抽出（フォールバック）
            else if (row.confirmation && typeof row.confirmation === 'string') {
                const dateMatch = row.confirmation.match(/(\d{1,2})\/(\d{1,2})/);
                if (dateMatch) {
                    const monthNum = parseInt(dateMatch[1]);
                    const dayNum = parseInt(dateMatch[2]);
                    const yearNum = targetMonth ? parseInt(targetMonth.split('-')[0]) : new Date().getFullYear();
                    dateForDaily = new Date(yearNum, monthNum - 1, dayNum);
                }
            }
            
            if (dateForDaily) {
                const dayKey = `${dateForDaily.getMonth() + 1}/${dateForDaily.getDate()}`;
                const count = dailyCount.get(dayKey) || 0;
                dailyCount.set(dayKey, count + 1);
            }
        });

        stats.totalOrders = filteredData.length;
        stats.averageAmount = stats.totalOrders > 0 ? Math.round(stats.totalAmount / stats.totalOrders) : 0;

        // 地区別受注件数ランキング
        stats.regionOrders = Array.from(regionCount.entries())
            .map(([region, count]) => ({ region, count }))
            .sort((a, b) => b.count - a.count);

        // 地区別受注金額ランキング
        stats.regionAmounts = Array.from(regionAmount.entries())
            .map(([region, amount]) => ({ region, amount }))
            .sort((a, b) => b.amount - a.amount);

        // 日別推移（該当月の1日～月末）
        const targetDate = targetMonth ? new Date(targetMonth + '-01') : new Date();
        const year = targetDate.getFullYear();
        const month = targetDate.getMonth();
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        
        for (let day = 1; day <= daysInMonth; day++) {
            const currentDate = new Date(year, month, day);
            const dayKey = `${month + 1}/${day}`;
            const dayType = this.getDayType(currentDate);
            
            
            stats.dailyTrend.push({
                day: this.formatDateWithDayOfWeek(currentDate),
                count: dailyCount.get(dayKey) || 0,
                dayType: dayType.type,
                dayLabel: dayType.label,
                dayColor: dayType.color
            });
        }

        // 担当者別受注件数ランキング（トップ10）
        stats.staffOrderRanking = Array.from(staffCount.entries())
            .map(([compositeKey, count]) => {
                // 複合キーから地区・所属・担当者名を抽出（地区№|所属№|担当者名）
                const parts = compositeKey.split('|');
                const regionNo = parts[0] || '';
                const departmentNo = parts[1] || '';
                const staffName = parts[2] || compositeKey; // フォールバック
                const regionName = this.getRegionName(regionNo);
                return { 
                    name: staffName, 
                    count, 
                    regionNo, 
                    departmentNo, 
                    regionName 
                };
            })
            .sort((a, b) => b.count - a.count)
            .slice(0, 10);

        // 担当者別受注金額ランキング（トップ10）
        stats.staffAmountRanking = Array.from(staffAmount.entries())
            .map(([compositeKey, amount]) => {
                // 複合キーから地区・所属・担当者名を抽出（地区№|所属№|担当者名）
                const parts = compositeKey.split('|');
                const regionNo = parts[0] || '';
                const departmentNo = parts[1] || '';
                const staffName = parts[2] || compositeKey; // フォールバック
                const regionName = this.getRegionName(regionNo);
                return { 
                    name: staffName, 
                    amount, 
                    regionNo, 
                    departmentNo, 
                    regionName 
                };
            })
            .sort((a, b) => b.amount - a.amount)
            .slice(0, 10);

        return stats;
    }

    // グラフを描画
    public renderCharts(data: any[], targetMonth?: string): void {
        const stats = this.calculateStats(data, targetMonth);
        
        // 既存のチャートを破棄
        this.charts.forEach(chart => chart.destroy());
        this.charts.clear();

        // 地区別受注件数円グラフ
        this.renderRegionOrdersChart(stats);
        
        // 地区別受注金額棒グラフ
        this.renderRegionAmountsChart(stats);
        
        // 日別推移線グラフ
        this.renderDailyTrendChart(stats);
    }

    private renderRegionOrdersChart(stats: BIStats): void {
        const ctx = document.getElementById('regionOrdersChart') as HTMLCanvasElement;
        if (!ctx) return;

        const topRegions = stats.regionOrders.slice(0, 8);
        const colors = [
            'rgba(255, 99, 132, 0.8)',
            'rgba(54, 162, 235, 0.8)',
            'rgba(255, 205, 86, 0.8)',
            'rgba(75, 192, 192, 0.8)',
            'rgba(153, 102, 255, 0.8)',
            'rgba(255, 159, 64, 0.8)',
            'rgba(199, 199, 199, 0.8)',
            'rgba(83, 102, 255, 0.8)'
        ];

        const chart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: topRegions.map(r => r.region),
                datasets: [{
                    data: topRegions.map(r => r.count),
                    backgroundColor: colors.slice(0, topRegions.length),
                    borderColor: colors.slice(0, topRegions.length).map(color => color.replace('0.8', '1')),
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom'
                    }
                }
            }
        });

        this.charts.set('regionOrders', chart);
    }

    private renderRegionAmountsChart(stats: BIStats): void {
        const ctx = document.getElementById('regionAmountsChart') as HTMLCanvasElement;
        if (!ctx) return;

        const topRegions = stats.regionAmounts.slice(0, 8);
        
        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: topRegions.map(r => r.region),
                datasets: [{
                    label: '受注金額',
                    data: topRegions.map(r => r.amount),
                    backgroundColor: 'rgba(54, 162, 235, 0.8)',
                    borderColor: 'rgba(54, 162, 235, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value: any) {
                                return '¥' + value.toLocaleString();
                            }
                        }
                    }
                }
            }
        });

        this.charts.set('regionAmounts', chart);
    }

    private renderDailyTrendChart(stats: BIStats): void {
        const ctx = document.getElementById('dailyTrendChart') as HTMLCanvasElement;
        if (!ctx) return;

        // 公休日の背景色を設定（公休日・禁止日は区別せず灰色）
                const backgroundColors = stats.dailyTrend.map(d => {
                    if (d.dayType === 'holiday') {
                        return 'rgba(108, 117, 125, 0.3)'; // 公休日・禁止日：薄い灰
                    } else {
                        return 'rgba(75, 192, 192, 0.2)'; // 通常日：薄い青緑
                    }
                });

        const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: stats.dailyTrend.map(d => d.day),
                datasets: [{
                    label: '受注件数',
                    data: stats.dailyTrend.map(d => d.count),
                    borderColor: 'rgba(75, 192, 192, 1)',
                    backgroundColor: backgroundColors,
                    tension: 0.1,
                    fill: true,
                    pointBackgroundColor: stats.dailyTrend.map(d => {
                        return d.dayColor || 'rgba(75, 192, 192, 1)';
                    }),
                    pointBorderColor: stats.dailyTrend.map(d => {
                        const color = d.dayColor || 'rgba(75, 192, 192, 1)';
                        return color;
                    }),
                    pointRadius: 4,
                    pointHoverRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    tooltip: {
                        callbacks: {
                            title: function(context: any) {
                                const dataIndex = context[0].dataIndex;
                                const dayData = stats.dailyTrend[dataIndex];
                                return `${dayData.day} (${dayData.dayLabel || '通常日'})`;
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true
                    },
                    x: {
                        ticks: {
                            maxRotation: 45,
                            minRotation: 45
                        }
                    }
                }
            }
        });

        this.charts.set('dailyTrend', chart);
    }

    // チャートを破棄
    public destroyCharts(): void {
        this.charts.forEach(chart => chart.destroy());
        this.charts.clear();
    }

}
