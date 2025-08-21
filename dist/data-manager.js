export class DataManager {
    constructor() {
        this.STORAGE_KEY = 'daily_monthly_report_data';
    }
    saveData(data) {
        try {
            const dataToSave = {
                timestamp: new Date().toISOString(),
                data: data
            };
            localStorage.setItem(this.STORAGE_KEY, JSON.stringify(dataToSave));
        }
        catch (error) {
            console.error('データ保存エラー:', error);
        }
    }
    loadData() {
        try {
            const saved = localStorage.getItem(this.STORAGE_KEY);
            if (saved) {
                const parsed = JSON.parse(saved);
                // データの有効期限チェック（24時間）
                const savedTime = new Date(parsed.timestamp);
                const now = new Date();
                const hoursDiff = (now.getTime() - savedTime.getTime()) / (1000 * 60 * 60);
                if (hoursDiff < 24) {
                    return parsed.data;
                }
                else {
                    // 古いデータを削除
                    this.clearData();
                }
            }
        }
        catch (error) {
            console.error('データ読み込みエラー:', error);
        }
        return null;
    }
    clearData() {
        try {
            localStorage.removeItem(this.STORAGE_KEY);
        }
        catch (error) {
            console.error('データ削除エラー:', error);
        }
    }
    createDataTableHTML(data, filterState) {
        if (!data || data.length === 0) {
            return '<div class="text-center text-muted py-5">データがありません</div>';
        }
        // 表示する列を制限（重要な列のみ）
        const displayColumns = [
            { key: 'date', label: '日付', formatter: (value) => this.formatDate(value) },
            { key: 'region', label: '地区', formatter: (value) => value },
            { key: 'contractor', label: '契約者', formatter: (value) => value },
            { key: 'contractorAge', label: '年齢', formatter: (value) => value },
            { key: 'isOvertime', label: '時間外', formatter: (value) => value ? '○' : '×' },
            { key: 'isElderly', label: '高齢者', formatter: (value) => value ? '○' : '×' },
            { key: 'isExcessive', label: '過量販売', formatter: (value) => value ? '○' : '×' },
            { key: 'isSingle', label: '単独契約', formatter: (value) => value ? '○' : '×' },
            { key: 'productName', label: '商品名', formatter: (value) => value },
            { key: 'amount', label: '金額', formatter: (value) => this.formatAmount(value) }
        ];
        // フィルタリングとソート
        const filteredData = this.filterAndSortData(data);
        // ページネーション設定
        const itemsPerPage = 50;
        const totalPages = Math.ceil(filteredData.length / itemsPerPage);
        let html = `
            <div class="d-flex justify-content-between align-items-center mb-3">
                <h5>データ一覧 (${filteredData.length}件)</h5>
                <div>
                    <button class="btn btn-outline-primary btn-sm" id="exportDataBtn">
                        <i class="fas fa-download me-2"></i>データエクスポート
                    </button>
                </div>
            </div>
            
            ${this.createFilterStatusHTML(filterState)}
            
            <!-- フィルタ -->
            <div class="row mb-3">
                <div class="col-md-3">
                    <select class="form-select form-select-sm" id="regionFilter">
                        <option value="">全地区</option>
                        <option value="九州地区" ${filterState?.region === '九州地区' ? 'selected' : ''}>九州地区</option>
                        <option value="中四国地区" ${filterState?.region === '中四国地区' ? 'selected' : ''}>中四国地区</option>
                        <option value="関西地区" ${filterState?.region === '関西地区' ? 'selected' : ''}>関西地区</option>
                        <option value="関東地区" ${filterState?.region === '関東地区' ? 'selected' : ''}>関東地区</option>
                    </select>
                </div>
                <div class="col-md-3">
                    <select class="form-select form-select-sm" id="ageFilter">
                        <option value="">全年齢</option>
                        <option value="elderly" ${filterState?.age === 'elderly' ? 'selected' : ''}>高齢者（70歳以上）</option>
                        <option value="normal" ${filterState?.age === 'normal' ? 'selected' : ''}>通常年齢（69歳以下）</option>
                    </select>
                </div>
                <div class="col-md-3">
                    <select class="form-select form-select-sm" id="overtimeFilter">
                        <option value="">全件</option>
                        <option value="true" ${filterState?.overtime === 'true' ? 'selected' : ''}>時間外対応のみ</option>
                        <option value="false" ${filterState?.overtime === 'false' ? 'selected' : ''}>通常時間のみ</option>
                    </select>
                </div>
                <div class="col-md-3">
                    <button class="btn btn-outline-secondary btn-sm w-100" id="applyFilterBtn">
                        <i class="fas fa-filter me-2"></i>フィルタ適用
                    </button>
                </div>
            </div>
            
            <!-- データテーブル -->
            <div class="table-responsive">
                <table class="table table-striped table-hover data-table">
                    <thead>
                        <tr>
                            ${displayColumns.map(col => `<th>${col.label}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
        `;
        // 最初のページのデータを表示
        const firstPageData = filteredData.slice(0, itemsPerPage);
        firstPageData.forEach(row => {
            html += '<tr>';
            displayColumns.forEach(col => {
                const value = row[col.key];
                const formattedValue = col.formatter ? col.formatter(value) : value;
                html += `<td>${formattedValue || '-'}</td>`;
            });
            html += '</tr>';
        });
        html += `
                    </tbody>
                </table>
            </div>
            
            <!-- ページネーション -->
            ${this.createPaginationHTML(totalPages, 1)}
        `;
        return html;
    }
    filterAndSortData(data) {
        // 日付でソート（新しい順）
        return data.sort((a, b) => {
            if (!a.date || !b.date)
                return 0;
            return b.date.getTime() - a.date.getTime();
        });
    }
    createPaginationHTML(totalPages, currentPage) {
        if (totalPages <= 1)
            return '';
        let html = '<nav aria-label="データページネーション"><ul class="pagination justify-content-center">';
        // 前のページ
        if (currentPage > 1) {
            html += `<li class="page-item"><a class="page-link" href="#" data-page="${currentPage - 1}">前へ</a></li>`;
        }
        // ページ番号
        const startPage = Math.max(1, currentPage - 2);
        const endPage = Math.min(totalPages, currentPage + 2);
        for (let i = startPage; i <= endPage; i++) {
            const activeClass = i === currentPage ? 'active' : '';
            html += `<li class="page-item ${activeClass}"><a class="page-link" href="#" data-page="${i}">${i}</a></li>`;
        }
        // 次のページ
        if (currentPage < totalPages) {
            html += `<li class="page-item"><a class="page-link" href="#" data-page="${currentPage + 1}">次へ</a></li>`;
        }
        html += '</ul></nav>';
        return html;
    }
    formatDate(date) {
        if (!date)
            return '-';
        return date.toLocaleDateString('ja-JP');
    }
    formatAmount(amount) {
        if (!amount)
            return '-';
        const num = parseFloat(amount);
        if (isNaN(num))
            return '-';
        return num.toLocaleString('ja-JP') + '円';
    }
    createFilterStatusHTML(filterState) {
        if (!filterState || (!filterState.region && !filterState.age && !filterState.overtime)) {
            return '';
        }
        const activeFilters = [];
        if (filterState.region)
            activeFilters.push(`地区: ${filterState.region}`);
        if (filterState.age) {
            const ageLabel = filterState.age === 'elderly' ? '高齢者（70歳以上）' : '通常年齢（69歳以下）';
            activeFilters.push(`年齢: ${ageLabel}`);
        }
        if (filterState.overtime) {
            const overtimeLabel = filterState.overtime === 'true' ? '時間外対応のみ' : '通常時間のみ';
            activeFilters.push(`時間: ${overtimeLabel}`);
        }
        if (activeFilters.length === 0)
            return '';
        return `
            <div class="alert alert-info mb-3">
                <i class="fas fa-filter me-2"></i>
                <strong>適用中のフィルター:</strong> ${activeFilters.join(' | ')}
                <button class="btn btn-outline-secondary btn-sm ms-3" id="clearFiltersBtn">
                    <i class="fas fa-times me-1"></i>フィルター解除
                </button>
            </div>
        `;
    }
    // CSVエクスポート機能
    async exportDataToCSV(data) {
        try {
            // ヘッダー行
            const headers = [
                '日付', '地区', '契約者', '年齢', '時間外対応', '高齢者', '過量販売', '単独契約',
                '商品名', '金額', '担当者名', '契約者TEL', '確認者', '確認者TEL'
            ];
            // データ行
            const csvData = [headers];
            data.forEach(row => {
                csvData.push([
                    this.formatDate(row.date) || '',
                    row.region || '',
                    row.contractor || '',
                    row.contractorAge || '',
                    row.isOvertime ? '○' : '×',
                    row.isElderly ? '○' : '×',
                    row.isExcessive ? '○' : '×',
                    row.isSingle ? '○' : '×',
                    row.productName || '',
                    this.formatAmount(row.amount) || '',
                    row.staffName || '',
                    row.contractorTel || '',
                    row.confirmer || '',
                    row.confirmerTel || ''
                ]);
            });
            // CSV文字列の作成（強化されたUTF-8エンコーディング）
            const csvContent = csvData.map(row => row.map(cell => {
                // セル内容をエスケープして文字化けを防止
                const escapedCell = String(cell).replace(/"/g, '""');
                return `"${escapedCell}"`;
            }).join(',')).join('\r\n'); // Windows互換の改行コード
            // BOM付きUTF-8でBlobを作成（Excel対応）
            const bom = '\uFEFF';
            const blob = new Blob([bom + csvContent], { type: 'text/csv;charset=utf-8;' });
            const fileName = `データエクスポート_${new Date().toISOString().split('T')[0]}.csv`;
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
            console.error('CSVエクスポートエラー:', error);
            throw error;
        }
    }
    // データの統計情報を取得
    getDataStatistics(data) {
        if (!data || data.length === 0) {
            return {
                totalRecords: 0,
                dateRange: { start: null, end: null },
                regionDistribution: {},
                ageDistribution: {},
                overtimeRate: 0
            };
        }
        const dates = data.filter(row => row.date).map(row => row.date);
        const startDate = dates.length > 0 ? new Date(Math.min(...dates.map(d => d.getTime()))) : null;
        const endDate = dates.length > 0 ? new Date(Math.max(...dates.map(d => d.getTime()))) : null;
        const regionDistribution = {};
        const ageDistribution = { elderly: 0, normal: 0 };
        let overtimeCount = 0;
        data.forEach(row => {
            // 地区分布
            const region = row.region || '不明';
            regionDistribution[region] = (regionDistribution[region] || 0) + 1;
            // 年齢分布
            if (row.isElderly) {
                ageDistribution.elderly++;
            }
            else {
                ageDistribution.normal++;
            }
            // 時間外対応
            if (row.isOvertime) {
                overtimeCount++;
            }
        });
        return {
            totalRecords: data.length,
            dateRange: { start: startDate, end: endDate },
            regionDistribution,
            ageDistribution,
            overtimeRate: data.length > 0 ? (overtimeCount / data.length * 100).toFixed(1) : 0
        };
    }
}
//# sourceMappingURL=data-manager.js.map