export class ReportGenerator {
    generateDailyReport(data, date) {
        const targetDate = new Date(date);
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
        return reportData;
    }
    calculateReportData(data, type) {
        // 基本統計
        const totalOrders = data.length;
        const overtimeOrders = data.filter(row => row.isOvertime).length;
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
            '九州地区': { orders: 0, overtime: 0, excessive: 0, single: 0 },
            '中四国地区': { orders: 0, overtime: 0, excessive: 0, single: 0 },
            '関西地区': { orders: 0, overtime: 0, excessive: 0, single: 0 },
            '関東地区': { orders: 0, overtime: 0, excessive: 0, single: 0 },
            'その他': { orders: 0, overtime: 0, excessive: 0, single: 0 }
        };
        data.forEach(row => {
            const region = row.region;
            if (regions[region]) {
                regions[region].orders++;
                if (row.isOvertime) {
                    regions[region].overtime++;
                }
                if (row.isExcessive) {
                    regions[region].excessive++;
                }
                if (row.isSingle) {
                    regions[region].single++;
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
                <div class="row mb-4">
                    <div class="col-md-6">
                        <div class="stat-card">
                            <div class="stat-number">${report.totalOrders}</div>
                            <div class="stat-label">総受注件数</div>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="stat-card">
                            <div class="stat-number">${report.overtimeOrders}</div>
                            <div class="stat-label">総時間外対応件数<br>(18:30以降)</div>
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
                <div class="row mb-4">
                    <div class="col-md-6">
                        <div class="stat-card">
                            <div class="stat-number">${report.totalOrders}</div>
                            <div class="stat-label">総受注件数</div>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="stat-card">
                            <div class="stat-number">${report.overtimeOrders}</div>
                            <div class="stat-label">総時間外対応件数<br>(18:30以降)</div>
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
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();
            // 日本語フォントの設定（デフォルトフォントを使用）
            doc.setFont('helvetica');
            // PDFのタイトル
            const title = type === 'daily' ? '日報' : '月報';
            doc.setFontSize(20);
            doc.text(title, 20, 20);
            // 基本統計
            doc.setFontSize(14);
            doc.text(`総受注件数: ${report.totalOrders}`, 20, 40);
            doc.text(`総時間外対応件数: ${report.overtimeOrders}`, 20, 50);
            // 地区別統計
            let yPos = 70;
            doc.setFontSize(12);
            doc.text('地区別受注件数:', 20, yPos);
            yPos += 10;
            Object.entries(report.regionStats).forEach(([region, stats]) => {
                if (stats.orders > 0) {
                    doc.text(`${region}: ${stats.orders}件 (時間外: ${stats.overtime}件, 過量: ${stats.excessive}件, 単独: ${stats.single}件)`, 30, yPos);
                    yPos += 7;
                }
            });
            // 年齢別統計
            yPos += 10;
            doc.text('年齢別集計:', 20, yPos);
            yPos += 10;
            doc.text(`高齢者(70歳以上): ${report.ageStats.elderly.total}件`, 30, yPos);
            yPos += 7;
            doc.text(`- 過量販売: ${report.ageStats.elderly.excessive}件`, 40, yPos);
            yPos += 7;
            doc.text(`- 単独契約: ${report.ageStats.elderly.single}件`, 40, yPos);
            yPos += 7;
            doc.text(`通常年齢(69歳以下): ${report.ageStats.normal.total}件`, 30, yPos);
            yPos += 7;
            doc.text(`- 過量販売: ${report.ageStats.normal.excessive}件`, 40, yPos);
            // PDFをダウンロード
            const fileName = `${type === 'daily' ? '日報' : '月報'}_${new Date().toISOString().split('T')[0]}.pdf`;
            doc.save(fileName);
        }
        catch (error) {
            console.error('PDF出力エラー:', error);
            alert('PDFの出力に失敗しました。');
        }
    }
    async exportToCSV(report, type) {
        try {
            // CSVデータの作成
            const csvData = [
                ['項目', '値'],
                ['総受注件数', report.totalOrders],
                ['総時間外対応件数', report.overtimeOrders],
                [''],
                ['地区別受注件数', ''],
                ...Object.entries(report.regionStats)
                    .filter(([_, stats]) => stats.orders > 0)
                    .map(([region, stats]) => [region, `${stats.orders}件 (時間外: ${stats.overtime}件, 過量: ${stats.excessive}件, 単独: ${stats.single}件)`]),
                [''],
                ['年齢別集計', ''],
                ['高齢者(70歳以上)', report.ageStats.elderly.total + '件'],
                ['- 過量販売', report.ageStats.elderly.excessive + '件'],
                ['- 単独契約', report.ageStats.elderly.single + '件'],
                ['通常年齢(69歳以下)', report.ageStats.normal.total + '件'],
                ['- 過量販売', report.ageStats.normal.excessive + '件']
            ];
            // CSV文字列の作成（BOM付きUTF-8）
            const csvContent = '\uFEFF' + csvData.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
            // CSVファイルをダウンロード
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
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
}
//# sourceMappingURL=report-generator.js.map