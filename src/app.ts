import { ExcelProcessor } from './excel-processor.js';
import { ReportGenerator } from './report-generator.js';
import { DataManager } from './data-manager.js';

class App {
    private excelProcessor: ExcelProcessor;
    private reportGenerator: ReportGenerator;
    private dataManager: DataManager;
    private currentData: any[] = [];

    constructor() {
        this.excelProcessor = new ExcelProcessor();
        this.reportGenerator = new ReportGenerator();
        this.dataManager = new DataManager();
        
        this.initializeApp();
    }

    private initializeApp(): void {
        this.setupEventListeners();
        this.setupDateDefaults();
        this.setupDragAndDrop();
    }

    private setupEventListeners(): void {
        // ファイル選択ボタン
        const fileInputBtn = document.getElementById('fileInputBtn');
        const fileInput = document.getElementById('fileInput') as HTMLInputElement;
        
        if (fileInputBtn && fileInput) {
            fileInputBtn.addEventListener('click', () => fileInput.click());
            fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        }

        // データ処理ボタン
        const processDataBtn = document.getElementById('processDataBtn');
        if (processDataBtn) {
            processDataBtn.addEventListener('click', () => this.processData());
        }

        // 日付変更時のイベントリスナー
        const reportDate = document.getElementById('reportDate') as HTMLInputElement;
        if (reportDate) {
            reportDate.addEventListener('change', () => this.handleDateChange());
        }

        // タブ切り替え
        const tabs = document.querySelectorAll('[data-bs-toggle="tab"]');
        tabs.forEach(tab => {
            tab.addEventListener('shown.bs.tab', (e) => this.handleTabChange(e));
        });
    }

    private setupDateDefaults(): void {
        const today = new Date();
        const reportDate = document.getElementById('reportDate') as HTMLInputElement;
        const reportMonth = document.getElementById('reportMonth') as HTMLInputElement;
        
        if (reportDate) {
            reportDate.value = today.toISOString().split('T')[0];
            // 日付変更時のイベントリスナー
            reportDate.addEventListener('change', () => this.handleDateChange());
        }
        
        if (reportMonth) {
            const year = today.getFullYear();
            const month = String(today.getMonth() + 1).padStart(2, '0');
            reportMonth.value = `${year}-${month}`;
            // 月変更時のイベントリスナー
            reportMonth.addEventListener('change', () => this.handleMonthChange());
        }
    }

    private setupDragAndDrop(): void {
        const dropZone = document.getElementById('dropZone');
        
        if (dropZone) {
            dropZone.addEventListener('dragover', (e) => {
                e.preventDefault();
                dropZone.classList.add('dragover');
            });

            dropZone.addEventListener('dragleave', () => {
                dropZone.classList.remove('dragover');
            });

            dropZone.addEventListener('drop', (e) => {
                e.preventDefault();
                dropZone.classList.remove('dragover');
                this.handleFileDrop(e);
            });
        }
    }

    private handleFileSelect(event: Event): void {
        const target = event.target as HTMLInputElement;
        if (target.files && target.files.length > 0) {
            this.processExcelFile(target.files[0]);
        }
    }

    private handleFileDrop(event: DragEvent): void {
        if (event.dataTransfer && event.dataTransfer.files.length > 0) {
            this.processExcelFile(event.dataTransfer.files[0]);
        }
    }

    private async processExcelFile(file: File): Promise<void> {
        try {
            this.showLoading(true);
            
            // ファイル情報を表示
            this.showFileInfo(file.name);
            
            // Excelファイルを読み込み
            this.currentData = await this.excelProcessor.readExcelFile(file);
            
            // データを保存
            this.dataManager.saveData(this.currentData);
            
            this.showLoading(false);
            
        } catch (error) {
            console.error('ファイル処理エラー:', error);
            this.showError('ファイルの処理中にエラーが発生しました。');
            this.showLoading(false);
        }
    }

    private async processData(): Promise<void> {
        if (this.currentData.length === 0) {
            this.showError('データが読み込まれていません。');
            return;
        }

        try {
            this.showLoading(true);
            
            console.log('データ処理開始 - 既存データを使用');
            console.log('現在の日付設定:', {
                reportDate: (document.getElementById('reportDate') as HTMLInputElement)?.value,
                currentDataLength: this.currentData.length
            });
            
            // 日報・月報の生成（既存データを新しい日付で再フィルタリング）
            await this.generateReports();
            
            this.showLoading(false);
            
        } catch (error) {
            console.error('データ処理エラー:', error);
            this.showError('データの処理中にエラーが発生しました。');
            this.showLoading(false);
        }
    }

    private async generateReports(): Promise<void> {
        const reportDate = (document.getElementById('reportDate') as HTMLInputElement)?.value;
        const reportMonth = (document.getElementById('reportMonth') as HTMLInputElement)?.value;
        
        if (reportDate) {
            const dailyReport = this.reportGenerator.generateDailyReport(this.currentData, reportDate);
            this.displayDailyReport(dailyReport);
        }
        
        if (reportMonth) {
            const monthlyReport = this.reportGenerator.generateMonthlyReport(this.currentData, reportMonth);
            this.displayMonthlyReport(monthlyReport);
        }
    }

    private async generateMonthlyReportOnly(): Promise<void> {
        const reportMonth = (document.getElementById('reportMonth') as HTMLInputElement)?.value;
        
        if (reportMonth && this.currentData.length > 0) {
            console.log('月報のみ再生成:', reportMonth);
            const monthlyReport = this.reportGenerator.generateMonthlyReport(this.currentData, reportMonth);
            this.displayMonthlyReport(monthlyReport);
        }
    }

    private displayDailyReport(report: any): void {
        const content = document.getElementById('dailyReportContent');
        if (!content) return;

        content.innerHTML = this.reportGenerator.createDailyReportHTML(report);
        this.setupExportButtons('daily', report);
    }

    private displayMonthlyReport(report: any): void {
        const content = document.getElementById('monthlyReportContent');
        if (!content) return;

        content.innerHTML = this.reportGenerator.createMonthlyReportHTML(report);
        this.setupExportButtons('monthly', report);
    }

    private setupExportButtons(type: string, report: any): void {
        const exportButtons = document.querySelectorAll(`#${type}ReportContent .btn-export`);
        
        exportButtons.forEach(button => {
            button.addEventListener('click', async (e) => {
                const target = e.target as HTMLElement;
                const format = target.getAttribute('data-format');
                
                if (format === 'pdf') {
                    await this.reportGenerator.exportToPDF(report, type);
                } else if (format === 'csv') {
                    await this.reportGenerator.exportToCSV(report, type);
                }
            });
        });
    }

    private handleDateChange(): void {
        console.log('日付変更を検出');
        // 日付が変更された場合、既存データでレポートを再生成
        if (this.currentData.length > 0) {
            console.log('既存データでレポートを再生成');
            this.generateReports();
        }
    }

    private handleMonthChange(): void {
        console.log('月変更を検出');
        // 月が変更された場合、既存データで月報を再生成
        if (this.currentData.length > 0) {
            console.log('既存データで月報を再生成');
            this.generateMonthlyReportOnly();
        }
    }

    private handleTabChange(event: Event): void {
        const target = event.target as HTMLElement;
        const tabId = target.getAttribute('id');
        
        if (tabId === 'data-tab' && this.currentData.length > 0) {
            this.displayDataTable();
        }
    }

    private displayDataTable(): void {
        const content = document.getElementById('dataContent');
        if (!content) return;

        content.innerHTML = this.dataManager.createDataTableHTML(this.currentData);
        
        // フィルタ機能のイベントリスナーを設定
        this.setupDataTableFilters();
    }
    
    private setupDataTableFilters(): void {
        const applyFilterBtn = document.getElementById('applyFilterBtn');
        if (applyFilterBtn) {
            applyFilterBtn.addEventListener('click', () => this.applyDataFilters());
        }
        
        // エクスポートボタンのイベントリスナー
        const exportDataBtn = document.getElementById('exportDataBtn');
        if (exportDataBtn) {
            exportDataBtn.addEventListener('click', () => this.exportFilteredData());
        }
        
        // フィルター解除ボタンのイベントリスナー
        const clearFiltersBtn = document.getElementById('clearFiltersBtn');
        if (clearFiltersBtn) {
            clearFiltersBtn.addEventListener('click', () => this.clearDataFilters());
        }
    }
    
    private applyDataFilters(): void {
        const regionFilter = (document.getElementById('regionFilter') as HTMLSelectElement)?.value;
        const ageFilter = (document.getElementById('ageFilter') as HTMLSelectElement)?.value;
        const overtimeFilter = (document.getElementById('overtimeFilter') as HTMLSelectElement)?.value;
        
        let filteredData = [...this.currentData];
        
        // 地区フィルタ
        if (regionFilter) {
            filteredData = filteredData.filter(row => row.region === regionFilter);
        }
        
        // 年齢フィルタ
        if (ageFilter) {
            if (ageFilter === 'elderly') {
                filteredData = filteredData.filter(row => row.isElderly);
            } else if (ageFilter === 'normal') {
                filteredData = filteredData.filter(row => !row.isElderly);
            }
        }
        
        // 時間外フィルタ
        if (overtimeFilter) {
            const isOvertime = overtimeFilter === 'true';
            filteredData = filteredData.filter(row => row.isOvertime === isOvertime);
        }
        
        // フィルタ結果を表示（フィルター状態を保持）
        const content = document.getElementById('dataContent');
        if (content) {
            content.innerHTML = this.dataManager.createDataTableHTML(filteredData, {
                region: regionFilter,
                age: ageFilter,
                overtime: overtimeFilter
            });
            this.setupDataTableFilters(); // イベントリスナーを再設定
        }
    }
    
    private clearDataFilters(): void {
        // フィルターをリセット
        const regionFilter = document.getElementById('regionFilter') as HTMLSelectElement;
        const ageFilter = document.getElementById('ageFilter') as HTMLSelectElement;
        const overtimeFilter = document.getElementById('overtimeFilter') as HTMLSelectElement;
        
        if (regionFilter) regionFilter.value = '';
        if (ageFilter) ageFilter.value = '';
        if (overtimeFilter) overtimeFilter.value = '';
        
        // 全データを表示
        const content = document.getElementById('dataContent');
        if (content) {
            content.innerHTML = this.dataManager.createDataTableHTML(this.currentData);
            this.setupDataTableFilters();
        }
    }
    
    private async exportFilteredData(): Promise<void> {
        try {
            // 現在のフィルタ条件を取得
            const regionFilter = (document.getElementById('regionFilter') as HTMLSelectElement)?.value;
            const ageFilter = (document.getElementById('ageFilter') as HTMLSelectElement)?.value;
            const overtimeFilter = (document.getElementById('overtimeFilter') as HTMLSelectElement)?.value;
            
            let filteredData = [...this.currentData];
            
            // フィルタを適用
            if (regionFilter) {
                filteredData = filteredData.filter(row => row.region === regionFilter);
            }
            if (ageFilter) {
                if (ageFilter === 'elderly') {
                    filteredData = filteredData.filter(row => row.isElderly);
                } else if (ageFilter === 'normal') {
                    filteredData = filteredData.filter(row => !row.isElderly);
                }
            }
            if (overtimeFilter) {
                const isOvertime = overtimeFilter === 'true';
                filteredData = filteredData.filter(row => row.isOvertime === isOvertime);
            }
            
            // CSVエクスポート
            await this.dataManager.exportDataToCSV(filteredData);
            
        } catch (error) {
            console.error('データエクスポートエラー:', error);
            alert('データのエクスポートに失敗しました。');
        }
    }

    private showFileInfo(fileName: string): void {
        const fileInfo = document.getElementById('fileInfo');
        const fileNameElement = document.getElementById('fileName');
        
        if (fileInfo && fileNameElement) {
            fileNameElement.textContent = fileName;
            fileInfo.style.display = 'block';
        }
    }

    private showLoading(show: boolean): void {
        const loadingElements = document.querySelectorAll('.loading');
        loadingElements.forEach(element => {
            if (show) {
                element.classList.add('show');
            } else {
                element.classList.remove('show');
            }
        });
    }

    private showError(message: string): void {
        // エラーメッセージを表示する処理
        alert(message);
    }
}

// アプリケーションの初期化
document.addEventListener('DOMContentLoaded', () => {
    new App();
});