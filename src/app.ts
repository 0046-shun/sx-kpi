import { ExcelProcessor } from './excel-processor.js';
import { ReportGenerator } from './report-generator.js';
import { DataManager } from './data-manager.js';
import { CalendarManager } from './calendar-manager.js';
import { HolidaySettings } from './types.js';

export class App {
    private excelProcessor: ExcelProcessor;
    private reportGenerator: ReportGenerator;
    private dataManager: DataManager;
    private calendarManager: CalendarManager;
    private currentData: any[] = [];

    constructor() {
        this.excelProcessor = new ExcelProcessor();
        this.reportGenerator = new ReportGenerator();
        this.dataManager = DataManager.getInstance();
        this.calendarManager = new CalendarManager();
        
        this.initializeApp();
    }

    private initializeApp(): void {
        this.setupEventListeners();
        this.initializeCalendarManager();
        this.loadSavedHolidaySettings();
        this.setDefaultDate();
    }

    private setupEventListeners(): void {
        // ファイルドロップエリアの設定
        const dropArea = document.getElementById('dropArea');
        console.log('ドロップエリア要素:', dropArea);
        
        if (dropArea) {
            console.log('ドロップエリアのイベントリスナーを設定中...');
            
            dropArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                console.log('dragover イベント発生');
                dropArea.classList.add('dragover');
            });
            
            dropArea.addEventListener('dragleave', (e) => {
                console.log('dragleave イベント発生');
                dropArea.classList.remove('dragover');
            });
            
            dropArea.addEventListener('drop', (e) => {
                e.preventDefault();
                console.log('drop イベント発生');
                dropArea.classList.remove('dragover');
                const files = e.dataTransfer?.files;
                console.log('ドロップされたファイル:', files);
                if (files && files.length > 0) {
                    this.handleFileUpload(files[0]);
                }
            });
            
            // ドラッグエンターイベントも追加
            dropArea.addEventListener('dragenter', (e) => {
                e.preventDefault();
                console.log('dragenter イベント発生');
            });
        } else {
            console.error('ドロップエリア要素が見つかりません');
        }
        
        // ファイル選択ボタン
        const fileInputBtn = document.getElementById('fileInputBtn');
        const fileInput = document.getElementById('fileInput') as HTMLInputElement;
        
        if (fileInputBtn && fileInput) {
            fileInputBtn.addEventListener('click', () => fileInput.click());
            fileInput.addEventListener('change', (e) => {
                const target = e.target as HTMLInputElement;
                if (target.files && target.files.length > 0) {
                    this.handleFileUpload(target.files[0]);
                }
            });
        }
        
        // 日報生成ボタン
        const generateDailyReportBtn = document.getElementById('generateDailyReport');
        if (generateDailyReportBtn) {
            generateDailyReportBtn.addEventListener('click', () => {
                this.generateDailyReport();
            });
        }
        
        // 月報生成ボタン
        const generateMonthlyReportBtn = document.getElementById('generateMonthlyReport');
        if (generateMonthlyReportBtn) {
            generateMonthlyReportBtn.addEventListener('click', () => {
                this.generateMonthlyReport();
            });
        }
        
        // 公休日・禁止日設定ボタン
        const openHolidaySettingsBtn = document.getElementById('openHolidaySettings');
        if (openHolidaySettingsBtn) {
            openHolidaySettingsBtn.addEventListener('click', () => {
                this.openHolidaySettings();
            });
        }
        
        // 公休日設定変更イベント
        document.addEventListener('holidaySettingsChanged', (e: any) => {
            this.handleHolidaySettingsChanged(e.detail);
        });
    }
    
    private initializeCalendarManager(): void {
        // カレンダーマネージャーの初期化
        this.calendarManager.initializeCalendars();
    }
    
    private loadSavedHolidaySettings(): void {
        // 保存された設定を読み込み
        const savedSettings = this.dataManager.getHolidaySettings();
        this.calendarManager.setHolidaySettings(savedSettings);
        this.reportGenerator.updateHolidaySettings(savedSettings);
    }
    
    private openHolidaySettings(): void {
        // モーダルを開く
        const modal = document.getElementById('holidaySettingsModal');
        if (modal) {
            const bootstrapModal = new (window as any).bootstrap.Modal(modal);
            bootstrapModal.show();
        }
    }
    
    private handleHolidaySettingsChanged(settings: HolidaySettings): void {
        // 設定を保存
        this.dataManager.setHolidaySettings(settings);
        
        // レポートジェネレーターに設定を反映
        this.reportGenerator.updateHolidaySettings(settings);
        
        // 現在のレポートがあれば再生成
        this.refreshCurrentReport();
    }
    
    private refreshCurrentReport(): void {
        // 現在表示されているレポートを再生成
        const dailyContainer = document.getElementById('dailyReportContent');
        const monthlyContainer = document.getElementById('monthlyReportContent');
        
        if (dailyContainer && dailyContainer.innerHTML.trim() !== '' && 
            dailyContainer.querySelector('.report-title')?.textContent?.includes('日報')) {
            this.generateDailyReport();
        } else if (monthlyContainer && monthlyContainer.innerHTML.trim() !== '' && 
                   monthlyContainer.querySelector('.report-title')?.textContent?.includes('月報')) {
            this.generateMonthlyReport();
        }
    }
    
    // ファイルアップロード処理
    private async handleFileUpload(file: File): Promise<void> {
        try {
            console.log('ファイルアップロード開始:', file.name);
            
            // Excelファイルを読み込み
            const data = await this.excelProcessor.readExcelFile(file);
            
            // データを保存
            this.dataManager.setData(data);
            
            // 現在のデータを保存
            this.currentData = data;
            
            // 担当別データを即座に更新
            this.updateStaffData();
            
            console.log('ファイル処理完了:', data.length, '件');
            
            // 成功メッセージを表示
            this.showMessage('ファイルの読み込みが完了しました。', 'success');
            
        } catch (error) {
            console.error('ファイル処理エラー:', error);
            this.showMessage('ファイルの処理中にエラーが発生しました。', 'error');
        }
    }
    
    // 日報生成
    private generateDailyReport(): void {
        const data = this.dataManager.getData();
        if (data.length === 0) {
            this.showMessage('データが読み込まれていません。', 'error');
            return;
        }
        
        const dateInput = document.getElementById('dateInput') as HTMLInputElement;
        if (!dateInput || !dateInput.value) {
            this.showMessage('日付を選択してください。', 'error');
            return;
        }

        try {
            const report = this.reportGenerator.generateDailyReport(data, dateInput.value);
            this.displayReport(report, 'daily');
        } catch (error) {
            console.error('日報生成エラー:', error);
            this.showMessage('日報の生成中にエラーが発生しました。', 'error');
        }
    }
    
    // 月報生成
    private generateMonthlyReport(): void {
        const data = this.dataManager.getData();
        if (data.length === 0) {
            this.showMessage('データが読み込まれていません。', 'error');
            return;
        }
        
        const dateInput = document.getElementById('dateInput') as HTMLInputElement;
        if (!dateInput || !dateInput.value) {
            this.showMessage('日付を選択してください。', 'error');
            return;
        }
        
        try {
            // 選択された日付から年月を取得
            const selectedDate = new Date(dateInput.value);
            const year = selectedDate.getFullYear();
            const month = String(selectedDate.getMonth() + 1).padStart(2, '0');
            const monthString = `${year}-${month}`;
            
            const report = this.reportGenerator.generateMonthlyReport(data, monthString);
            this.displayReport(report, 'monthly');
        } catch (error) {
            console.error('月報生成エラー:', error);
            this.showMessage('月報の生成中にエラーが発生しました。', 'error');
        }
    }
    
    // レポート表示
    private displayReport(report: any, type: 'daily' | 'monthly'): void {
        if (type === 'daily') {
            const dailyContainer = document.getElementById('dailyReportContent');
            if (dailyContainer) {
                dailyContainer.innerHTML = this.reportGenerator.createDailyReportHTML(report);
            }
        } else {
            const monthlyContainer = document.getElementById('monthlyReportContent');
            if (monthlyContainer) {
                monthlyContainer.innerHTML = this.reportGenerator.createMonthlyReportHTML(report);
            }
        }
        
        // 担当別データも更新
        this.updateStaffData();
        
        // エクスポートボタンのイベントリスナーを設定
        this.setupExportButtons(type, report);
    }

    // 担当別データを更新
    private updateStaffData(): void {
        console.log('担当別データ更新開始, currentDataの件数:', this.currentData?.length || 0);
        if (this.currentData && this.currentData.length > 0) {
            const staffData = this.reportGenerator.generateStaffData(this.currentData);
            console.log('担当別データ生成完了, 担当者数:', staffData?.length || 0);
            const staffContainer = document.getElementById('staffDataContent');
            if (staffContainer) {
                staffContainer.innerHTML = this.reportGenerator.createStaffDataHTML(staffData);
                console.log('担当別データHTMLを更新しました');
            } else {
                console.error('staffDataContent要素が見つかりません');
            }
        } else {
            console.log('データがないため担当別データを更新できません');
        }
    }
    
    // エクスポートボタンの設定
    private setupExportButtons(type: string, report: any): void {
        const containerSelector = type === 'daily' ? '#dailyReportContent' : '#monthlyReportContent';
        const exportButtons = document.querySelectorAll(`${containerSelector} .btn-export`);
        
        exportButtons.forEach(button => {
            button.addEventListener('click', async (e) => {
                const target = e.target as HTMLElement;
                const format = target.getAttribute('data-format');
                
                try {
                if (format === 'pdf') {
                    await this.reportGenerator.exportToPDF(report, type);
                } else if (format === 'csv') {
                    await this.reportGenerator.exportToCSV(report, type);
                    }
                } catch (error) {
                    console.error('エクスポートエラー:', error);
                    this.showMessage('エクスポート中にエラーが発生しました。', 'error');
                }
            });
        });
    }

    // メッセージ表示
    private showMessage(message: string, type: 'success' | 'error' | 'info'): void {
        // 既存のメッセージを削除
        const existingMessage = document.querySelector('.alert');
        if (existingMessage) {
            existingMessage.remove();
        }
        
        // 新しいメッセージを作成
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type === 'error' ? 'danger' : type} alert-dismissible fade show`;
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        
        // メッセージを表示
        const mainContent = document.querySelector('.main-content');
        if (mainContent) {
            mainContent.insertBefore(alertDiv, mainContent.firstChild);
        }
        
        // 3秒後に自動で消す
        setTimeout(() => {
            if (alertDiv.parentNode) {
                alertDiv.remove();
            }
        }, 3000);
    }
    
    // デフォルト日付の設定
    private setDefaultDate(): void {
        const dateInput = document.getElementById('dateInput') as HTMLInputElement;
        if (dateInput) {
            const today = new Date();
            dateInput.value = today.toISOString().split('T')[0];
        }
    }
}

// アプリケーションの初期化
document.addEventListener('DOMContentLoaded', () => {
    new App();
});