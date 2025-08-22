import { ExcelProcessor } from './excel-processor.js';
import { ReportGenerator } from './report-generator.js';
import { DataManager } from './data-manager.js';
import { CalendarManager } from './calendar-manager.js';
export class App {
    constructor() {
        this.excelProcessor = new ExcelProcessor();
        this.reportGenerator = new ReportGenerator();
        this.dataManager = DataManager.getInstance();
        this.calendarManager = new CalendarManager();
        this.initializeApp();
    }
    initializeApp() {
        this.setupEventListeners();
        this.initializeCalendarManager();
        this.loadSavedHolidaySettings();
        this.setDefaultDate();
    }
    setupEventListeners() {
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
        }
        else {
            console.error('ドロップエリア要素が見つかりません');
        }
        // ファイル選択ボタン
        const fileInputBtn = document.getElementById('fileInputBtn');
        const fileInput = document.getElementById('fileInput');
        if (fileInputBtn && fileInput) {
            fileInputBtn.addEventListener('click', () => fileInput.click());
            fileInput.addEventListener('change', (e) => {
                const target = e.target;
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
        document.addEventListener('holidaySettingsChanged', (e) => {
            this.handleHolidaySettingsChanged(e.detail);
        });
    }
    initializeCalendarManager() {
        // カレンダーマネージャーの初期化
        this.calendarManager.initializeCalendars();
    }
    loadSavedHolidaySettings() {
        // 保存された設定を読み込み
        const savedSettings = this.dataManager.getHolidaySettings();
        this.calendarManager.setHolidaySettings(savedSettings);
        this.reportGenerator.updateHolidaySettings(savedSettings);
    }
    openHolidaySettings() {
        // モーダルを開く
        const modal = document.getElementById('holidaySettingsModal');
        if (modal) {
            const bootstrapModal = new window.bootstrap.Modal(modal);
            bootstrapModal.show();
        }
    }
    handleHolidaySettingsChanged(settings) {
        // 設定を保存
        this.dataManager.setHolidaySettings(settings);
        // レポートジェネレーターに設定を反映
        this.reportGenerator.updateHolidaySettings(settings);
        // 現在のレポートがあれば再生成
        this.refreshCurrentReport();
    }
    refreshCurrentReport() {
        // 現在表示されているレポートを再生成
        const reportContainer = document.getElementById('reportContainer');
        if (reportContainer && reportContainer.innerHTML.trim() !== '') {
            // 日報か月報かを判定して再生成
            if (reportContainer.querySelector('.report-title')?.textContent?.includes('日報')) {
                this.generateDailyReport();
            }
            else if (reportContainer.querySelector('.report-title')?.textContent?.includes('月報')) {
                this.generateMonthlyReport();
            }
        }
    }
    // ファイルアップロード処理
    async handleFileUpload(file) {
        try {
            console.log('ファイルアップロード開始:', file.name);
            // Excelファイルを読み込み
            const data = await this.excelProcessor.readExcelFile(file);
            // データを保存
            this.dataManager.setData(data);
            console.log('ファイル処理完了:', data.length, '件');
            // 成功メッセージを表示
            this.showMessage('ファイルの読み込みが完了しました。', 'success');
        }
        catch (error) {
            console.error('ファイル処理エラー:', error);
            this.showMessage('ファイルの処理中にエラーが発生しました。', 'error');
        }
    }
    // 日報生成
    generateDailyReport() {
        const data = this.dataManager.getData();
        if (data.length === 0) {
            this.showMessage('データが読み込まれていません。', 'error');
            return;
        }
        const dateInput = document.getElementById('dateInput');
        if (!dateInput || !dateInput.value) {
            this.showMessage('日付を選択してください。', 'error');
            return;
        }
        try {
            const report = this.reportGenerator.generateDailyReport(data, dateInput.value);
            this.displayReport(report, 'daily');
        }
        catch (error) {
            console.error('日報生成エラー:', error);
            this.showMessage('日報の生成中にエラーが発生しました。', 'error');
        }
    }
    // 月報生成
    generateMonthlyReport() {
        const data = this.dataManager.getData();
        if (data.length === 0) {
            this.showMessage('データが読み込まれていません。', 'error');
            return;
        }
        const dateInput = document.getElementById('dateInput');
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
        }
        catch (error) {
            console.error('月報生成エラー:', error);
            this.showMessage('月報の生成中にエラーが発生しました。', 'error');
        }
    }
    // レポート表示
    displayReport(report, type) {
        const reportContainer = document.getElementById('reportContainer');
        if (!reportContainer)
            return;
        if (type === 'daily') {
            reportContainer.innerHTML = this.reportGenerator.createDailyReportHTML(report);
        }
        else {
            reportContainer.innerHTML = this.reportGenerator.createMonthlyReportHTML(report);
        }
        // エクスポートボタンのイベントリスナーを設定
        this.setupExportButtons(type, report);
    }
    // エクスポートボタンの設定
    setupExportButtons(type, report) {
        const exportButtons = document.querySelectorAll(`#reportContainer .btn-export`);
        exportButtons.forEach(button => {
            button.addEventListener('click', async (e) => {
                const target = e.target;
                const format = target.getAttribute('data-format');
                try {
                    if (format === 'pdf') {
                        await this.reportGenerator.exportToPDF(report, type);
                    }
                    else if (format === 'csv') {
                        await this.reportGenerator.exportToCSV(report, type);
                    }
                }
                catch (error) {
                    console.error('エクスポートエラー:', error);
                    this.showMessage('エクスポート中にエラーが発生しました。', 'error');
                }
            });
        });
    }
    // メッセージ表示
    showMessage(message, type) {
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
    setDefaultDate() {
        const dateInput = document.getElementById('dateInput');
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
//# sourceMappingURL=app.js.map