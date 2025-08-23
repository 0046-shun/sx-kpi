import { ExcelProcessor } from './excel-processor.js';
import { ReportGenerator } from './report-generator.js';
import { DataManager } from './data-manager.js';
import { CalendarManager } from './calendar-manager.js';
export class App {
    constructor() {
        this.currentData = [];
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
        if (dropArea) {
            dropArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                dropArea.classList.add('dragover');
            });
            dropArea.addEventListener('dragleave', (e) => {
                dropArea.classList.remove('dragover');
            });
            dropArea.addEventListener('drop', (e) => {
                e.preventDefault();
                dropArea.classList.remove('dragover');
                const files = e.dataTransfer?.files;
                if (files && files.length > 0) {
                    this.handleFileUpload(files[0]);
                }
            });
            // ドラッグエンターイベントも追加
            dropArea.addEventListener('dragenter', (e) => {
                e.preventDefault();
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
        const dailyContainer = document.getElementById('dailyReportContent');
        const monthlyContainer = document.getElementById('monthlyReportContent');
        if (dailyContainer && dailyContainer.innerHTML.trim() !== '' &&
            dailyContainer.querySelector('.report-title')?.textContent?.includes('日報')) {
            this.generateDailyReport();
        }
        else if (monthlyContainer && monthlyContainer.innerHTML.trim() !== '' &&
            monthlyContainer.querySelector('.report-title')?.textContent?.includes('月報')) {
            this.generateMonthlyReport();
        }
    }
    // ファイルアップロード処理
    async handleFileUpload(file) {
        try {
            // スタートメッセージを表示
            this.showMessage(`ファイル「${file.name}」の読み込みを開始しました...`, 'info');
            // Excelファイルを読み込み
            const data = await this.excelProcessor.readExcelFile(file);
            // データを保存
            this.dataManager.setData(data);
            // 現在のデータを保存
            this.currentData = data;
            // 担当別データと確認データを即座に更新
            this.updateStaffData();
            this.updateDataConfirmation();
            // 完了メッセージを表示
            this.showMessage(`ファイル「${file.name}」の読み込みが完了しました。総データ件数: ${data.length}件`, 'success');
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
        const reportMonthInput = document.getElementById('reportMonth');
        if (!reportMonthInput || !reportMonthInput.value) {
            this.showMessage('月報作成月を選択してください。', 'error');
            return;
        }
        try {
            // 選択された月から年月を取得
            const monthString = reportMonthInput.value; // 既に "YYYY-MM" 形式
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
        if (type === 'daily') {
            const dailyContainer = document.getElementById('dailyReportContent');
            if (dailyContainer) {
                dailyContainer.innerHTML = this.reportGenerator.createDailyReportHTML(report);
            }
        }
        else {
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
    updateStaffData() {
        if (this.currentData && this.currentData.length > 0) {
            const staffData = this.reportGenerator.generateStaffData(this.currentData);
            const staffContainer = document.getElementById('staffDataContent');
            if (staffContainer) {
                staffContainer.innerHTML = this.reportGenerator.createStaffDataHTML(staffData);
            }
            else {
                console.error('staffDataContent要素が見つかりません');
            }
        }
        else {
        }
    }
    // 確認データを更新
    updateDataConfirmation() {
        if (this.currentData && this.currentData.length > 0) {
            const dataContainer = document.getElementById('dataConfirmationContent');
            if (dataContainer) {
                dataContainer.innerHTML = this.reportGenerator.createDataConfirmationHTML(this.currentData);
                // フィルター機能を設定
                this.setupDataFilters();
            }
            else {
                console.error('dataConfirmationContent要素が見つかりません');
            }
        }
        else {
        }
    }
    // データフィルター機能を設定
    setupDataFilters() {
        const staffFilter = document.getElementById('staffFilter');
        const regionFilter = document.getElementById('regionFilter');
        const departmentFilter = document.getElementById('departmentFilter');
        const table = document.getElementById('dataConfirmationTable');
        if (!table)
            return;
        const filterData = () => {
            const staffValue = staffFilter?.value.toLowerCase() || '';
            const regionValue = regionFilter?.value.toLowerCase() || '';
            const departmentValue = departmentFilter?.value.toLowerCase() || '';
            const rows = table.querySelectorAll('tbody tr');
            let visibleCount = 0;
            rows.forEach((row) => {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 5) {
                    const staffName = cells[2]?.textContent?.toLowerCase() || '';
                    const regionNo = cells[3]?.textContent?.toLowerCase() || '';
                    const departmentNo = cells[4]?.textContent?.toLowerCase() || '';
                    const matchesStaff = !staffValue || staffName.includes(staffValue);
                    const matchesRegion = !regionValue || regionNo.includes(regionValue);
                    const matchesDepartment = !departmentValue || departmentNo.includes(departmentValue);
                    if (matchesStaff && matchesRegion && matchesDepartment) {
                        row.style.display = '';
                        visibleCount++;
                    }
                    else {
                        row.style.display = 'none';
                    }
                }
            });
            // 表示件数を更新
            const infoAlert = table.parentElement?.nextElementSibling;
            if (infoAlert && infoAlert.classList.contains('alert-info')) {
                infoAlert.innerHTML = `<small>※ フィルター適用中（表示件数: ${visibleCount}件 / 総件数: ${this.currentData.length}件）</small>`;
            }
        };
        // フィルター入力時のイベントリスナーを設定
        if (staffFilter)
            staffFilter.addEventListener('input', filterData);
        if (regionFilter)
            regionFilter.addEventListener('input', filterData);
        if (departmentFilter)
            departmentFilter.addEventListener('input', filterData);
    }
    // エクスポートボタンの設定
    setupExportButtons(type, report) {
        const containerSelector = type === 'daily' ? '#dailyReportContent' : '#monthlyReportContent';
        const exportButtons = document.querySelectorAll(`${containerSelector} .btn-export`);
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
        const existingMessages = document.querySelectorAll('.alert-message');
        existingMessages.forEach(msg => msg.remove());
        // 新しいメッセージを作成
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type === 'error' ? 'danger' : type} alert-dismissible fade show alert-message`;
        alertDiv.style.position = 'fixed';
        alertDiv.style.top = '20px';
        alertDiv.style.right = '20px';
        alertDiv.style.zIndex = '9999';
        alertDiv.style.minWidth = '300px';
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
        `;
        // メッセージをbodyに追加
        document.body.appendChild(alertDiv);
        // 閉じるボタンのイベントリスナー
        const closeButton = alertDiv.querySelector('.btn-close');
        if (closeButton) {
            closeButton.addEventListener('click', () => {
                alertDiv.remove();
            });
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
        const reportMonthInput = document.getElementById('reportMonth');
        if (dateInput) {
            const today = new Date();
            dateInput.value = today.toISOString().split('T')[0];
        }
        if (reportMonthInput) {
            const today = new Date();
            const year = today.getFullYear();
            const month = String(today.getMonth() + 1).padStart(2, '0');
            reportMonthInput.value = `${year}-${month}`;
        }
    }
}
// アプリケーションの初期化
document.addEventListener('DOMContentLoaded', () => {
    new App();
});
//# sourceMappingURL=app.js.map