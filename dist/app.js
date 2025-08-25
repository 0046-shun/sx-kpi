import { ExcelProcessor } from './excel-processor.js';
import { ReportGenerator } from './report-generator.js';
import { DataManager } from './data-manager.js';
import { CalendarManager } from './calendar-manager.js';
export class App {
    constructor() {
        this.currentData = [];
        this.calendarStates = {};
        this.excelProcessor = new ExcelProcessor();
        this.calendarManager = new CalendarManager();
        this.reportGenerator = new ReportGenerator(this.excelProcessor, this.calendarManager);
        this.dataManager = DataManager.getInstance();
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
        // 公休日・禁止日設定の初期化
        this.setupHolidaySettingsModal();
    }
    setupHolidaySettingsModal() {
        // モーダル表示時の処理
        const modal = document.getElementById('holidaySettingsModal');
        if (modal) {
            modal.addEventListener('show.bs.modal', () => {
                this.loadHolidaySettings();
                this.generateCalendar('publicHolidayCalendar', 'public');
                this.generateCalendar('prohibitedDayCalendar', 'prohibited');
            });
        }
        // 保存ボタンのイベントリスナー
        const saveButton = document.getElementById('saveHolidaySettings');
        if (saveButton) {
            saveButton.addEventListener('click', () => {
                this.saveHolidaySettings();
            });
        }
        // 公休日追加ボタンのイベントリスナー
        const addPublicHolidayButton = document.getElementById('addPublicHoliday');
        if (addPublicHolidayButton) {
            addPublicHolidayButton.addEventListener('click', () => {
                this.addPublicHoliday();
            });
        }
        // 禁止日追加ボタンのイベントリスナー
        const addProhibitedDayButton = document.getElementById('addProhibitedDay');
        if (addProhibitedDayButton) {
            addProhibitedDayButton.addEventListener('click', () => {
                this.addProhibitedDay();
            });
        }
        // 月移動ボタンのイベントリスナー
        const prevMonthPublicButton = document.getElementById('prevMonthPublic');
        const nextMonthPublicButton = document.getElementById('nextMonthPublic');
        if (prevMonthPublicButton) {
            prevMonthPublicButton.addEventListener('click', () => {
                this.changeMonth('public', -1);
            });
        }
        if (nextMonthPublicButton) {
            nextMonthPublicButton.addEventListener('click', () => {
                this.changeMonth('public', 1);
            });
        }
        const prevMonthProhibitedButton = document.getElementById('prevMonthProhibited');
        const nextMonthProhibitedButton = document.getElementById('nextMonthProhibited');
        if (prevMonthProhibitedButton) {
            prevMonthProhibitedButton.addEventListener('click', () => {
                this.changeMonth('prohibited', -1);
            });
        }
        if (nextMonthProhibitedButton) {
            nextMonthProhibitedButton.addEventListener('click', () => {
                this.changeMonth('prohibited', 1);
            });
        }
        // 動的に追加される削除ボタンのイベントリスナー（委譲）
        const modalBody = modal?.querySelector('.modal-body');
        if (modalBody) {
            modalBody.addEventListener('click', (e) => {
                const target = e.target;
                if (target.classList.contains('holiday-remove')) {
                    const holidayItem = target.closest('.holiday-item');
                    if (holidayItem) {
                        const dateText = holidayItem.querySelector('.holiday-date')?.textContent;
                        const type = holidayItem.classList.contains('public') ? 'public' : 'prohibited';
                        if (dateText) {
                            const date = new Date(dateText);
                            this.removeHolidayDate(date, type);
                        }
                    }
                }
            });
        }
    }
    addPublicHoliday() {
        const dateInput = document.getElementById('publicHolidayDate');
        if (dateInput && dateInput.value) {
            const date = new Date(dateInput.value);
            this.calendarManager.addPublicHoliday(date);
            this.loadHolidaySettings(); // リストを更新
        }
    }
    addProhibitedDay() {
        const dateInput = document.getElementById('prohibitedDayDate');
        if (dateInput && dateInput.value) {
            const date = new Date(dateInput.value);
            this.calendarManager.addProhibitedDay(date);
            this.loadHolidaySettings(); // リストを更新
        }
    }
    removeHolidayDate(date, type) {
        if (type === 'public') {
            this.calendarManager.removePublicHoliday(date);
        }
        else {
            this.calendarManager.removeProhibitedDay(date);
        }
        this.loadHolidaySettings(); // リストを更新
    }
    // 月を変更
    changeMonth(type, direction) {
        const containerId = type === 'public' ? 'publicHolidayCalendar' : 'prohibitedDayCalendar';
        const state = this.calendarStates[containerId];
        if (state) {
            state.currentMonth += direction;
            // 年をまたぐ場合の処理
            if (state.currentMonth < 0) {
                state.currentMonth = 11;
                state.currentYear--;
            }
            else if (state.currentMonth > 11) {
                state.currentMonth = 0;
                state.currentYear++;
            }
            // カレンダーを再生成
            this.generateCalendar(containerId, type);
        }
    }
    loadHolidaySettings() {
        // 現在の設定をモーダルに表示
        const settings = this.calendarManager.getSettings();
        this.displayHolidaySettings(settings);
        // 日付入力を現在の日付に設定
        const today = new Date();
        const todayStr = today.toISOString().split('T')[0];
        const publicHolidayDateInput = document.getElementById('publicHolidayDate');
        const prohibitedDayDateInput = document.getElementById('prohibitedDayDate');
        if (publicHolidayDateInput) {
            publicHolidayDateInput.value = todayStr;
        }
        if (prohibitedDayDateInput) {
            prohibitedDayDateInput.value = todayStr;
        }
    }
    saveHolidaySettings() {
        // モーダルから設定を取得して保存
        const settings = this.getHolidaySettingsFromModal();
        this.calendarManager.updateSettings(settings);
        // ReportGeneratorにも設定を反映
        this.reportGenerator.updateHolidaySettings(settings);
        // 設定変更を通知
        this.notifyHolidaySettingsChanged();
        // モーダルを閉じる
        const modal = document.getElementById('holidaySettingsModal');
        if (modal) {
            const bootstrapModal = window.bootstrap?.Modal.getInstance(modal);
            if (bootstrapModal) {
                bootstrapModal.hide();
            }
        }
    }
    displayHolidaySettings(settings) {
        // 公休日リストの表示
        const publicHolidayList = document.getElementById('publicHolidayList');
        if (publicHolidayList) {
            publicHolidayList.innerHTML = settings.publicHolidays.length === 0
                ? '<p class="text-muted">設定された公休日はありません</p>'
                : settings.publicHolidays
                    .sort((a, b) => a.getTime() - b.getTime())
                    .map(date => `
                        <div class="holiday-item public">
                            <span class="holiday-date">${this.formatDate(date)}</span>
                            <button class="holiday-remove" data-date="${date.toISOString()}" data-type="public">
                                <i class="fas fa-times"></i>
                            </button>
                        </div>
                    `).join('');
        }
        // 禁止日リストの表示
        const prohibitedDayList = document.getElementById('prohibitedDayList');
        if (prohibitedDayList) {
            prohibitedDayList.innerHTML = settings.prohibitedDays.length === 0
                ? '<p class="text-muted">設定された禁止日はありません</p>'
                : settings.prohibitedDays
                    .sort((a, b) => a.getTime() - b.getTime())
                    .map(date => `
                        <div class="holiday-item prohibited">
                            <span class="holiday-date">${this.formatDate(date)}</span>
                            <button class="holiday-remove" data-date="${date.toISOString()}" data-type="prohibited">
                                <i class="fas fa-times"></i>
                            </button>
                        </div>
                    `).join('');
        }
    }
    getHolidaySettingsFromModal() {
        // 現在の設定を取得
        return this.calendarManager.getSettings();
    }
    notifyHolidaySettingsChanged() {
        // 設定変更を通知
        const event = new CustomEvent('holidaySettingsChanged', {
            detail: this.calendarManager.getSettings()
        });
        document.dispatchEvent(event);
    }
    formatDate(date) {
        return date.toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
    }
    loadSavedHolidaySettings() {
        // 保存された設定があれば読み込み、なければデフォルト設定を追加
        const settings = this.calendarManager.getSettings();
        if (settings.publicHolidays.length === 0 && settings.prohibitedDays.length === 0) {
            // デフォルト設定を追加（画像の設定に基づく）
            this.addDefaultHolidaySettings();
        }
        // 設定をReportGeneratorに反映
        this.reportGenerator.updateHolidaySettings(settings);
    }
    addDefaultHolidaySettings() {
        // 画像の設定に基づくデフォルト値を追加
        const defaultPublicHolidays = [
            new Date(2025, 7, 24), // 8月24日
            new Date(2025, 7, 25), // 8月25日
            new Date(2025, 7, 31) // 8月31日
        ];
        const defaultProhibitedDays = [
            new Date(2025, 7, 14), // 8月14日
            new Date(2025, 7, 15), // 8月15日
            new Date(2025, 7, 16), // 8月16日
            new Date(2025, 8, 1) // 9月1日
        ];
        // デフォルト設定を追加
        defaultPublicHolidays.forEach(date => {
            this.calendarManager.addPublicHoliday(date);
        });
        defaultProhibitedDays.forEach(date => {
            this.calendarManager.addProhibitedDay(date);
        });
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
            // 担当別データのCSV出力ボタンのイベントリスナーを設定
            this.setupStaffDataExportButton();
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
            // 担当別データも月報データで更新
            this.updateStaffData(report.rawData);
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
        // 担当別データのCSV出力ボタンのイベントリスナーを設定
        this.setupStaffDataExportButton();
    }
    // 担当別データを更新
    updateStaffData(monthlyData) {
        let targetData = this.currentData;
        // 月報データが指定されている場合は、その月のデータのみを使用
        if (monthlyData && monthlyData.length > 0) {
            targetData = monthlyData;
        }
        if (targetData && targetData.length > 0) {
            // 月報の場合はその月の最初の日、そうでなければ現在の日付を使用
            let targetDate;
            if (monthlyData && monthlyData.length > 0) {
                const reportMonthInput = document.getElementById('reportMonth');
                if (reportMonthInput && reportMonthInput.value) {
                    const [year, month] = reportMonthInput.value.split('-').map(Number);
                    targetDate = new Date(year, month - 1, 1);
                }
                else {
                    targetDate = new Date();
                }
            }
            else {
                targetDate = new Date();
            }
            const staffData = this.reportGenerator.generateStaffData(targetData, targetDate);
            const staffContainer = document.getElementById('staffDataContent');
            if (staffContainer) {
                // 月情報を表示するヘッダーを追加
                const monthHeader = this.createMonthHeaderForStaffData();
                staffContainer.innerHTML = monthHeader + this.reportGenerator.createStaffDataHTML(staffData);
            }
            else {
                console.error('staffDataContent要素が見つかりません');
            }
        }
        else {
            const staffContainer = document.getElementById('staffDataContent');
            if (staffContainer) {
                staffContainer.innerHTML = '<div class="alert alert-info">データがありません</div>';
            }
        }
    }
    // 担当別データ用の月ヘッダーを生成
    createMonthHeaderForStaffData() {
        const reportMonthInput = document.getElementById('reportMonth');
        if (reportMonthInput && reportMonthInput.value) {
            const [year, month] = reportMonthInput.value.split('-').map(Number);
            const monthText = `${year}年${month}月`;
            return `
                <div class="alert alert-info mb-3">
                    <h5><i class="fas fa-calendar-alt me-2"></i>対象期間: ${monthText}</h5>
                </div>
            `;
        }
        return `
            <div class="alert alert-info mb-3">
                <h5><i class="fas fa-calendar-alt me-2"></i>対象期間: 全期間</h5>
            </div>
        `;
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
    // 担当別データのCSV出力ボタンのイベントリスナーを設定
    setupStaffDataExportButton() {
        const exportButton = document.querySelector('.btn-export-staff-data');
        if (exportButton) {
            exportButton.addEventListener('click', async () => {
                try {
                    // 現在の担当別データを取得
                    const currentData = this.currentData;
                    if (!currentData || currentData.length === 0) {
                        this.showMessage('データが読み込まれていません。', 'error');
                        return;
                    }
                    // 月報データが存在する場合は、その月のデータを使用
                    const monthlyData = this.getCurrentMonthlyData();
                    const targetData = monthlyData && monthlyData.length > 0 ? monthlyData : currentData;
                    // 月報の場合はその月の最初の日、そうでなければ現在の日付を使用
                    let targetDate;
                    if (monthlyData && monthlyData.length > 0) {
                        const reportMonthInput = document.getElementById('reportMonth');
                        if (reportMonthInput && reportMonthInput.value) {
                            const [year, month] = reportMonthInput.value.split('-').map(Number);
                            targetDate = new Date(year, month - 1, 1);
                        }
                        else {
                            targetDate = new Date();
                        }
                    }
                    else {
                        targetDate = new Date();
                    }
                    const staffData = this.reportGenerator.generateStaffData(targetData, targetDate);
                    await this.reportGenerator.exportStaffDataToCSV(staffData);
                    this.showMessage('担当別データのCSV出力が完了しました。', 'success');
                }
                catch (error) {
                    console.error('担当別データCSV出力エラー:', error);
                    this.showMessage('CSVの出力に失敗しました。', 'error');
                }
            });
        }
    }
    // 現在の月報データを取得
    getCurrentMonthlyData() {
        const reportMonthInput = document.getElementById('reportMonth');
        if (reportMonthInput && reportMonthInput.value) {
            const data = this.dataManager.getData();
            if (data.length > 0) {
                try {
                    const monthString = reportMonthInput.value;
                    const report = this.reportGenerator.generateMonthlyReport(data, monthString);
                    return report.rawData || null;
                }
                catch (error) {
                    console.error('月報データ取得エラー:', error);
                    return null;
                }
            }
        }
        return null;
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
    // カレンダー生成
    generateCalendar(containerId, type) {
        const container = document.getElementById(containerId);
        if (!container)
            return;
        const currentDate = new Date();
        const currentMonth = currentDate.getMonth();
        const currentYear = currentDate.getFullYear();
        // カレンダーの状態を保存
        if (!this.calendarStates) {
            this.calendarStates = {};
        }
        if (!this.calendarStates[containerId]) {
            this.calendarStates[containerId] = {
                currentMonth: currentMonth,
                currentYear: currentYear,
                selectedDates: new Set()
            };
        }
        const state = this.calendarStates[containerId];
        const month = state.currentMonth;
        const year = state.currentYear;
        // カレンダーHTMLを生成
        const calendarHTML = this.createCalendarHTML(year, month, type, state.selectedDates);
        container.innerHTML = calendarHTML;
        // 追加ボタンをカレンダーグリッドの外側に配置
        const addButtonHTML = this.createAddButtonHTML(type, state.selectedDates);
        if (addButtonHTML) {
            // 既存のボタンがあれば削除
            const existingButton = container.parentElement?.querySelector(`#addSelectedDates_${type}`);
            if (existingButton) {
                existingButton.remove();
            }
            // カレンダーグリッドの下にボタンを挿入
            const buttonContainer = document.createElement('div');
            buttonContainer.innerHTML = addButtonHTML;
            container.parentElement?.insertBefore(buttonContainer, container.nextSibling);
        }
        // 日付クリックイベントを設定
        const dayElements = container.querySelectorAll('.day[data-date]');
        dayElements.forEach(dayElement => {
            dayElement.addEventListener('click', (e) => {
                const target = e.target;
                const dateStr = target.dataset.date;
                if (dateStr) {
                    this.toggleDateSelection(containerId, dateStr, type);
                }
            });
        });
        // 追加ボタンのイベントリスナーを設定
        const addButton = container.parentElement?.querySelector(`#addSelectedDates_${type}`);
        if (addButton) {
            addButton.addEventListener('click', () => {
                this.addSelectedDates(containerId, type);
            });
        }
    }
    // カレンダーHTML生成
    createCalendarHTML(year, month, type, selectedDates) {
        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);
        const startDate = new Date(firstDay);
        startDate.setDate(startDate.getDate() - firstDay.getDay());
        const settings = this.calendarManager.getSettings();
        const holidayDates = type === 'public' ? settings.publicHolidays : settings.prohibitedDays;
        let html = '';
        const today = new Date();
        let currentDate = new Date(startDate);
        // 曜日ヘッダー行（日曜日始まり）
        html += '<div class="calendar-week">';
        html += '<div class="day weekday">日</div>';
        html += '<div class="day weekday">月</div>';
        html += '<div class="day weekday">火</div>';
        html += '<div class="day weekday">水</div>';
        html += '<div class="day weekday">木</div>';
        html += '<div class="day weekday">金</div>';
        html += '<div class="day weekday">土</div>';
        html += '</div>';
        // 日付行を生成
        for (let week = 0; week < 6; week++) {
            html += '<div class="calendar-week">';
            for (let day = 0; day < 7; day++) {
                const dateString = currentDate.toISOString().split('T')[0];
                const isCurrentMonth = currentDate.getMonth() === month;
                const isToday = this.isSameDate(currentDate, today);
                const isSelected = selectedDates.has(dateString);
                const isHoliday = holidayDates.some(h => this.isSameDate(h, currentDate));
                let classes = 'day';
                if (!isCurrentMonth)
                    classes += ' other-month';
                if (isToday)
                    classes += ' today';
                if (isSelected)
                    classes += ' selected';
                if (isHoliday) {
                    classes += type === 'public' ? ' public' : ' prohibited';
                }
                html += `<div class="${classes}" data-date="${dateString}">${currentDate.getDate()}</div>`;
                currentDate.setDate(currentDate.getDate() + 1);
            }
            html += '</div>';
        }
        return html;
    }
    // 選択された日付の追加ボタンを生成（カレンダーグリッドの外側）
    createAddButtonHTML(type, selectedDates) {
        if (selectedDates.size === 0)
            return '';
        const buttonColor = type === 'public' ? 'success' : 'danger';
        return `
            <div class="mt-3 text-center">
                <button class="btn btn-${buttonColor} btn-sm" id="addSelectedDates_${type}">
                    <i class="fas fa-plus me-1"></i>選択された${selectedDates.size}日を追加
                </button>
            </div>
        `;
    }
    // 日付選択の切り替え
    toggleDateSelection(containerId, dateStr, type) {
        const state = this.calendarStates[containerId];
        if (!state)
            return;
        if (state.selectedDates.has(dateStr)) {
            state.selectedDates.delete(dateStr);
        }
        else {
            state.selectedDates.add(dateStr);
        }
        // カレンダーを再生成
        this.generateCalendar(containerId, type);
    }
    // 選択された日付を一括追加
    addSelectedDates(containerId, type) {
        const settings = this.calendarManager.getSettings();
        const state = this.calendarStates[containerId];
        if (!state)
            return;
        state.selectedDates.forEach(dateStr => {
            const date = new Date(dateStr);
            if (type === 'public') {
                if (!settings.publicHolidays.some(d => this.isSameDate(d, date))) {
                    settings.publicHolidays.push(date);
                }
            }
            else {
                if (!settings.prohibitedDays.some(d => this.isSameDate(d, date))) {
                    settings.prohibitedDays.push(date);
                }
            }
        });
        // 設定を保存
        this.calendarManager.updateSettings(settings);
        // 表示を更新
        this.displayHolidaySettings(settings);
        // 選択状態をクリア
        if (this.calendarStates) {
            if (this.calendarStates[containerId]) {
                this.calendarStates[containerId].selectedDates.clear();
            }
        }
        // カレンダーを再生成
        this.generateCalendar(containerId, type);
    }
    // 日付比較
    isSameDate(date1, date2) {
        return date1.getFullYear() === date2.getFullYear() &&
            date1.getMonth() === date2.getMonth() &&
            date1.getDate() === date2.getDate();
    }
}
// アプリケーションの初期化
document.addEventListener('DOMContentLoaded', () => {
    new App();
});
//# sourceMappingURL=app.js.map