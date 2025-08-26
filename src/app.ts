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
    private calendarStates: { [key: string]: { currentMonth: number; currentYear: number; selectedDates: Set<string> } } = {};

    constructor() {
        this.excelProcessor = new ExcelProcessor();
        this.calendarManager = new CalendarManager();
        this.reportGenerator = new ReportGenerator(this.excelProcessor, this.calendarManager);
        this.dataManager = DataManager.getInstance();
        
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
        // 公休日・禁止日設定の初期化
        this.setupHolidaySettingsModal();
    }

    // 公休日・禁止日設定モーダルの設定
    private setupHolidaySettingsModal(): void {
        // 削除ボタンのイベントリスナーを設定
        this.setupHolidayRemoveButtons();

        // 月変更ボタンのイベントリスナーを設定
        const prevMonthPublic = document.getElementById('prevMonthPublic');
        const nextMonthPublic = document.getElementById('nextMonthPublic');
        const prevMonthProhibited = document.getElementById('prevMonthProhibited');
        const nextMonthProhibited = document.getElementById('nextMonthProhibited');

        if (prevMonthPublic) {
            prevMonthPublic.addEventListener('click', () => this.changeMonth('public', -1));
        }
        if (nextMonthPublic) {
            nextMonthPublic.addEventListener('click', () => this.changeMonth('public', 1));
        }
        if (prevMonthProhibited) {
            prevMonthProhibited.addEventListener('click', () => this.changeMonth('prohibited', -1));
        }
        if (nextMonthProhibited) {
            nextMonthProhibited.addEventListener('click', () => this.changeMonth('prohibited', 1));
        }

        // 設定保存ボタンのイベントリスナーを設定
        const saveButton = document.getElementById('saveHolidaySettings');
        if (saveButton) {
            saveButton.addEventListener('click', () => this.saveHolidaySettings());
        }

        // 個別追加ボタンのイベントリスナーを設定
        const addPublicHolidayButton = document.getElementById('addPublicHoliday');
        if (addPublicHolidayButton) {
            addPublicHolidayButton.addEventListener('click', () => this.addPublicHoliday());
        }

        const addProhibitedDayButton = document.getElementById('addProhibitedDay');
        if (addProhibitedDayButton) {
            addProhibitedDayButton.addEventListener('click', () => this.addProhibitedDay());
        }

        // モーダルが表示されたときの処理
        const modal = document.getElementById('holidaySettingsModal');
        if (modal) {
            modal.addEventListener('shown.bs.modal', () => {
                // カレンダーを生成
                this.generateCalendar('publicHolidayCalendar', 'public');
                this.generateCalendar('prohibitedDayCalendar', 'prohibited');
                
                // 設定済み日付一覧を表示
                this.loadHolidaySettings();
                
                // 削除ボタンのイベントリスナーを再設定
                this.setupHolidayRemoveButtons();
                
                console.log('モーダル表示完了 - イベントリスナー設定済み');
            });
        }
    }

    // 削除ボタンのイベントリスナーを設定
    private setupHolidayRemoveButtons(): void {
        // 公休日削除ボタンのイベントリスナー
        const publicRemoveButtons = document.querySelectorAll('.holiday-remove[data-type="public"]');
        publicRemoveButtons.forEach(button => {
            // 既存のイベントリスナーを削除
            button.removeEventListener('click', this.handleHolidayRemove);
            // 新しいイベントリスナーを追加
            button.addEventListener('click', this.handleHolidayRemove.bind(this));
        });

        // 禁止日削除ボタンのイベントリスナー
        const prohibitedRemoveButtons = document.querySelectorAll('.holiday-remove[data-type="prohibited"]');
        prohibitedRemoveButtons.forEach(button => {
            // 既存のイベントリスナーを削除
            button.removeEventListener('click', this.handleHolidayRemove);
            // 新しいイベントリスナーを追加
            button.addEventListener('click', this.handleHolidayRemove.bind(this));
        });
        
        console.log(`削除ボタンのイベントリスナーを設定: 公休日${publicRemoveButtons.length}個, 禁止日${prohibitedRemoveButtons.length}個`);
    }

    // 削除ボタンのクリックイベントハンドラー
    private handleHolidayRemove(e: Event): void {
        const target = e.target as HTMLElement;
        const buttonElement = target.closest('.holiday-remove') as HTMLElement;
        
        if (buttonElement) {
            const dateStr = buttonElement.dataset.date;
            const type = buttonElement.dataset.type as 'public' | 'prohibited';
            
            if (dateStr && type) {
                console.log(`削除ボタンクリック: ${dateStr}, タイプ: ${type}`);
                
                // 日付文字列を正しく解析
                const [year, month, day] = dateStr.split('-').map(Number);
                const date = new Date(year, month - 1, day); // monthは0ベースなので-1
                
                this.removeHolidayDate(date, type);
            } else {
                console.error('削除ボタンのデータ属性が不正:', { dateStr, type });
            }
        } else {
            console.error('削除ボタンが見つかりません');
        }
    }

    private addPublicHoliday(): void {
        const dateInput = document.getElementById('publicHolidayDate') as HTMLInputElement;
        if (dateInput && dateInput.value) {
            // 日付文字列を正しく解析（タイムゾーンの影響を排除）
            const [year, month, day] = dateInput.value.split('-').map(Number);
            const date = new Date(year, month - 1, day); // monthは0ベースなので-1
            
            // 既存の設定を確認
            const settings = this.calendarManager.getSettings();
            if (settings.publicHolidays.some(d => this.isSameDate(d, date))) {
                this.showMessage('この日付は既に設定されています。', 'info');
                return;
            }
            
            // CalendarManagerに追加
            this.calendarManager.addPublicHoliday(date);
            
            // DataManagerにも同期
            const updatedSettings = this.calendarManager.getSettings();
            this.dataManager.setHolidaySettings(updatedSettings);
            
            // ReportGeneratorにも設定を反映
            this.reportGenerator.updateHolidaySettings(updatedSettings);
            
            // 設定済み日付一覧を更新
            this.loadHolidaySettings();
            
            // カレンダーを再生成
            this.generateCalendar('publicHolidayCalendar', 'public');
            
            // 成功メッセージを表示
            this.showMessage(`公休日「${dateInput.value}」が正常に追加されました。`, 'success');
            console.log(`公休日を追加: ${dateInput.value}`);
        }
    }

    private addProhibitedDay(): void {
        const dateInput = document.getElementById('prohibitedDayDate') as HTMLInputElement;
        if (dateInput && dateInput.value) {
            // 日付文字列を正しく解析（タイムゾーンの影響を排除）
            const [year, month, day] = dateInput.value.split('-').map(Number);
            const date = new Date(year, month - 1, day); // monthは0ベースなので-1
            
            // 既存の設定を確認
            const settings = this.calendarManager.getSettings();
            if (settings.prohibitedDays.some(d => this.isSameDate(d, date))) {
                this.showMessage('この日付は既に設定されています。', 'info');
                return;
            }
            
            // CalendarManagerに追加
            this.calendarManager.addProhibitedDay(date);
            
            // DataManagerにも同期
            const updatedSettings = this.calendarManager.getSettings();
            this.dataManager.setHolidaySettings(updatedSettings);
            
            // ReportGeneratorにも設定を反映
            this.reportGenerator.updateHolidaySettings(updatedSettings);
            
            // 設定済み日付一覧を更新
            this.loadHolidaySettings();
            
            // カレンダーを再生成
            this.generateCalendar('prohibitedDayCalendar', 'prohibited');
            
            // 成功メッセージを表示
            this.showMessage(`禁止日「${dateInput.value}」が正常に追加されました。`, 'success');
            console.log(`禁止日を追加: ${dateInput.value}`);
        }
    }

    private removeHolidayDate(date: Date, type: 'public' | 'prohibited'): void {
        const dateStr = this.formatDate(date);
        
        console.log(`${type === 'public' ? '公休日' : '禁止日'}を削除開始: ${dateStr}`);
        
        if (type === 'public') {
            this.calendarManager.removePublicHoliday(date);
            console.log(`公休日を削除: ${dateStr}`);
        } else {
            this.calendarManager.removeProhibitedDay(date);
            console.log(`禁止日を削除: ${dateStr}`);
        }
        
        // 設定を取得してDataManagerにも同期
        const settings = this.calendarManager.getSettings();
        this.dataManager.setHolidaySettings(settings);
        
        // ReportGeneratorにも設定を反映
        this.reportGenerator.updateHolidaySettings(settings);
        
        // 設定済み日付一覧を更新
        this.loadHolidaySettings();
        
        // カレンダーを再生成して削除されたことを表示
        const containerId = type === 'public' ? 'publicHolidayCalendar' : 'prohibitedDayCalendar';
        this.generateCalendar(containerId, type);
        
        // 削除完了メッセージを表示
        const message = type === 'public' ? '公休日' : '禁止日';
        this.showMessage(`${message}「${dateStr}」が削除されました。`, 'success');
        
        console.log(`${type === 'public' ? '公休日' : '禁止日'}の削除完了: ${dateStr}`);
    }

    // 月を変更
    private changeMonth(type: 'public' | 'prohibited', direction: number): void {
        const containerId = type === 'public' ? 'publicHolidayCalendar' : 'prohibitedDayCalendar';
        const state = this.calendarStates[containerId];
        
        if (state) {
            state.currentMonth += direction;
            
            // 年をまたぐ場合の処理
            if (state.currentMonth < 0) {
                state.currentMonth = 11;
                state.currentYear--;
            } else if (state.currentMonth > 11) {
                state.currentMonth = 0;
                state.currentYear++;
            }
            
            // ヘッダーの月表示を更新
            this.updateMonthDisplay(type, state.currentYear, state.currentMonth);
            
            // カレンダーを再生成
            this.generateCalendar(containerId, type);
        }
    }

    // ヘッダーの月表示を更新
    private updateMonthDisplay(type: 'public' | 'prohibited', year: number, month: number): void {
        const monthDisplayId = type === 'public' ? 'publicMonthDisplay' : 'prohibitedMonthDisplay';
        const monthDisplay = document.getElementById(monthDisplayId);
        if (monthDisplay) {
            monthDisplay.textContent = `${year}年${month + 1}月`;
        }
    }

    private loadHolidaySettings(): void {
        // CalendarManagerから現在の設定を取得
        const settings = this.calendarManager.getSettings();
        
        // DataManagerにも設定を同期
        this.dataManager.setHolidaySettings(settings);
        
        // 現在の設定をモーダルに表示
        this.displayHolidaySettings(settings);
        
        // 日付入力を現在の日付に設定
        const today = new Date();
        const todayStr = today.toISOString().split('T')[0];
        
        const publicHolidayDateInput = document.getElementById('publicHolidayDate') as HTMLInputElement;
        const prohibitedDayDateInput = document.getElementById('prohibitedDayDate') as HTMLInputElement;
        
        if (publicHolidayDateInput) {
            publicHolidayDateInput.value = todayStr;
        }
        if (prohibitedDayDateInput) {
            prohibitedDayDateInput.value = todayStr;
        }
        
        // 削除ボタンのイベントリスナーを再設定
        this.setupHolidayRemoveButtons();
        
        // デバッグ用ログ
        console.log('設定を読み込み:', {
            publicHolidays: settings.publicHolidays.map(d => d.toISOString().split('T')[0]),
            prohibitedDays: settings.prohibitedDays.map(d => d.toISOString().split('T')[0])
        });
    }

    private saveHolidaySettings(): void {
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
            const bootstrapModal = (window as any).bootstrap?.Modal.getInstance(modal);
            if (bootstrapModal) {
                bootstrapModal.hide();
            }
        }
    }

    private displayHolidaySettings(settings: HolidaySettings): void {
        // 公休日リストの表示
        const publicHolidayList = document.getElementById('publicHolidayList');
        if (publicHolidayList) {
            publicHolidayList.innerHTML = settings.publicHolidays.length === 0 
                ? '<p class="text-muted">設定された公休日はありません</p>'
                : settings.publicHolidays
                    .sort((a, b) => a.getTime() - b.getTime())
                    .map(date => {
                        // 日付文字列を正しく生成（タイムゾーンの影響を排除）
                        const dateStr = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
                        return `
                            <div class="holiday-item public">
                                <span class="holiday-date">${this.formatDate(date)}</span>
                                <button class="holiday-remove" data-date="${dateStr}" data-type="public">
                                    <i class="fas fa-times"></i>
                                </button>
                            </div>
                        `;
                    }).join('');
        }

        // 禁止日リストの表示
        const prohibitedDayList = document.getElementById('prohibitedDayList');
        if (prohibitedDayList) {
            prohibitedDayList.innerHTML = settings.prohibitedDays.length === 0 
                ? '<p class="text-muted">設定された禁止日はありません</p>'
                : settings.prohibitedDays
                    .sort((a, b) => a.getTime() - b.getTime())
                    .map(date => {
                        // 日付文字列を正しく生成（タイムゾーンの影響を排除）
                        const dateStr = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
                        return `
                            <div class="holiday-item prohibited">
                                <span class="holiday-date">${this.formatDate(date)}</span>
                                <button class="holiday-remove" data-date="${dateStr}" data-type="prohibited">
                                    <i class="fas fa-times"></i>
                                </button>
                            </div>
                        `;
                    }).join('');
        }
        
        // デバッグ用ログ
        console.log('設定済み日付一覧を表示:', {
            publicHolidays: settings.publicHolidays.map(d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`),
            prohibitedDays: settings.prohibitedDays.map(d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`)
        });
    }

    private getHolidaySettingsFromModal(): HolidaySettings {
        // 現在の設定を取得
        return this.calendarManager.getSettings();
    }

    private notifyHolidaySettingsChanged(): void {
        // 設定変更を通知
        const event = new CustomEvent('holidaySettingsChanged', {
            detail: this.calendarManager.getSettings()
        });
        document.dispatchEvent(event);
    }

    private formatDate(date: Date): string {
        return date.toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
    }
    
    private loadSavedHolidaySettings(): void {
        // 保存された設定があれば読み込み、なければデフォルト設定を追加
        const settings = this.calendarManager.getSettings();
        
        if (settings.publicHolidays.length === 0 && settings.prohibitedDays.length === 0) {
            // デフォルト設定を追加（画像の設定に基づく）
            this.addDefaultHolidaySettings();
        }
        
        // 設定をReportGeneratorに反映
        this.reportGenerator.updateHolidaySettings(settings);
    }

    private addDefaultHolidaySettings(): void {
        // 画像の設定に基づくデフォルト値を追加
        const defaultPublicHolidays = [
            new Date(2025, 7, 24), // 8月24日
            new Date(2025, 7, 25), // 8月25日
            new Date(2025, 7, 31)  // 8月31日
        ];
        
        const defaultProhibitedDays = [
            new Date(2025, 7, 14), // 8月14日
            new Date(2025, 7, 15), // 8月15日
            new Date(2025, 7, 16), // 8月16日
            new Date(2025, 8, 1)   // 9月1日
        ];

        // デフォルト設定を追加
        defaultPublicHolidays.forEach(date => {
            this.calendarManager.addPublicHoliday(date);
        });
        
        defaultProhibitedDays.forEach(date => {
            this.calendarManager.addProhibitedDay(date);
        });
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
            
            // 月報も自動生成して担当別データと連動
            this.generateMonthlyReportForFileDrop();
            
            // 完了メッセージを表示
            this.showMessage(`ファイル「${file.name}」の読み込みが完了しました。総データ件数: ${data.length}件`, 'success');
            
            // 担当別データを生成
            this.updateStaffData();
            
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
        
        const reportMonthInput = document.getElementById('reportMonth') as HTMLInputElement;
        if (!reportMonthInput || !reportMonthInput.value) {
            this.showMessage('月報作成月を選択してください。', 'error');
            return;
        }
        
        try {
            // 選択された月から年月を取得
            const monthString = reportMonthInput.value; // 既に "YYYY-MM" 形式
            
            const report = this.reportGenerator.generateMonthlyReport(data, monthString);
            // 月報を表示
            this.displayReport(report, 'monthly');
            
            // 担当別データも月報データで更新
            this.updateStaffDataWithMonthlyData(report.rawData, monthString);
            
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
        
        // 担当別データも更新（月報の場合は月報データを渡す）
        if (type === 'monthly' && report.rawData) {
            this.updateStaffData(report.rawData);
        } else {
            this.updateStaffData();
        }
        
        // エクスポートボタンのイベントリスナーを設定
        this.setupExportButtons(type, report);
        
        // 担当別データのCSV出力ボタンのイベントリスナーを設定
        this.setupStaffDataExportButton();
    }

    // データの月分布を取得
    private getMonthDistribution(data: any[]): { [key: string]: number } {
        const distribution: { [key: string]: number } = {};
        
        data.forEach(row => {
            if (row.date) {
                const year = row.date.getFullYear();
                const month = row.date.getMonth() + 1; // 0ベースなので+1
                const key = `${year}年${month}月`;
                distribution[key] = (distribution[key] || 0) + 1;
            }
        });
        
        return distribution;
    }

    // 担当別データを更新
    private updateStaffData(monthlyData?: any[]): void {
        // ドロップされたファイルのデータを直接使用
        const rawData = this.currentData || [];
        if (rawData.length === 0) {
            console.log('Excelデータがありません');
            return;
        }

        // ドロップされたデータの最初の行から月を取得
        const firstRow = rawData[0];
        if (firstRow && firstRow.date) {
            const year = firstRow.date.getFullYear();
            const month = firstRow.date.getMonth();
            const targetDate = new Date(year, month, 15); // 月の15日を基準に設定
            
            const staffData = this.reportGenerator.generateStaffData(rawData, targetDate);
            
            // 担当別データのHTMLを生成して表示
            const staffContainer = document.getElementById('staffDataContent');
            if (staffContainer) {
                const monthHeader = this.createMonthHeaderForStaffData();
                staffContainer.innerHTML = monthHeader + this.reportGenerator.createStaffDataHTML(staffData);
            }
            
            this.setupStaffDataSearchAndSort();
        } else {
            console.log('ドロップデータから日付を取得できません');
        }
    }

    // 担当別データを月報データで更新
    private updateStaffDataWithMonthlyData(monthlyData: any[], monthString: string): void {
        if (monthlyData.length === 0) {
            return;
        }

        // 月報データの最初の行から月を取得
        const firstRow = monthlyData[0];
        if (firstRow && firstRow.date) {
            const year = firstRow.date.getFullYear();
            const month = firstRow.date.getMonth();
            const targetDate = new Date(year, month, 15); // 月の15日を基準に設定
            
            const staffData = this.reportGenerator.generateStaffData(monthlyData, targetDate);
            
            // 担当別データのHTMLを生成して表示
            const staffContainer = document.getElementById('staffDataContent');
            if (staffContainer) {
                const monthHeader = this.createMonthHeaderForStaffData(monthString);
                staffContainer.innerHTML = monthHeader + this.reportGenerator.createStaffDataHTML(staffData);
            }
            
            this.setupStaffDataSearchAndSort();
        }
    }

    // 担当別データ用の月ヘッダーを生成
    private createMonthHeaderForStaffData(monthString?: string): string {
        if (!monthString) {
            return `
                <div class="alert alert-info mb-3">
                    <h5><i class="fas fa-calendar-alt me-2"></i>対象期間: 全期間</h5>
                </div>
            `;
        }
        const [year, month] = monthString.split('-').map(Number);
        const monthText = `${year}年${month}月`;
        return `
            <div class="alert alert-info mb-3">
                <h5><i class="fas fa-calendar-alt me-2"></i>対象期間: ${monthText}</h5>
            </div>
        `;
    }

    // 確認データを更新
    private updateDataConfirmation(): void {

        if (this.currentData && this.currentData.length > 0) {
            const dataContainer = document.getElementById('dataConfirmationContent');
            if (dataContainer) {
                dataContainer.innerHTML = this.reportGenerator.createDataConfirmationHTML(this.currentData);
        
                
                // フィルター機能を設定
                this.setupDataFilters();
            } else {
                console.error('dataConfirmationContent要素が見つかりません');
            }
        } else {
    
        }
    }
    
    // 担当別データの検索・ソート機能を設定
    private setupStaffDataSearchAndSort(): void {
        // 検索機能の設定
        this.setupStaffDataSearch();
        
        // ソート機能の設定
        this.setupStaffDataSort();
    }
    
    // 担当別データの検索機能を設定
    private setupStaffDataSearch(): void {
        const searchExecute = document.getElementById('searchExecute');
        const searchClear = document.getElementById('searchClear');
        const staffNameSearch = document.getElementById('staffNameSearch') as HTMLInputElement;
        const regionSearch = document.getElementById('regionSearch') as HTMLInputElement;
        const departmentSearch = document.getElementById('departmentSearch') as HTMLInputElement;
        const ordersSearch = document.getElementById('ordersSearch') as HTMLSelectElement;
        const elderlySearch = document.getElementById('elderlySearch') as HTMLSelectElement;
        
        if (searchExecute) {
            searchExecute.addEventListener('click', () => {
                this.executeStaffDataSearch();
            });
        }
        
        if (searchClear) {
            searchClear.addEventListener('click', () => {
                this.clearStaffDataSearch();
            });
        }
        
        // リアルタイム検索（入力と同時に検索実行）
        [staffNameSearch, regionSearch, departmentSearch, ordersSearch, elderlySearch].forEach(element => {
            if (element) {
                element.addEventListener('input', () => {
                    this.executeStaffDataSearch();
                });
                element.addEventListener('change', () => {
                    this.executeStaffDataSearch();
                });
            }
        });
    }
    
    // 担当別データのソート機能を設定
    private setupStaffDataSort(): void {
        const sortableHeaders = document.querySelectorAll('#staffDataTable th.sortable');
        
        sortableHeaders.forEach(header => {
            header.addEventListener('click', (e) => {
                const target = e.currentTarget as HTMLElement;
                const sortKey = target.getAttribute('data-sort');
                if (sortKey) {
                    this.executeStaffDataSort(sortKey);
                }
            });
        });
    }
    
    // 担当別データの検索を実行
    private executeStaffDataSearch(): void {
        const staffNameSearch = document.getElementById('staffNameSearch') as HTMLInputElement;
        const regionSearch = document.getElementById('regionSearch') as HTMLInputElement;
        const departmentSearch = document.getElementById('departmentSearch') as HTMLInputElement;
        const ordersSearch = document.getElementById('ordersSearch') as HTMLSelectElement;
        const elderlySearch = document.getElementById('elderlySearch') as HTMLSelectElement;
        
        const staffNameValue = staffNameSearch?.value.toLowerCase() || '';
        const regionValue = regionSearch?.value.toLowerCase() || '';
        const departmentValue = departmentSearch?.value.toLowerCase() || '';
        const ordersValue = ordersSearch?.value || '';
        const elderlyValue = elderlySearch?.value || '';
        
        const table = document.getElementById('staffDataTable');
        if (!table) return;
        
        const rows = table.querySelectorAll('tbody tr');
        let visibleCount = 0;
        
        rows.forEach((row: Element) => {
            const rowElement = row as HTMLElement;
            const region = rowElement.getAttribute('data-region') || '';
            const department = rowElement.getAttribute('data-department') || '';
            const staff = rowElement.getAttribute('data-staff') || '';
            const orders = parseInt(rowElement.getAttribute('data-orders') || '0');
            const elderly = parseInt(rowElement.getAttribute('data-elderly') || '0');
            
            // 検索条件の判定
            const matchesStaff = !staffNameValue || staff.toLowerCase().includes(staffNameValue);
            const matchesRegion = !regionValue || region.toLowerCase().includes(regionValue);
            const matchesDepartment = !departmentValue || department.toLowerCase().includes(departmentValue);
            const matchesOrders = !ordersValue || this.matchesOrderCondition(orders, ordersValue);
            const matchesElderly = !elderlyValue || this.matchesOrderCondition(elderly, elderlyValue);
            
            if (matchesStaff && matchesRegion && matchesDepartment && matchesOrders && matchesElderly) {
                rowElement.style.display = '';
                visibleCount++;
                
                // 検索結果のハイライト
                if (staffNameValue || regionValue || departmentValue || ordersValue || elderlyValue) {
                    rowElement.classList.add('search-result-highlight');
                } else {
                    rowElement.classList.remove('search-result-highlight');
                }
            } else {
                rowElement.style.display = 'none';
                rowElement.classList.remove('search-result-highlight');
            }
        });
        
        // 検索結果件数を更新
        this.updateSearchResultInfo(visibleCount, rows.length);
    }
    
    // 検索条件の判定（件数）
    private matchesOrderCondition(actual: number, condition: string): boolean {
        if (!condition) return true;
        
        switch (condition) {
            case '0': return actual === 0;
            case '5': return actual >= 5;
            case '10': return actual >= 10;
            case '15': return actual >= 15;
            case '20': return actual >= 20;
            case '30': return actual >= 30;
            default: return true;
        }
    }
    
    // 検索結果件数を更新
    private updateSearchResultInfo(visibleCount: number, totalCount: number): void {
        const searchResultInfo = document.getElementById('searchResultInfo');
        if (searchResultInfo) {
            searchResultInfo.textContent = `検索結果: ${visibleCount}件 / 総件数: ${totalCount}件`;
        }
    }
    
    // 担当別データの検索条件をクリア
    private clearStaffDataSearch(): void {
        const staffNameSearch = document.getElementById('staffNameSearch') as HTMLInputElement;
        const regionSearch = document.getElementById('regionSearch') as HTMLInputElement;
        const departmentSearch = document.getElementById('departmentSearch') as HTMLInputElement;
        const ordersSearch = document.getElementById('ordersSearch') as HTMLSelectElement;
        const elderlySearch = document.getElementById('elderlySearch') as HTMLSelectElement;
        
        if (staffNameSearch) staffNameSearch.value = '';
        if (regionSearch) regionSearch.value = '';
        if (departmentSearch) departmentSearch.value = '';
        if (ordersSearch) ordersSearch.value = '';
        if (elderlySearch) elderlySearch.value = '';
        
        // 検索を実行して全件表示
        this.executeStaffDataSearch();
    }
    
    // 担当別データのソートを実行
    private executeStaffDataSort(sortKey: string): void {
        const table = document.getElementById('staffDataTable');
        if (!table) return;
        
        const tbody = table.querySelector('tbody');
        if (!tbody) return;
        
        const rows = Array.from(tbody.querySelectorAll('tr'));
        const header = table.querySelector(`th[data-sort="${sortKey}"]`);
        
        if (!header) return;
        
        // 現在のソート状態を確認
        const currentSort = header.classList.contains('sort-asc') ? 'asc' : 
                           header.classList.contains('sort-desc') ? 'desc' : 'none';
        
        // ソート状態をリセット
        table.querySelectorAll('th.sortable').forEach(th => {
            th.classList.remove('sort-asc', 'sort-desc');
        });
        
        // 新しいソート状態を設定
        let newSort: string;
        if (currentSort === 'none' || currentSort === 'desc') {
            newSort = 'asc';
            header.classList.add('sort-asc');
        } else {
            newSort = 'desc';
            header.classList.add('sort-desc');
        }
        
        // データをソート
        rows.sort((a, b) => {
            const aValue = this.getSortValue(a, sortKey);
            const bValue = this.getSortValue(b, sortKey);
            
            if (newSort === 'asc') {
                return this.compareValues(aValue, bValue);
            } else {
                return this.compareValues(bValue, aValue);
            }
        });
        
        // ソートされた行を再配置
        rows.forEach(row => tbody.appendChild(row));
    }
    
    // ソート用の値を取得
    private getSortValue(row: Element, sortKey: string): any {
        switch (sortKey) {
            case 'region':
            case 'department':
                const numValue = parseInt(row.getAttribute(`data-${sortKey}`) || '0');
                return isNaN(numValue) ? 0 : numValue;
            case 'staff':
                return row.getAttribute('data-staff') || '';
            case 'orders':
            case 'elderly':
            case 'single':
            case 'excessive':
            case 'overtime':
                const value = parseInt(row.getAttribute(`data-${sortKey}`) || '0');
                return isNaN(value) ? 0 : value;
            case 'normalAge':
                // 69歳以下の件数は data-normal-age 属性から取得
                const normalAgeValue = parseInt(row.getAttribute('data-normal-age') || '0');
                return isNaN(normalAgeValue) ? 0 : normalAgeValue;
            default:
                return '';
        }
    }
    
    // 値の比較
    private compareValues(a: any, b: any): number {
        if (typeof a === 'number' && typeof b === 'number') {
            return a - b;
        }
        if (typeof a === 'string' && typeof b === 'string') {
            return a.localeCompare(b, 'ja');
        }
        return 0;
    }
    
    // 担当別データのCSV出力ボタンのイベントリスナーを設定
    private setupStaffDataExportButton(): void {
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
                    
                    // 月報の場合はその月の代表的な日付（15日）、そうでなければ現在の日付を使用
                    let targetDate: Date;
                    if (monthlyData && monthlyData.length > 0) {
                        const reportMonthInput = document.getElementById('reportMonth') as HTMLInputElement;
                        if (reportMonthInput && reportMonthInput.value) {
                            const [year, month] = reportMonthInput.value.split('-').map(Number);
                            // 月の代表的な日付として15日を使用（月の真ん中）
                            targetDate = new Date(year, month - 1, 15);
                        } else {
                            targetDate = new Date();
                        }
                    } else {
                        targetDate = new Date();
                    }
                    
                    const staffData = this.reportGenerator.generateStaffData(targetData, targetDate);
                    await this.reportGenerator.exportStaffDataToCSV(staffData);
                    
                    this.showMessage('担当別データのCSV出力が完了しました。', 'success');
                } catch (error) {
                    console.error('担当別データCSV出力エラー:', error);
                    this.showMessage('CSVの出力に失敗しました。', 'error');
                }
            });
        }
    }

    // 現在の月報データを取得
    private getCurrentMonthlyData(): any[] | null {
        const reportMonthInput = document.getElementById('reportMonth') as HTMLInputElement;
        if (reportMonthInput && reportMonthInput.value) {
            const data = this.dataManager.getData();
            if (data.length > 0) {
                try {
                    const monthString = reportMonthInput.value;
                    const report = this.reportGenerator.generateMonthlyReport(data, monthString);
                    return report.rawData || null;
                } catch (error) {
                    console.error('月報データ取得エラー:', error);
                    return null;
                }
            }
        }
        return null;
    }

    // 月報を自動生成して担当別データと連動
    private generateMonthlyReportForFileDrop(): void {
        const data = this.dataManager.getData();
        if (data.length === 0) {
            console.log('ファイルドロップでデータが読み込まれていません。月報は生成されません。');
            return;
        }

        const reportMonthInput = document.getElementById('reportMonth') as HTMLInputElement;
        if (!reportMonthInput || !reportMonthInput.value) {
            console.log('ファイルドロップで月報作成月が選択されていません。月報は生成されません。');
            return;
        }

        try {
            const monthString = reportMonthInput.value; // 既に "YYYY-MM" 形式
            const report = this.reportGenerator.generateMonthlyReport(data, monthString);
            
            // 月報を表示
            this.displayReport(report, 'monthly');
            
            // 担当別データも月報データで更新
            this.updateStaffDataWithMonthlyData(report.rawData, monthString);
            
        } catch (error) {
            console.error('ファイルドロップで月報生成エラー:', error);
            this.showMessage('ファイルドロップで月報の生成中にエラーが発生しました。', 'error');
        }
    }

    // データフィルター機能を設定
    private setupDataFilters(): void {
        const staffFilter = document.getElementById('staffFilter') as HTMLInputElement;
        const regionFilter = document.getElementById('regionFilter') as HTMLInputElement;
        const departmentFilter = document.getElementById('departmentFilter') as HTMLInputElement;
        const table = document.getElementById('dataConfirmationTable');
        
        if (!table) return;
        
        const filterData = () => {
            const staffValue = staffFilter?.value.toLowerCase() || '';
            const regionValue = regionFilter?.value.toLowerCase() || '';
            const departmentValue = departmentFilter?.value.toLowerCase() || '';
            
            const rows = table.querySelectorAll('tbody tr');
            let visibleCount = 0;
            
            rows.forEach((row: Element) => {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 5) {
                    const staffName = cells[2]?.textContent?.toLowerCase() || '';
                    const regionNo = cells[3]?.textContent?.toLowerCase() || '';
                    const departmentNo = cells[4]?.textContent?.toLowerCase() || '';
                    
                    const matchesStaff = !staffValue || staffName.includes(staffValue);
                    const matchesRegion = !regionValue || regionNo.includes(regionValue);
                    const matchesDepartment = !departmentValue || departmentNo.includes(departmentValue);
                    
                    if (matchesStaff && matchesRegion && matchesDepartment) {
                        (row as HTMLElement).style.display = '';
                        visibleCount++;
                    } else {
                        (row as HTMLElement).style.display = 'none';
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
        if (staffFilter) staffFilter.addEventListener('input', filterData);
        if (regionFilter) regionFilter.addEventListener('input', filterData);
        if (departmentFilter) departmentFilter.addEventListener('input', filterData);
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
    private setDefaultDate(): void {
        const dateInput = document.getElementById('dateInput') as HTMLInputElement;
        const reportMonthInput = document.getElementById('reportMonth') as HTMLInputElement;
        
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
    private generateCalendar(containerId: string, type: 'public' | 'prohibited'): void {
        const container = document.getElementById(containerId);
        if (!container) return;

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
                selectedDates: new Set<string>()
            };
        }

        const state = this.calendarStates[containerId];
        const month = state.currentMonth;
        const year = state.currentYear;

        // カレンダーHTMLを生成
        const calendarHTML = this.createCalendarHTML(year, month, type, state.selectedDates);
        container.innerHTML = calendarHTML;

        // ヘッダーの月表示を更新
        this.updateMonthDisplay(type, year, month);

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
                const target = e.target as HTMLElement;
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
    private createCalendarHTML(year: number, month: number, type: 'public' | 'prohibited', selectedDates: Set<string>): string {
        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);
        const startDate = new Date(firstDay);
        startDate.setDate(startDate.getDate() - firstDay.getDay());

        const settings = this.calendarManager.getSettings();
        const holidayDates = type === 'public' ? settings.publicHolidays : settings.prohibitedDays;

        // デバッグ用ログ
        console.log(`${type}カレンダー生成 - 対象年月: ${year}年${month + 1}月`);
        console.log(`設定済み${type === 'public' ? '公休日' : '禁止日'}:`, holidayDates.map(d => d.toISOString().split('T')[0]));

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
                // 日付文字列を正しく生成（タイムゾーンの影響を排除）
                const dateString = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}-${String(currentDate.getDate()).padStart(2, '0')}`;
                
                const isCurrentMonth = currentDate.getMonth() === month;
                const isToday = this.isSameDate(currentDate, today);
                const isSelected = selectedDates.has(dateString);
                
                // 設定済み日付との比較（タイムゾーンの影響を排除）
                const isHoliday = holidayDates.some(h => {
                    const holidayString = `${h.getFullYear()}-${String(h.getMonth() + 1).padStart(2, '0')}-${String(h.getDate()).padStart(2, '0')}`;
                    return holidayString === dateString;
                });

                // デバッグ用ログ（特定の日付の場合）
                if (isHoliday) {
                    console.log(`${dateString} は${type === 'public' ? '公休日' : '禁止日'}として設定済み`);
                }

                let classes = 'day';
                if (!isCurrentMonth) classes += ' other-month';
                if (isToday) classes += ' today';
                if (isSelected) classes += ' selected';
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
    private createAddButtonHTML(type: 'public' | 'prohibited', selectedDates: Set<string>): string {
        if (selectedDates.size === 0) return '';

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
    private toggleDateSelection(containerId: string, dateStr: string, type: 'public' | 'prohibited'): void {
        const state = this.calendarStates[containerId];
        if (!state) return;

        // 日付文字列を正しく処理（タイムゾーンの影響を排除）
        const [year, month, day] = dateStr.split('-').map(Number);
        const date = new Date(year, month - 1, day); // monthは0ベースなので-1
        
        // 日付文字列を正しく生成
        const normalizedDateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

        if (state.selectedDates.has(normalizedDateStr)) {
            state.selectedDates.delete(normalizedDateStr);
        } else {
            state.selectedDates.add(normalizedDateStr);
        }

        // カレンダーを再生成
        this.generateCalendar(containerId, type);
        
        // デバッグ用ログ
        console.log(`${type}カレンダー - 日付選択切り替え: ${normalizedDateStr}, 選択状態: ${state.selectedDates.has(normalizedDateStr)}`);
    }

    // 選択された日付を一括追加
    private addSelectedDates(containerId: string, type: 'public' | 'prohibited'): void {
        const settings = this.calendarManager.getSettings();
        
        const state = this.calendarStates[containerId];
        if (!state) return;

        // 追加する日付を記録
        const addedDates: string[] = [];
        
        state.selectedDates.forEach(dateStr => {
            // 日付文字列を正しく解析（タイムゾーンの影響を排除）
            const [year, month, day] = dateStr.split('-').map(Number);
            const date = new Date(year, month - 1, day); // monthは0ベースなので-1
            
            if (type === 'public') {
                if (!settings.publicHolidays.some(d => this.isSameDate(d, date))) {
                    settings.publicHolidays.push(date);
                    addedDates.push(dateStr);
                }
            } else {
                if (!settings.prohibitedDays.some(d => this.isSameDate(d, date))) {
                    settings.prohibitedDays.push(date);
                    addedDates.push(dateStr);
                }
            }
        });

        // CalendarManagerに設定を保存
        this.calendarManager.updateSettings(settings);
        
        // DataManagerにも設定を同期
        this.dataManager.setHolidaySettings(settings);
        
        // ReportGeneratorにも設定を反映
        this.reportGenerator.updateHolidaySettings(settings);
        
        // 設定済み日付一覧を更新
        this.loadHolidaySettings();
        
        // 選択状態をクリア
        if (this.calendarStates) {
            if (this.calendarStates[containerId]) {
                this.calendarStates[containerId].selectedDates.clear();
            }
        }
        
        // カレンダーを再生成（追加された日付を表示）
        this.generateCalendar(containerId, type);
        
        // 成功メッセージを表示
        const message = type === 'public' ? '公休日' : '禁止日';
        if (addedDates.length > 0) {
            this.showMessage(`${message}が${addedDates.length}日正常に追加されました。`, 'success');
            console.log(`追加された日付: ${addedDates.join(', ')}`);
            } else {
            this.showMessage('追加する日付がありませんでした。', 'info');
        }
    }

    // 日付比較
    private isSameDate(date1: Date, date2: Date): boolean {
        return date1.getFullYear() === date2.getFullYear() &&
               date1.getMonth() === date2.getMonth() &&
               date1.getDate() === date2.getDate();
    }
}

// アプリケーションの初期化
document.addEventListener('DOMContentLoaded', () => {
    new App();
});