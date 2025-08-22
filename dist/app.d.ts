export declare class App {
    private excelProcessor;
    private reportGenerator;
    private dataManager;
    private calendarManager;
    constructor();
    private initializeApp;
    private setupEventListeners;
    private initializeCalendarManager;
    private loadSavedHolidaySettings;
    private openHolidaySettings;
    private handleHolidaySettingsChanged;
    private refreshCurrentReport;
    private handleFileUpload;
    private generateDailyReport;
    private generateMonthlyReport;
    private displayReport;
    private setupExportButtons;
    private showMessage;
    private setDefaultDate;
}
