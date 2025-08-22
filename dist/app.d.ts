export declare class App {
    private excelProcessor;
    private reportGenerator;
    private dataManager;
    private calendarManager;
    private currentData;
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
    private updateStaffData;
    private setupExportButtons;
    private showMessage;
    private setDefaultDate;
}
