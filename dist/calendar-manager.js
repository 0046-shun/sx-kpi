export class CalendarManager {
    constructor() {
        this.holidaySettings = {
            publicHolidays: [],
            prohibitedDays: []
        };
        this.loadSettings();
    }
    // 設定を更新
    updateSettings(settings) {
        this.holidaySettings = settings;
        this.saveSettings();
    }
    // 設定を取得
    getSettings() {
        return this.holidaySettings;
    }
    // 公休日かどうかの判定
    isPublicHoliday(date) {
        return this.holidaySettings.publicHolidays.some(holiday => this.isSameDate(date, holiday));
    }
    // 禁止日かどうかの判定
    isProhibitedDay(date) {
        return this.holidaySettings.prohibitedDays.some(prohibited => this.isSameDate(date, prohibited));
    }
    // 日付が同じかどうかの判定
    isSameDate(date1, date2) {
        return date1.getFullYear() === date2.getFullYear() &&
            date1.getMonth() === date2.getMonth() &&
            date1.getDate() === date2.getDate();
    }
    // 設定をローカルストレージに保存
    saveSettings() {
        try {
            const settingsData = {
                publicHolidays: this.holidaySettings.publicHolidays.map(date => date.toISOString()),
                prohibitedDays: this.holidaySettings.prohibitedDays.map(date => date.toISOString())
            };
            localStorage.setItem('holidaySettings', JSON.stringify(settingsData));
        }
        catch (error) {
            console.error('設定の保存に失敗しました:', error);
        }
    }
    // 設定をローカルストレージから読み込み
    loadSettings() {
        try {
            const savedSettings = localStorage.getItem('holidaySettings');
            if (savedSettings) {
                const settingsData = JSON.parse(savedSettings);
                this.holidaySettings = {
                    publicHolidays: settingsData.publicHolidays.map((dateStr) => new Date(dateStr)),
                    prohibitedDays: settingsData.prohibitedDays.map((dateStr) => new Date(dateStr))
                };
            }
        }
        catch (error) {
            console.error('設定の読み込みに失敗しました:', error);
        }
    }
    // 公休日を追加
    addPublicHoliday(date) {
        if (!this.isPublicHoliday(date)) {
            this.holidaySettings.publicHolidays.push(date);
            this.saveSettings();
        }
    }
    // 禁止日を追加
    addProhibitedDay(date) {
        if (!this.isProhibitedDay(date)) {
            this.holidaySettings.prohibitedDays.push(date);
            this.saveSettings();
        }
    }
    // 公休日を削除
    removePublicHoliday(date) {
        this.holidaySettings.publicHolidays = this.holidaySettings.publicHolidays.filter(holiday => !this.isSameDate(holiday, date));
        this.saveSettings();
    }
    // 禁止日を削除
    removeProhibitedDay(date) {
        this.holidaySettings.prohibitedDays = this.holidaySettings.prohibitedDays.filter(prohibited => !this.isSameDate(prohibited, date));
        this.saveSettings();
    }
    // 設定をクリア
    clearSettings() {
        this.holidaySettings = {
            publicHolidays: [],
            prohibitedDays: []
        };
        this.saveSettings();
    }
}
//# sourceMappingURL=calendar-manager.js.map