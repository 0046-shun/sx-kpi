import { HolidaySettings } from './types.js';
export declare class CalendarManager {
    private holidaySettings;
    constructor();
    updateSettings(settings: HolidaySettings): void;
    getSettings(): HolidaySettings;
    isPublicHoliday(date: Date): boolean;
    isProhibitedDay(date: Date): boolean;
    private isSameDate;
    private saveSettings;
    private loadSettings;
    addPublicHoliday(date: Date): void;
    addProhibitedDay(date: Date): void;
    removePublicHoliday(date: Date): void;
    removeProhibitedDay(date: Date): void;
    clearSettings(): void;
}
