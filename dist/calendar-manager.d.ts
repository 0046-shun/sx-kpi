import { HolidaySettings } from './types.js';
export declare class CalendarManager {
    private currentYear;
    private currentMonth;
    private holidaySettings;
    constructor();
    initializeCalendars(): void;
    private renderPublicHolidayCalendar;
    private renderProhibitedDayCalendar;
    private generateCalendarHTML;
    private isDateSelected;
    private isToday;
    private formatMonthYear;
    private updateHolidayLists;
    private updatePublicHolidayList;
    private updateProhibitedDayList;
    private formatDate;
    private setupEventListeners;
    private toggleDateSelection;
    private removeDateSelection;
    private saveSettings;
    getHolidaySettings(): HolidaySettings;
    setHolidaySettings(settings: HolidaySettings): void;
}
