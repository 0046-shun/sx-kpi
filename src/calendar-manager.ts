import { HolidaySettings } from './types.js';

export class CalendarManager {
    private currentYear: number;
    private currentMonth: number;
    private holidaySettings: HolidaySettings;
    
    constructor() {
        const now = new Date();
        this.currentYear = now.getFullYear();
        this.currentMonth = now.getMonth();
        this.holidaySettings = {
            publicHolidays: [],
            prohibitedDays: []
        };
    }
    
    // カレンダーの初期化
    initializeCalendars(): void {
        this.renderPublicHolidayCalendar();
        this.renderProhibitedDayCalendar();
        this.updateHolidayLists();
        this.setupEventListeners();
    }
    
    // 公休日カレンダーの描画
    private renderPublicHolidayCalendar(): void {
        const calendar = document.getElementById('publicHolidayCalendar');
        const monthDisplay = document.getElementById('currentMonthPublic');
        
        if (!calendar || !monthDisplay) return;
        
        monthDisplay.textContent = this.formatMonthYear(this.currentYear, this.currentMonth);
        calendar.innerHTML = this.generateCalendarHTML(this.currentYear, this.currentMonth, 'public');
    }
    
    // 禁止日カレンダーの描画
    private renderProhibitedDayCalendar(): void {
        const calendar = document.getElementById('prohibitedDayCalendar');
        const monthDisplay = document.getElementById('currentMonthProhibited');
        
        if (!calendar || !monthDisplay) return;
        
        monthDisplay.textContent = this.formatMonthYear(this.currentYear, this.currentMonth);
        calendar.innerHTML = this.generateCalendarHTML(this.currentYear, this.currentMonth, 'prohibited');
    }
    
    // カレンダーHTMLの生成
    private generateCalendarHTML(year: number, month: number, type: 'public' | 'prohibited'): string {
        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);
        
        // カレンダーの開始日を正しく計算（前月の日曜日から開始）
        const startDate = new Date(firstDay);
        const firstDayOfWeek = firstDay.getDay(); // 0=日曜日, 1=月曜日, ..., 6=土曜日
        startDate.setDate(startDate.getDate() - firstDayOfWeek);
        
        let html = '<div class="calendar-weekdays">';
        ['日', '月', '火', '水', '木', '金', '土'].forEach(day => {
            html += `<div class="calendar-weekday">${day}</div>`;
        });
        html += '</div>';
        
        html += '<div class="calendar-days">';
        
        // 6週間分（42日）のカレンダーを生成
        for (let i = 0; i < 42; i++) {
            const currentDate = new Date(startDate);
            currentDate.setDate(startDate.getDate() + i);
            
            const isCurrentMonth = currentDate.getMonth() === month;
            const isToday = this.isToday(currentDate);
            const isSelected = this.isDateSelected(currentDate, type);
            const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;
            
            let classes = 'calendar-day';
            if (!isCurrentMonth) classes += ' other-month';
            if (isToday) classes += ' today';
            if (isSelected) classes += type === 'public' ? ' selected-public' : ' selected-prohibited';
            if (isWeekend) classes += ' weekend';
            
            html += `
                <div class="${classes}" data-date="${currentDate.toISOString()}" data-type="${type}">
                    ${currentDate.getDate()}
                </div>
            `;
        }
        
        html += '</div>';
        return html;
    }
    
    // 日付が選択されているかチェック
    private isDateSelected(date: Date, type: 'public' | 'prohibited'): boolean {
        const dateStr = date.toISOString().split('T')[0];
        if (type === 'public') {
            return this.holidaySettings.publicHolidays.some(d => d.toISOString().split('T')[0] === dateStr);
        } else {
            return this.holidaySettings.prohibitedDays.some(d => d.toISOString().split('T')[0] === dateStr);
        }
    }
    
    // 今日の日付かチェック
    private isToday(date: Date): boolean {
        const today = new Date();
        return date.toDateString() === today.toDateString();
    }
    
    // 月年のフォーマット
    private formatMonthYear(year: number, month: number): string {
        return `${year}年${month + 1}月`;
    }
    
    // 公休日・禁止日リストの更新
    private updateHolidayLists(): void {
        this.updatePublicHolidayList();
        this.updateProhibitedDayList();
    }
    
    // 公休日リストの更新
    private updatePublicHolidayList(): void {
        const list = document.getElementById('publicHolidayList');
        if (!list) return;
        
        list.innerHTML = this.holidaySettings.publicHolidays.length === 0 
            ? '<p class="text-muted">設定された公休日はありません</p>'
            : this.holidaySettings.publicHolidays
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
    
    // 禁止日リストの更新
    private updateProhibitedDayList(): void {
        const list = document.getElementById('prohibitedDayList');
        if (!list) return;
        
        list.innerHTML = this.holidaySettings.prohibitedDays.length === 0 
            ? '<p class="text-muted">設定された禁止日はありません</p>'
            : this.holidaySettings.prohibitedDays
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
    
    // 日付のフォーマット
    private formatDate(date: Date): string {
        return date.toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
    }
    
    // イベントリスナーの設定
    private setupEventListeners(): void {
        // 月切り替えボタン
        document.getElementById('prevMonthPublic')?.addEventListener('click', () => {
            this.currentMonth--;
            if (this.currentMonth < 0) {
                this.currentMonth = 11;
                this.currentYear--;
            }
            this.renderPublicHolidayCalendar();
        });
        
        document.getElementById('nextMonthPublic')?.addEventListener('click', () => {
            this.currentMonth++;
            if (this.currentMonth > 11) {
                this.currentMonth = 0;
                this.currentYear++;
            }
            this.renderPublicHolidayCalendar();
        });
        
        document.getElementById('prevMonthProhibited')?.addEventListener('click', () => {
            this.currentMonth--;
            if (this.currentMonth < 0) {
                this.currentMonth = 11;
                this.currentYear--;
            }
            this.renderProhibitedDayCalendar();
        });
        
        document.getElementById('nextMonthProhibited')?.addEventListener('click', () => {
            this.currentMonth++;
            if (this.currentMonth > 11) {
                this.currentMonth = 0;
                this.currentYear++;
            }
            this.renderProhibitedDayCalendar();
        });
        
        // カレンダー日付クリック
        document.addEventListener('click', (e) => {
            const target = e.target as HTMLElement;
            if (target.classList.contains('calendar-day')) {
                const date = new Date(target.dataset.date || '');
                const type = target.dataset.type as 'public' | 'prohibited';
                this.toggleDateSelection(date, type);
            }
        });
        
        // 削除ボタンクリック
        document.addEventListener('click', (e) => {
            const target = e.target as HTMLElement;
            if (target.classList.contains('holiday-remove')) {
                const date = new Date(target.dataset.date || '');
                const type = target.dataset.type as 'public' | 'prohibited';
                this.removeDateSelection(date, type);
            }
        });
        
        // 保存ボタンクリック
        document.getElementById('saveHolidaySettings')?.addEventListener('click', () => {
            this.saveSettings();
        });
    }
    
    // 日付選択の切り替え
    private toggleDateSelection(date: Date, type: 'public' | 'prohibited'): void {
        const dateStr = date.toISOString().split('T')[0];
        
        if (type === 'public') {
            const isSelected = this.holidaySettings.publicHolidays.some(d => d.toISOString().split('T')[0] === dateStr);
            if (isSelected) {
                this.holidaySettings.publicHolidays = this.holidaySettings.publicHolidays.filter(
                    d => d.toISOString().split('T')[0] !== dateStr
                );
            } else {
                this.holidaySettings.publicHolidays.push(date);
            }
        } else {
            const isSelected = this.holidaySettings.prohibitedDays.some(d => d.toISOString().split('T')[0] === dateStr);
            if (isSelected) {
                this.holidaySettings.prohibitedDays = this.holidaySettings.prohibitedDays.filter(
                    d => d.toISOString().split('T')[0] !== dateStr
                );
            } else {
                this.holidaySettings.prohibitedDays.push(date);
            }
        }
        
        this.renderPublicHolidayCalendar();
        this.renderProhibitedDayCalendar();
        this.updateHolidayLists();
    }
    
    // 日付選択の削除
    private removeDateSelection(date: Date, type: 'public' | 'prohibited'): void {
        const dateStr = date.toISOString().split('T')[0];
        
        if (type === 'public') {
            this.holidaySettings.publicHolidays = this.holidaySettings.publicHolidays.filter(
                d => d.toISOString().split('T')[0] !== dateStr
            );
        } else {
            this.holidaySettings.prohibitedDays = this.holidaySettings.prohibitedDays.filter(
                d => d.toISOString().split('T')[0] !== dateStr
            );
        }
        
        this.renderPublicHolidayCalendar();
        this.renderProhibitedDayCalendar();
        this.updateHolidayLists();
    }
    
    // 設定の保存
    private saveSettings(): void {
        // カスタムイベントで設定変更を通知
        const event = new CustomEvent('holidaySettingsChanged', {
            detail: this.holidaySettings
        });
        document.dispatchEvent(event);
        
        // モーダルを閉じる
        const modal = document.getElementById('holidaySettingsModal');
        if (modal) {
            const bootstrapModal = (window as any).bootstrap?.Modal.getInstance(modal);
            if (bootstrapModal) {
                bootstrapModal.hide();
            }
        }
    }
    
    // 設定の取得
    getHolidaySettings(): HolidaySettings {
        return { ...this.holidaySettings };
    }
    
    // 設定の設定
    setHolidaySettings(settings: HolidaySettings): void {
        this.holidaySettings = settings;
        this.renderPublicHolidayCalendar();
        this.renderProhibitedDayCalendar();
        this.updateHolidayLists();
    }
}
