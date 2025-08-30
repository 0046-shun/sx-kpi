import { HolidaySettings, StaffData } from './types.js';

export class ExcelProcessor {
    
    async readExcelFile(file: File): Promise<any[]> {
        
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target?.result as ArrayBuffer);
                    const workbook = (window as any).XLSX.read(data, { type: 'array' });
                    
                    // 最初のシートを取得
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    
                    // シートをJSONに変換（ヘッダー行を保持）
                    const jsonData = (window as any).XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // データの検証とクリーニング
                    const cleanedData = this.cleanAndValidateData(jsonData);
                    
                    resolve(cleanedData);
                    
                } catch (error) {
                    console.error('Excelファイル読み込みエラー:', error);
                    reject(new Error(`Excelファイルの読み込みに失敗しました: ${error}`));
                }
            };
            
            reader.onerror = () => {
                reject(new Error('ファイルの読み込みに失敗しました'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }
    
    private cleanAndValidateData(rawData: any[]): any[] {
        
        if (rawData.length < 10) {
            throw new Error('データが不足しています。最低10行のデータが必要です。');
        }
        
        // ヘッダー行（1-8行目）と項目行（9行目）をスキップ
        const allDataRows = rawData.slice(9);
        
        // 空行を除外（データがある行のみを抽出）
        const dataRows = allDataRows.filter((row, index) => {
            const hasData = row && Array.isArray(row) && row.length > 0 && row[0] !== undefined && row[0] !== null && row[0] !== '';
            return hasData;
        });
        
        // データ行をオブジェクトに変換
        const processedData = dataRows.map((row, index) => {
            const rowNumber = index + 10; // 実際の行番号
            const parsedDate = this.parseDate(row[0]); // A列：日付
            
            return {
                rowNumber,
                date: parsedDate,
                time: this.parseTime(row[1]), // B列：時間
                regionNumber: row[2], // C列：地区№
                departmentNumber: row[3], // D列：所属№
                staffName: this.normalizeStaffName(row[4]), // E列：担当者名（正規化）
                contractor: row[5], // F列：契約者
                contractorAge: this.parseAge(row[6]), // G列：年齢
                contractorRelation: row[7], // H列：相手
                contractorTel: row[8], // I列：契約者TEL
                confirmation: row[9], // J列：確認
                confirmationDateTime: row[10], // K列：確認者日時
                confirmer: row[11], // L列：確認者
                confirmerAge: this.parseAge(row[12]), // M列：年齢
                confirmerRelation: row[13], // N列：相手
                confirmerTel: row[14], // O列：確認者TEL
                productName: row[15], // P列：商品名
                quantity: row[16], // Q列：数量
                amount: row[17], // R列：金額
                contractDate: this.parseDate(row[18]), // S列：契約予定日
                startDate: this.parseDate(row[19]), // T列：着工日
                startTime: row[20], // U列：着工時間
                completionDate: this.parseDate(row[21]), // V列：完工予定日
                completionTime: row[22], // W列：完工予定時間
                notes: row[23], // X列：備考
                // 追加の列があればここに追加
                debugA: row[0], // A列の日付
                debugB: row[1]  // B列の時間
            };
        }); // 全データを保持（フィルタリングはレポート生成時に実行）
        
        return processedData;
    }
    
    private parseDateFromKColumn(kColumnValue: any): Date | null {
        if (!kColumnValue || typeof kColumnValue !== 'string') {
            return null;
        }
        
        const kColumnStr = String(kColumnValue);
        
        // K列から日付を抽出（例：'8/19 19:18 本室' → 8/19）
        const dateMatch = kColumnStr.match(/(\d{1,2})\/(\d{1,2})/);
        if (dateMatch) {
            const month = parseInt(dateMatch[1]);
            const day = parseInt(dateMatch[2]);
            const year = 2025; // 固定年（必要に応じて調整）
            const result = new Date(year, month - 1, day); // JavaScriptの月は0ベース
            return result;
        }
        
        return null;
    }
    
    private parseDate(value: any): Date | null {
        if (!value) {
            return null;
        }
        
        // Excelの日付形式（数値）を処理
        if (typeof value === 'number') {
            // Excelの日付は1900年1月1日からの日数（修正版）
            // 45889 = 2025年8月20日になるように調整
            const excelEpoch = new Date(1900, 0, 1);
            const date = new Date(excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000);
            return date;
        }
        
        // 文字列形式の日付を処理
        if (typeof value === 'string') {
            const trimmedValue = value.trim();
            
            // "YYYY年MM月DD日" 形式
            const japaneseMatch = trimmedValue.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
            if (japaneseMatch) {
                const [_, year, month, day] = japaneseMatch;
                const result = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                return result;
            }
            
            // "YYYY/MM/DD" 形式
            const slashMatch = trimmedValue.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
            if (slashMatch) {
                const [_, year, month, day] = slashMatch;
                const result = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                return result;
            }
            
            // "YYYY-MM-DD" 形式
            const dashMatch = trimmedValue.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
            if (dashMatch) {
                const [_, year, month, day] = dashMatch;
                const result = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                return result;
            }
            
            // "MM/DD" 形式（年は現在年を使用）
            const shortMatch = trimmedValue.match(/(\d{1,2})\/(\d{1,2})/);
            if (shortMatch) {
                const [_, month, day] = shortMatch;
                const currentYear = new Date().getFullYear();
                const result = new Date(currentYear, parseInt(month) - 1, parseInt(day));
                return result;
            }
            
            // 直接変換を試行
            try {
                const result = new Date(trimmedValue);
                if (!isNaN(result.getTime())) {
                    return result;
                }
            } catch (error) {
                // 変換失敗は無視
            }
        }
        
        return null;
    }
    
    private parseTime(value: any): string | null {
        if (!value) {
            return null;
        }
        
        // Excelの時間形式（数値）を処理
        if (typeof value === 'number') {
            // Excelの時間は0.0〜1.0の小数（例：0.5 = 12:00）
            const totalMinutes = Math.round(value * 24 * 60);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        }
        
        // 文字列形式の時間を処理
        if (typeof value === 'string') {
            const trimmedValue = value.trim();
            
            // "HH:MM" 形式
            const timeMatch = trimmedValue.match(/(\d{1,2}):(\d{1,2})/);
            if (timeMatch) {
                const [_, hours, minutes] = timeMatch;
                return `${hours.padStart(2, '0')}:${minutes.padStart(2, '0')}`;
            }
            
            // "HH：MM" 形式（全角コロン）
            const fullTimeMatch = trimmedValue.match(/(\d{1,2})：(\d{1,2})/);
            if (fullTimeMatch) {
                const [_, hours, minutes] = fullTimeMatch;
                return `${hours.padStart(2, '0')}:${minutes.padStart(2, '0')}`;
            }
        }
        
        return null;
    }
    
    private parseAge(value: any): number | null {
        if (value === null || value === undefined || value === '') {
            return null;
        }
        
        // すでに数値
        if (typeof value === 'number') {
            const n = Math.floor(value);
            return n > 0 && n < 150 ? n : null;
        }
        
        // 文字列から数字を抽出（例: '84歳', ' 84 ', '84.0', '84　歳'）
        if (typeof value === 'string') {
            const trimmedValue = value.trim();
            
            // 数字のみの文字列
            if (/^\d+$/.test(trimmedValue)) {
                const n = parseInt(trimmedValue, 10);
                return n > 0 && n < 150 ? n : null;
            }
            
            // 数字を含む文字列（例: '84歳', '84　歳'）
            const matched = trimmedValue.match(/(\d+)/);
            if (matched) {
                const n = parseInt(matched[1], 10);
                return n > 0 && n < 150 ? n : null;
            }
        }
        
        // それ以外は数値変換を試みる
        const numValue = Number(value);
        if (!isNaN(numValue) && numValue > 0 && numValue < 150) {
            return Math.floor(numValue);
        }
        
        return null;
    }
    
    /**
     * 担当者名を正規化する
     * 括弧内の情報（役職、担当範囲、技術分野など）を除いて基本の担当者名を抽出
     * 例：
     * - "山田(SE)" → "山田"
     * - "山田(岡田)" → "山田"
     * - "山田(技)" → "山田"
     * - "山田（技術）" → "山田"
     */
    private normalizeStaffName(staffName: any): string {
        if (!staffName || typeof staffName !== 'string') {
            return '';
        }
        
        const trimmedName = staffName.trim();
        if (trimmedName === '') {
            return '';
        }
        
        // 括弧（半角・全角）で囲まれた部分を除去
        // 半角括弧: ( )
        // 全角括弧: （ ）
        const normalizedName = trimmedName.replace(/[\(（].*?[\)）]/g, '');
        
        // 前後の空白を除去して返す
        return normalizedName.trim();
    }
    
    /**
     * 指定された日付の受注かどうかを判定
     * @param row データ行
     * @param targetDate 対象日付
     * @param isDailyReport 日報生成時かどうか（true: 日報、false: 月報）
     */
    public isOrderForDate(row: any, targetDate: Date, isDailyReport: boolean = true): boolean {
        // 基本チェック
        if (!row.date || !targetDate) {
            return false;
        }

        // 受注日（A列）と確認日時（K列）の日付を比較して、遅い日付を取得
        let effectiveDate = row.date; // デフォルトは受注日

        if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
            const confirmationStr = row.confirmationDateTime;

            // K列から日付を抽出（8/30 10:03 大城 のような形式）
            const dateTimePattern = confirmationStr.match(/(\d{1,2})\/(\d{1,2})/);
            if (dateTimePattern) {
                const kColumnMonth = parseInt(dateTimePattern[1]);
                const kColumnDay = parseInt(dateTimePattern[2]);
                const kColumnDate = new Date(targetDate.getFullYear(), kColumnMonth - 1, kColumnDay);

                // 受注日と確認日を比較して、遅い日付を採用
                if (kColumnDate > row.date) {
                    effectiveDate = kColumnDate;
                }
            }
        }

        // 日付単位での判定（日報の場合）
        const effectiveDateOnly = new Date(effectiveDate.getFullYear(), effectiveDate.getMonth(), effectiveDate.getDate());
        const targetDateOnly = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
        const dateMatch = effectiveDateOnly.getTime() === targetDateOnly.getTime();

        if (dateMatch) {
            // J列条件チェック（1、2、5の場合は除外）
            const confirmation = row.confirmation;
            if (typeof confirmation === 'number') {
                if (confirmation === 1 || confirmation === 2 || confirmation === 5) {
                    return false;
                }
            } else if (typeof confirmation === 'string') {
                const trimmedValue = confirmation.trim();
                if (trimmedValue !== '') {
                    const confirmationNum = parseInt(trimmedValue);
                    if (!isNaN(confirmationNum) && (confirmationNum === 1 || confirmationNum === 2 || confirmationNum === 5)) {
                        return false;
                    }
                }
            }

            // K列条件チェック（受注パターン）
            // K列に日付がない場合は、A列の日付で受注カウント
            if (!row.confirmationDateTime || typeof row.confirmationDateTime !== 'string') {
                return true;
            }

            const confirmationStr = row.confirmationDateTime;

            // 除外キーワードのチェック
            if (confirmationStr.includes('担当待ち') || confirmationStr.includes('直電') ||
                confirmationStr.includes('契約時') || confirmationStr.includes('待ち') ||
                confirmationStr.includes('入電予定')) {
                return false;
            }

            // 「同時」パターンのチェック
            if (confirmationStr.includes('同時')) {
                return true;
            }

            // 有効な受注パターンのチェック
            if (confirmationStr.includes('単独') || confirmationStr.includes('過量')) {
                return true;
            }

            // 69歳以下パターンのチェック
            if (row.contractorAge && typeof row.contractorAge === 'number' && row.contractorAge <= 69) {
                if (confirmationStr.includes('69歳以下')) {
                    return true;
                }
            }

            // 日付パターンのチェック（8/25 11:55 大城 のような形式）
            const dateTimePattern = confirmationStr.match(/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{1,2})/);
            if (dateTimePattern) {
                const kColumnMonth = parseInt(dateTimePattern[1]);
                const kColumnDay = parseInt(dateTimePattern[2]);
                const kColumnDateOnly = new Date(targetDate.getFullYear(), kColumnMonth - 1, kColumnDay);
                const kDateMatch = kColumnDateOnly.getTime() === targetDateOnly.getTime();

                if (kDateMatch) {
                    return true;
                }
            }

            // 単純な日付パターンのチェック（8/25 のような形式）
            const simpleDatePattern = confirmationStr.match(/(\d{1,2})\/(\d{1,2})/);
            if (simpleDatePattern) {
                const kColumnMonth = parseInt(simpleDatePattern[1]);
                const kColumnDay = parseInt(simpleDatePattern[2]);
                const kColumnDateOnly = new Date(targetDate.getFullYear(), kColumnMonth - 1, kColumnDay);
                const kDateMatch = kColumnDateOnly.getTime() === targetDateOnly.getTime();

                if (kDateMatch) {
                    return true;
                }
            }

            return false;
        }

        // 日付が一致しない場合の処理
        if (isDailyReport) {
            // 日報生成時は、日付が一致しない場合は受注として扱わない
            return false;
        }

        // 月報生成時のみ月報判定を実行
        // 月単位での判定（月報の場合）
        const monthMatch = effectiveDate.getFullYear() === targetDate.getFullYear() && 
                          effectiveDate.getMonth() === targetDate.getMonth();
        
        if (monthMatch) {
            // 月が一致する場合は、基本的に受注として扱う
            // ただし、明らかに除外すべき条件のみチェック
            
            // J列条件チェック（1、2、5の場合は除外）
            const confirmation = row.confirmation;
            if (typeof confirmation === 'number') {
                if (confirmation === 1 || confirmation === 2 || confirmation === 5) {
                    return false;
                }
            } else if (typeof confirmation === 'string') {
                const trimmedValue = confirmation.trim();
                if (trimmedValue !== '') {
                    const confirmationNum = parseInt(trimmedValue);
                    if (!isNaN(confirmationNum) && (confirmationNum === 1 || confirmationNum === 2 || confirmationNum === 5)) {
                        return false;
                    }
                }
            }
            
            // K列の明らかな除外キーワードのみチェック
            if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
                const confirmationStr = row.confirmationDateTime;
                // より厳密な除外条件のみ
                if (confirmationStr.includes('担当待ち') || confirmationStr.includes('直電')) {
                    return false;
                }
            }
            
            return true;
        }

        return false;
    }
        
    /**
     * 単独契約かどうかを判定
     */
    public isSingle(row: any): boolean {
        if (!row.confirmationDateTime || typeof row.confirmationDateTime !== 'string') {
            return false;
        }
        
        const confirmationStr = String(row.confirmationDateTime);
        const result = confirmationStr.includes('単独');
        
        return result;
    }
    
    /**
     * 過量販売かどうかを判定
     */
    public isExcessive(row: any): boolean {
        if (!row.confirmationDateTime || typeof row.confirmationDateTime !== 'string') {
            return false;
        }
        
        const confirmationStr = String(row.confirmationDateTime);
        const result = confirmationStr.includes('過量');
        
        return result;
    }
    
    /**
     * 時間外対応かどうかを判定
     * 18:30以降の対応を時間外とする
     */
    public isOvertime(row: any): boolean {
        
        // 基本的なデータ存在チェック
        if (!row.date || !row.time) {
            return false;
        }
        
        // 受注条件を満たすかチェック
        if (!this.isOrderForDate(row, row.date, true)) {  // 日報判定として扱う
            return false;
        }
        
        // 時間外判定の優先順位
        
        // 優先度1: K列が「同時」の場合
        if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string' && row.confirmationDateTime.includes('同時')) {
            if (row.date && row.time) {
                const checkTime = new Date(row.date);
                const [hours, minutes] = row.time.split(':').map(Number);
                checkTime.setHours(hours, minutes, 0, 0);
                
                const overtimeThreshold = new Date(row.date);
                overtimeThreshold.setHours(18, 30, 0, 0);
                
                const result = checkTime >= overtimeThreshold;
                return result;
            }
        }
        
        // 優先度2: K列に時間情報が含まれる場合
        if (row.confirmationDateTime && typeof row.confirmationDateTime === 'string') {
            const timeMatch = row.confirmationDateTime.match(/(\d{1,2}):(\d{1,2})/);
                if (timeMatch) {
                const hours = parseInt(timeMatch[1]);
                const minutes = parseInt(timeMatch[2]);
                
                if (hours > 18 || (hours === 18 && minutes >= 30)) {
                    return true;
                }
            }
        }
        
        // 優先度3: 標準時間フィールドを使用
        if (row.time) {
            const [hours, minutes] = row.time.split(':').map(Number);
            if (hours > 18 || (hours === 18 && minutes >= 30)) {
                return true;
            }
        }
        
        return false;
    }
    
    /**
     * 現在の日付フィルターを取得（動的）
     */
    public getCurrentDateFilter(): Date {
        // UI要素から日付を取得
        const dateInput = document.getElementById('dateInput') as HTMLInputElement;
        if (dateInput && dateInput.value) {
            const inputDate = new Date(dateInput.value);
            if (!isNaN(inputDate.getTime())) {
                return inputDate;
            }
        }
        
        // UI要素が取得できない場合は現在日付を使用
        return new Date();
    }
}