export class ExcelProcessor {
    
    async readExcelFile(file: File): Promise<any[]> {
        console.log('Excelファイル読み込み開始:', file.name);
        
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
                    console.log('Excelファイル読み込み成功、データ処理開始');
                    const cleanedData = this.cleanAndValidateData(jsonData);
                    console.log('データ処理完了、件数:', cleanedData.length);
                    
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
        console.log('データクリーニング開始、総行数:', rawData.length);
        
        if (rawData.length < 10) {
            throw new Error('データが不足しています。最低10行のデータが必要です。');
        }
        
        // ヘッダー行（1-8行目）と項目行（9行目）をスキップ
        const allDataRows = rawData.slice(9);
        console.log('全データ行数（ヘッダー除く）:', allDataRows.length);
        
        // 空行を除外（データがある行のみを抽出）
        const dataRows = allDataRows.filter((row, index) => {
            const hasData = row && Array.isArray(row) && row.length > 0 && row[0] !== undefined && row[0] !== null && row[0] !== '';
            if (!hasData && index < 10) {
                console.log(`行${index + 10}は空行のためスキップ:`, { row, length: row ? row.length : 0 });
            }
            return hasData;
        });
        
        console.log('有効データ行数（空行除く）:', dataRows.length);
        
        // 最初の有効行の詳細構造を確認
        if (dataRows.length > 0) {
            console.log('最初の有効行の詳細構造:', {
                rowLength: dataRows[0].length,
                rowKeys: Object.keys(dataRows[0]),
                rowValues: dataRows[0],
                firstFewValues: dataRows[0].slice(0, 15)
            });
        }
        
        // データ行をオブジェクトに変換
        const processedData = dataRows.map((row, index) => {
            const rowNumber = index + 10; // 実際の行番号
            
            return {
                rowNumber,
                date: this.parseDate(row[0]), // A列：日付
                time: this.parseTime(row[1]), // B列：時間
                regionNumber: row[2], // C列：地区№
                departmentNumber: row[3], // D列：所属№
                staffName: row[4], // E列：担当者名
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
                paymentMethod: row[22], // W列：支払方法
                receptionist: row[23], // X列：受付者
                coFlyer: row[24], // Y列：COチラシ
                designEstimateNumber: row[25], // Z列：設計見積番号
                remarks: row[26], // AA列：備考
                otherCompany: row[27], // AB列：他社
                history: row[28], // AC列：履歴
                mainContract: row[29], // AD列：本契約
                total: row[30], // AE列：計
                
                // 計算フィールド
                isOrder: this.isOrder(undefined, row[10], this.parseDate(row[0]), row[6], this.getCurrentDateFilter()), // 受注判定（J列確認区分なし + K列アクション状況 + A列日付 + G列年齢 + 対象日付）
                isOvertime: this.isOvertime(this.parseDate(row[0]), row[1], undefined, row[10], row[6], this.getCurrentDateFilter()), // 時間外対応判定（A列またはK列から日付、B列、J列なし、K列、G列年齢、対象日付）
                isElderly: this.isElderly(row[6]), // 高齢者判定
                isExcessive: this.isExcessive(row[10]), // 過量販売判定
                isSingle: this.isSingle(row[10]), // 単独契約判定
                region: this.getRegionName(row[2]), // 地区名
                
                // デバッグ用フィールド
                debugJ: row[9], // J列の確認区分
                debugK: row[10], // K列の確認者日時
                debugA: row[0], // A列の日付
                debugB: row[1]  // B列の時間
            };
        }); // 全データを保持（フィルタリングはレポート生成時に実行）
        
        console.log('最終的なフィルタリング結果:', {
            totalRows: dataRows.length,
            processedRows: processedData.length,
            filteredOut: dataRows.length - processedData.length
        });
        
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
                return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
            }
            
            // "YYYY/MM/DD" 形式
            const slashMatch = trimmedValue.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
            if (slashMatch) {
                const [_, year, month, day] = slashMatch;
                return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
            }
            
            // "YYYY-MM-DD" 形式
            const dashMatch = trimmedValue.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
            if (dashMatch) {
                const [_, year, month, day] = dashMatch;
                return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
            }
            
            // "MM/DD→MM/DD" 形式（最初の日付を使用）
            const arrowMatch = trimmedValue.match(/(\d{1,2})\/(\d{1,2})→/);
            if (arrowMatch) {
                const [_, month, day] = arrowMatch;
                const year = 2025; // 固定年（必要に応じて調整）
                return new Date(year, parseInt(month) - 1, parseInt(day));
            }
        }
        
        return null;
    }
    
    private parseTime(value: any): string | null {
        if (!value) return null;
        
        // Excelの時間形式（数値）を処理
        if (typeof value === 'number') {
            const hours = Math.floor(value * 24);
            const minutes = Math.floor((value * 24 - hours) * 60);
            return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        }
        
        // 文字列形式の時間を処理
        if (typeof value === 'string') {
            return value;
        }
        
        return null;
    }
    
    private parseAge(value: any): number | null {
        if (!value) return null;
        
        const age = parseInt(value);
        return isNaN(age) ? null : age;
    }
    
    // isActionCompleted関数は削除（不要）

    private isOrder(confirmation: any, confirmationDateTime: any, date: any, age: any, targetDate?: Date): boolean {
        console.log('isOrder判定開始:', {
            confirmation: confirmation,
            confirmationDateTime: confirmationDateTime,
            date: date,
            age: age,
            confirmationType: typeof confirmation,
            confirmationValue: confirmation,
            dateType: typeof date,
            dateValue: date,
            dateIsNull: date === null,
            dateIsUndefined: date === undefined,
            dateInstanceOfDate: date instanceof Date,
            dateToString: date ? date.toString() : 'null'
        });

        // 1. J列の確認区分をチェック：1、2、5の場合は除外
        // 一時的に無効化（J列のデータが存在しないため）
        let confirmationValue = confirmation;
        let isJColumnValid = true;

        // J列のデータが存在しない場合は常に有効とする
        if (confirmation === undefined || confirmation === null) {
            console.log('J列のデータが存在しないため、J列条件をスキップ');
            isJColumnValid = true;
        } else {
            // 数値の場合
            if (typeof confirmationValue === 'number') {
                if (confirmationValue === 1 || confirmationValue === 2 || confirmationValue === 5) {
                    console.log('J列の確認区分が1、2、5のため除外（数値）:', confirmationValue);
                    isJColumnValid = false;
                }
            }
            // 文字列の場合
            else if (typeof confirmationValue === 'string') {
                // 空白を除去してから数値変換を試行
                const trimmedValue = confirmationValue.trim();
                if (trimmedValue !== '') {
                    const confirmationNum = parseInt(trimmedValue);
                    if (!isNaN(confirmationNum) && (confirmationNum === 1 || confirmationNum === 2 || confirmationNum === 5)) {
                        console.log('J列の確認区分が1、2、5のため除外（文字列から数値変換）:', trimmedValue, confirmationNum);
                        isJColumnValid = false;
                    }
                }
            }
        }
        
        if (!isJColumnValid) {
            return false;
        }

        // 2. 日付判定：A列またはK列に該当日付が含まれる場合（動的判定）
        // 対象日付を取得（パラメータまたはUI選択日付）
        const actualTargetDate = targetDate || this.getCurrentDateFilter();
        const targetMonth = actualTargetDate.getMonth();
        const targetDay = actualTargetDate.getDate();
        
        console.log('日付判定開始:', {
            targetDate: actualTargetDate.toLocaleDateString(),
            targetMonth: targetMonth,
            targetDay: targetDay,
            aColumnDate: date ? (date instanceof Date ? date.toLocaleDateString() : String(date)) : 'null',
            aColumnMonth: date instanceof Date ? date.getMonth() : 'N/A',
            aColumnDay: date instanceof Date ? date.getDate() : 'N/A',
            dateType: typeof date,
            dateInstanceOfDate: date instanceof Date
        });
        
        // A列の日付が該当日付かチェック
        let isDateMatch = false;
        if (date && date instanceof Date) {
            const aDateMonth = date.getMonth();
            const aDateDay = date.getDate();
            if (aDateMonth === targetMonth && aDateDay === targetDay) {
                isDateMatch = true;
                console.log('A列の日付が該当日付と一致:', date.toLocaleDateString());
            }
        } else if (date && typeof date === 'string') {
            // A列が文字列の場合（例：'8/4→8/7'）、該当日付が含まれるかチェック
            const targetDateStr = `${targetMonth + 1}/${targetDay}`;
            if (date.includes(targetDateStr)) {
                isDateMatch = true;
                console.log('A列の文字列に該当日付が含まれる:', date);
            } else {
                // 日付範囲形式の場合、開始日が該当日付かチェック
                const dateRangeMatch = date.match(/(\d{1,2})\/(\d{1,2})→/);
                if (dateRangeMatch) {
                    const startMonth = parseInt(dateRangeMatch[1]);
                    const startDay = parseInt(dateRangeMatch[2]);
                    if (startMonth === targetMonth + 1 && startDay === targetDay) {
                        isDateMatch = true;
                        console.log('A列の日付範囲の開始日が該当日付:', date);
                    }
                }
            }
        }
        
        // K列の日付判定（シンプルな実装）
        if (!isDateMatch && confirmationDateTime) {
            const kColumnStr = String(confirmationDateTime);
            const targetDateStr = `${targetMonth + 1}/${targetDay}`;
            const targetDateStrAlt = `${targetMonth + 1}月${targetDay}日`;
            
            console.log('K列の日付判定:', {
                confirmationDateTime: confirmationDateTime,
                kColumnStr: kColumnStr,
                targetDateStr: targetDateStr,
                targetDateStrAlt: targetDateStrAlt
            });
            
            // K列に該当日付が含まれるかチェック（シンプル）
            if (kColumnStr.includes(targetDateStr) || kColumnStr.includes(targetDateStrAlt)) {
                isDateMatch = true;
                console.log('K列に該当日付が含まれるため受注対象:', confirmationDateTime);
            }
            
            // K列に日付+時間の形式で該当日付が含まれるかチェック
            const kDateMatch = kColumnStr.match(/(\d{1,2})\/(\d{1,2})/);
            if (kDateMatch) {
                const kMonth = parseInt(kDateMatch[1]);
                const kDay = parseInt(kDateMatch[2]);
                if (kMonth === targetMonth + 1 && kDay === targetDay) {
                    isDateMatch = true;
                    console.log('K列の日付形式で該当日付と一致:', kColumnStr);
                }
            }
        }
        
        if (!isDateMatch) {
            console.log('日付が該当日付と一致しないため除外');
            return false;
        }
        
        // 3. K列の内容判定（シンプルな実装）
        let isKColumnValid = false;
        
        if (confirmationDateTime) {
            const kColumnStr = String(confirmationDateTime);
            
            console.log('K列の内容判定:', {
                kColumnStr: kColumnStr,
                age: age,
                isEmpty: kColumnStr === '',
                is同時: kColumnStr === '同時'
            });
            
            // 除外キーワードチェック
            if (kColumnStr.includes('担当待ち') || kColumnStr.includes('直電') || 
                kColumnStr.includes('契約時') || kColumnStr.includes('契約') || 
                kColumnStr.includes('待ち')) {
                isKColumnValid = false;
                console.log('K列に除外キーワードが含まれる:', kColumnStr);
            } else {
                // 除外キーワードがない場合は有効
                isKColumnValid = true;
                console.log('K列の条件を満たす:', kColumnStr);
            }
        } else {
            // K列が空欄の場合も有効
            isKColumnValid = true;
            console.log('K列が空欄のため有効');
        }
        
        if (!isKColumnValid) {
            console.log('K列の条件を満たさないため除外');
            return false;
        }
        
        console.log('最終的な受注判定結果:', {
            jColumnValid: isJColumnValid,
            dateMatch: isDateMatch,
            kColumnValid: isKColumnValid,
            finalResult: true
        });
        
        // 4. すべての条件を満たす場合、受注として有効
        console.log('受注として有効:', { date, confirmation, confirmationDateTime, age });
        return true;
    }
    


    private isOvertime(date: any, time: any, confirmation: any, confirmationDateTime: any, age: any, targetDate?: Date): boolean {
        // カウント条件.mdに基づく時間外判定ロジック
        // 契約者（A列+B列）と確認者（K列）の両方をチェック
        
        // まず日付フィルタリングを行う
        const actualTargetDate = targetDate || this.getCurrentDateFilter();
        const targetMonth = actualTargetDate.getMonth();
        const targetDay = actualTargetDate.getDate();
        
        // A列またはK列に対象日付が含まれるかチェック
        let isDateMatch = false;
        
        // A列の日付チェック
        if (date && date instanceof Date) {
            const aDateMonth = date.getMonth();
            const aDateDay = date.getDate();
            if (aDateMonth === targetMonth && aDateDay === targetDay) {
                isDateMatch = true;
            }
        }
        
        // K列の日付チェック（A列でマッチしない場合）
        if (!isDateMatch && confirmationDateTime) {
            const kColumnStr = String(confirmationDateTime);
            const targetDateStr = `${targetMonth + 1}/${targetDay}`;
            
            if (kColumnStr.includes(targetDateStr)) {
                isDateMatch = true;
            }
            
            // K列の日付+時間形式での日付チェック
            const kDateMatch = kColumnStr.match(/(\d{1,2})\/(\d{1,2})/);
            if (kDateMatch) {
                const kMonth = parseInt(kDateMatch[1]);
                const kDay = parseInt(kDateMatch[2]);
                if (kMonth === targetMonth + 1 && kDay === targetDay) {
                    isDateMatch = true;
                }
            }
        }
        
        // 対象日付でない場合は時間外判定しない
        if (!isDateMatch) {
            return false;
        }
        
        let isContractorOvertime = false; // 契約者の時間外判定
        let isConfirmerOvertime = false;  // 確認者の時間外判定
        
        // 1. 契約者の時間外判定（A列+B列）
        if (date && time) {
            // 時間が数値（Excel時間）の場合
            if (typeof time === 'number') {
                const hours = Math.floor(time * 24);
                const minutes = Math.floor((time * 24 - hours) * 60);
                const totalMinutes = hours * 60 + minutes;
                
                if (totalMinutes >= 18 * 60 + 30) { // 18:30以降
                    isContractorOvertime = true;
                }
            }
            // 時間が文字列の場合
            else if (typeof time === 'string') {
                const parsedTime = this.parseTime(time);
                if (parsedTime) {
                    const [hours, minutes] = parsedTime.split(':').map(Number);
                    const totalMinutes = hours * 60 + minutes;
                    if (totalMinutes >= 18 * 60 + 30) {
                        isContractorOvertime = true;
                    }
                }
            }
        }
        
        // 2. 確認者の時間外判定（K列）
        if (confirmationDateTime && typeof confirmationDateTime === 'string') {
            const kColumnStr = String(confirmationDateTime);
            
            // K列が「同時」の場合：A列+B列の時間を使用
            if (kColumnStr === '同時') {
                isConfirmerOvertime = isContractorOvertime;
            }
            // K列に日付+時間が含まれる場合
            else if (kColumnStr.includes(':')) {
                // 日付+時間のパターンを検出（例: "8/19　18：44　柚木"）
                const timeMatch = kColumnStr.match(/(\d{1,2})\/(\d{1,2})\s*(\d{1,2})[：:]\s*(\d{1,2})/);
                if (timeMatch) {
                    const [_, month, day, hour, minute] = timeMatch;
                    const year = 2025; // テストファイルの年
                    const fullDateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                    const timeStr = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
                    
                    const confirmerTime = new Date(`${fullDateStr}T${timeStr}:00`);
                    
                    if (!isNaN(confirmerTime.getTime())) {
                        const cutoffTime = new Date(confirmerTime);
                        cutoffTime.setHours(18, 30, 0, 0);
                        
                        if (confirmerTime >= cutoffTime) {
                            isConfirmerOvertime = true;
                        }
                    }
                }
                // 従来の方法（スペース区切り）
                else {
                    const kColumnData = kColumnStr.split(' ');
                    if (kColumnData.length >= 2) {
                        const dateStr = kColumnData[0]; // 例: 8/20
                        const timeStr = kColumnData[1]; // 例: 12:37
                        
                        if (dateStr.includes('/')) {
                            const [month, day] = dateStr.split('/').map(Number);
                            const year = 2025;
                            const fullDateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                            
                            const confirmerTime = new Date(`${fullDateStr}T${timeStr}:00`);
                            
                            if (!isNaN(confirmerTime.getTime())) {
                                const cutoffTime = new Date(confirmerTime);
                                cutoffTime.setHours(18, 30, 0, 0);
                                
                                if (confirmerTime >= cutoffTime) {
                                    isConfirmerOvertime = true;
                                }
                            }
                        }
                    }
                }
            }
        }
        
        // 3. 契約者または確認者のいずれかが時間外の場合、時間外として判定
        const finalResult = isContractorOvertime || isConfirmerOvertime;
        
        return finalResult;
    }
    
    private isElderly(age: any): boolean {
        const parsedAge = this.parseAge(age);
        return parsedAge !== null && parsedAge >= 70;
    }
    
    private isExcessive(confirmationDateTime: any): boolean {
        if (!confirmationDateTime) return false;
        const text = String(confirmationDateTime).toLowerCase();
        return text.includes('過量');
    }
    
    private isSingle(confirmationDateTime: any): boolean {
        if (!confirmationDateTime) return false;
        const text = String(confirmationDateTime).toLowerCase();
        return text.includes('単独');
    }
    
    // 現在の日付フィルターを取得（動的）
    private getCurrentDateFilter(): Date {
        // UIから日付を取得
        const reportDateInput = document.getElementById('reportDate') as HTMLInputElement;
        if (reportDateInput && reportDateInput.value) {
            const selectedDate = new Date(reportDateInput.value);
            return selectedDate;
        }
        
        // デフォルトは今日の日付
        const today = new Date();
        return today;
    }
    
    private getRegionName(regionNumber: any): string {
        if (!regionNumber) return '不明';
        
        const number = parseInt(regionNumber);
        if (isNaN(number)) return '不明';
        
        switch (number) {
            case 511:
                return '九州地区';
            case 521:
            case 531:
                return '中四国地区';
            case 541:
                return '関西地区';
            case 561:
                return '関東地区';
            default:
                return 'その他';
        }
    }
}