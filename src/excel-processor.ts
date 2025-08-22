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
            const parsedDate = this.parseDate(row[0]); // A列：日付
            
            return {
                rowNumber,
                date: parsedDate,
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
                
                // 計算フィールド（日付に依存しないもののみ事前計算）
                // isOrder: 動的に計算（ReportGeneratorで実行）
                isOvertime: this.isOvertime(parsedDate, row[1], row[10], row[11], row[6]), // 時間外判定
                isElderly: this.isElderly(row[6]), // 高齢者判定
                isExcessive: this.isExcessive(row[10]), // 過量販売判定
                isSingle: this.isSingle(row[10]), // 単独契約判定
                region: this.getRegionName(row[2]), // 地区名
                isHolidayConstruction: false, // 公休日施工（動的に計算）
                isProhibitedConstruction: false, // 禁止日施工（動的に計算）
                
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
        
        // データの詳細分析を追加
        if (processedData.length > 0) {
            const sampleData = processedData.slice(0, 5);
            console.log('サンプルデータ（最初の5件）:', sampleData.map((row, index) => ({
                index: index + 1,
                staffName: row.staffName,
                staffNameType: typeof row.staffName,
                staffNameLength: row.staffName ? row.staffName.length : 0,
                staffNameTrimmed: row.staffName ? row.staffName.trim() : '',
                isOrder: undefined, // 動的に計算されるため
                age: row.contractorAge,
                regionNumber: row.regionNumber,
                departmentNumber: row.departmentNumber
            })));
            
            // 担当者名の統計
            const staffNameStats = {
                total: processedData.length,
                hasStaffName: processedData.filter(row => row.staffName && row.staffName.trim() !== '').length,
                emptyStaffName: processedData.filter(row => !row.staffName || row.staffName.trim() === '').length,
                staffNameTypes: processedData.reduce((acc, row) => {
                    const type = typeof row.staffName;
                    acc[type] = (acc[type] || 0) + 1;
                    return acc;
                }, {} as Record<string, number>)
            };
            console.log('担当者名の統計:', staffNameStats);
        }
        
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
        
        console.log('parseDate入力値:', { value, type: typeof value });
        
        // Excelの日付形式（数値）を処理
        if (typeof value === 'number') {
            // Excelの日付は1900年1月1日からの日数（修正版）
            // 45889 = 2025年8月20日になるように調整
            const excelEpoch = new Date(1900, 0, 1);
            const date = new Date(excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000);
            console.log('parseDate数値結果:', date);
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
                console.log('parseDate和暦結果:', result);
                return result;
            }
            
            // "YYYY/MM/DD" 形式
            const slashMatch = trimmedValue.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
            if (slashMatch) {
                const [_, year, month, day] = slashMatch;
                const result = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                console.log('parseDate YYYY/MM/DD結果:', result);
                return result;
            }
            
            // "YYYY-MM-DD" 形式
            const dashMatch = trimmedValue.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
            if (dashMatch) {
                const [_, year, month, day] = dashMatch;
                const result = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                console.log('parseDate YYYY-MM-DD結果:', result);
                return result;
            }
            
            // "MM/DD→MM/DD" 形式（最初の日付を使用）
            const arrowMatch = trimmedValue.match(/(\d{1,2})\/(\d{1,2})→/);
            if (arrowMatch) {
                const [_, month, day] = arrowMatch;
                const year = 2025; // 固定年（必要に応じて調整）
                const result = new Date(year, parseInt(month) - 1, parseInt(day));
                console.log('parseDate MM/DD→結果:', result);
                return result;
            }
            
            // 直接Date()コンストラクタを試行（フォールバック）
            try {
                const result = new Date(trimmedValue);
                if (!isNaN(result.getTime())) {
                    console.log('parseDate直接変換結果:', result);
                    return result;
                }
            } catch (error) {
                console.log('parseDate直接変換失敗:', error);
            }
        }
        
        console.log('parseDate変換失敗:', value);
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
        // 一時的に簡素化：基本的なデータチェックのみ
        console.log('isOrder判定開始（簡素化版）:', {
            confirmation: confirmation,
            confirmationDateTime: confirmationDateTime,
            date: date,
            age: age
        });

        // 基本的なデータ存在チェック
        if (!date) {
            console.log('日付が存在しないため除外');
            return false;
        }

        // 担当者名の存在チェック（E列）
        // このチェックは後で行うため、ここではスキップ

        console.log('isOrder判定結果: true（簡素化版）');
        return true;
    }
    


    private isSingle(confirmation: any): boolean {
        if (!confirmation) {
            return false;
        }
        
        const confirmationStr = String(confirmation).toLowerCase();
        const result = confirmationStr.includes('単独');
        
        console.log('isSingle判定:', { confirmation, confirmationStr, result });
        return result;
    }

    private isExcessive(confirmation: any): boolean {
        if (!confirmation) {
            return false;
        }
        
        const confirmationStr = String(confirmation).toLowerCase();
        const result = confirmationStr.includes('過量');
        
        console.log('isExcessive判定:', { confirmation, confirmationStr, result });
        return result;
    }

    private isOvertime(date: any, time: any, confirmation: any, confirmationDateTime: any, age: any, targetDate?: Date): boolean {
        console.log('isOvertime判定開始:', {
            date: date instanceof Date ? date.toLocaleDateString() : date,
            time: time,
            confirmation: confirmation,
            confirmationDateTime: confirmationDateTime
        });
        
        // 基本的なデータ存在チェック
        if (!date) {
            console.log('日付が存在しないため除外');
            return false;
        }
        
        // 時間外判定の優先順位（カウント条件.mdに基づく）
        let checkTime: Date | null = null;
        
        // 1. K列が「同時」の場合：A列+B列の日付時間を使用
        if (confirmation && String(confirmation).toLowerCase() === '同時') {
            if (date instanceof Date && time) {
                const timeStr = String(time);
                const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})/);
                if (timeMatch) {
                    checkTime = new Date(date);
                    checkTime.setHours(parseInt(timeMatch[1]), parseInt(timeMatch[2]), 0, 0);
                }
            }
        }
        // 2. K列に日付+時間が含まれる場合：K列の時間を使用
        else if (confirmation) {
            const confirmationStr = String(confirmation);
            const timeMatch = confirmationStr.match(/(\d{1,2})[：:]\s*(\d{2})/);
            if (timeMatch) {
                checkTime = new Date();
                checkTime.setHours(parseInt(timeMatch[1]), parseInt(timeMatch[2]), 0, 0);
            }
        }
        
        // 3. K列に時間情報がない場合：A列+B列の時間を使用
        if (!checkTime && date instanceof Date && time) {
            const timeStr = String(time);
            const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})/);
            if (timeMatch) {
                checkTime = new Date(date);
                checkTime.setHours(parseInt(timeMatch[1]), parseInt(timeMatch[2]), 0, 0);
            }
        }
        
        // 18:30以降かチェック
        if (checkTime) {
            const overtimeThreshold = new Date(checkTime);
            overtimeThreshold.setHours(18, 30, 0, 0);
            
            const result = checkTime >= overtimeThreshold;
            console.log('isOvertime判定結果:', {
                checkTime: checkTime.toLocaleTimeString(),
                threshold: '18:30',
                result: result
            });
            return result;
        }
        
        console.log('isOvertime判定結果: false（時間情報なし）');
        return false;
    }
    
    private isElderly(age: any): boolean {
        const parsedAge = this.parseAge(age);
        return parsedAge !== null && parsedAge >= 70;
    }
    

    
    // 現在の日付フィルターを取得（動的）
    private getCurrentDateFilter(): Date {
        // UIから日付を取得
        const reportDateInput = document.getElementById('reportDate') as HTMLInputElement;
        console.log('getCurrentDateFilter - UI要素チェック:', {
            element: !!reportDateInput,
            value: reportDateInput?.value,
            type: reportDateInput?.type
        });
        
        if (reportDateInput && reportDateInput.value) {
            // タイムゾーンの問題を回避するため、YYYY-MM-DD形式から直接作成
            const [year, month, day] = reportDateInput.value.split('-').map(Number);
            const selectedDate = new Date(year, month - 1, day); // monthは0ベース
            console.log('getCurrentDateFilter - 正常パス:', {
                inputValue: reportDateInput.value,
                splitResult: [year, month, day],
                selectedDate: selectedDate.toLocaleDateString(),
                selectedDateISO: selectedDate.toISOString()
            });
            return selectedDate;
        }
        
        // デフォルトは今日の日付
        const today = new Date();
        console.log('getCurrentDateFilter - デフォルトパス:', {
            today: today.toLocaleDateString(),
            todayISO: today.toISOString()
        });
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