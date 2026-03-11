/**
 * 📊 データアクセス
 * =========================================================
 * Google スプレッドシートへの読み書き操作
 */

/**
 * スプレッドシートのセル値（Dateオブジェクト or 文字列）を
 * 統一的に解析して { year, month, day, dateStr } を返すヘルパー
 * @param {Date|string} cellValue - A列（日付）のセル値
 * @returns {{ year: number, month: number, day: number, dateStr: string } | null}
 */
function parseRowDate(cellValue) {
    if (!cellValue) return null;
    if (cellValue instanceof Date) {
        const year = cellValue.getFullYear();
        const month = cellValue.getMonth() + 1; // 1-based
        const day = cellValue.getDate();
        const dateStr = Utilities.formatDate(cellValue, 'Asia/Tokyo', 'yyyy/MM/dd');
        return { year, month, day, dateStr };
    }
    const str = String(cellValue);
    const m = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (!m) return null;
    const year = parseInt(m[1], 10);
    const month = parseInt(m[2], 10);
    const day = parseInt(m[3], 10);
    const dateStr = `${m[1]}/${String(m[2]).padStart(2, '0')}/${String(m[3]).padStart(2, '0')}`;
    return { year, month, day, dateStr };
}

/**
 * スプレッドシートに1行追加
 */
function writeToSpreadsheet(memo, amount, category, method, dateStrInput, account, entryTypeInput) {
    if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID未設定");

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('家計簿');
    if (!sheet) {
        sheet = ss.insertSheet('家計簿');
        sheet.appendRow(["Date", "Amount", "Category", "Memo", "Type", "Method", "IsFixed", "Account"]);
        sheet.getRange(1, 1, 1, 8).setBackground('#e0f7fa').setFontWeight('bold');
    }

    const dateStr = dateStrInput || Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
    const accountName = account || '';
    const entryType = entryTypeInput || '支出';
    const isFixedVal = account ? false : "";

    sheet.appendRow([dateStr, amount, category || '未分類', memo, entryType, method || 'LINE手入力', isFixedVal, accountName]);
}

/**
 * 📱 ダッシュボードから支出を記録するAPI
 */
function addExpenseFromDashboard(memo, amount, category, dateStr, account, typeLabel) {
    try {
        if (!memo) return { success: false, message: "品名を入力してください" };
        const amountValidation = validateAmount(amount);
        if (!amountValidation.valid) return { success: false, message: amountValidation.message };
        const numAmount = amountValidation.value;
        const entryType = typeLabel || '支出';
        writeToSpreadsheet(memo, numAmount, category || '未分類', 'ダッシュボード入力', dateStr, account, entryType);
        return {
            success: true,
            message: `${memo}: ¥${numAmount.toLocaleString()} を記録しました`,
            data: { memo: memo, amount: numAmount, category: category || '未分類', account: account, type: entryType }
        };
    } catch (error) {
        console.error("ダッシュボード入力エラー:", error);
        return { success: false, message: "記録に失敗しました: " + error.message };
    }
}

/**
 * 📋 月別の取引データを取得する（ダッシュボード一覧表示用）
 */
function getMonthlyRecords(year, month) {
    try {
        if (!SPREADSHEET_ID) return { success: false, message: 'SPREADSHEET_ID未設定' };
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('家計簿');
        if (!sheet) return { success: true, records: [] };

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return { success: true, records: [] };

        const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
        const records = [];
        const targetYear = Number(year);
        const targetMonth = Number(month);

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const parsed = parseRowDate(row[0]);
            if (!parsed) continue;
            const { year: rowYear, month: rowMonth, dateStr } = parsed;

            if (rowYear === targetYear && rowMonth === targetMonth) {
                records.push({
                    rowIndex: i + 2,
                    date: dateStr,
                    amount: Number(row[1]) || 0,
                    category: String(row[2] || '未分類'),
                    memo: String(row[3] || ''),
                    type: String(row[4] || '支出'),
                    method: String(row[5] || ''),
                    account: String(row[7] || '')
                });
            }
        }

        records.sort((a, b) => b.date.localeCompare(a.date));
        return { success: true, records: records };
    } catch (error) {
        console.error('取引データ取得エラー:', error);
        return { success: false, message: error.message };
    }
}

/**
 * ✏️ 取引レコードを更新する（ダッシュボード編集用）
 */
function updateRecord(rowIndex, newCategory, newMemo) {
    try {
        if (!SPREADSHEET_ID) return { success: false, message: 'SPREADSHEET_ID未設定' };
        const rowValidation = validateRowIndex(rowIndex);
        if (!rowValidation.valid) return { success: false, message: rowValidation.message + '。全件表示してから再度お試しください。' };
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('家計簿');
        if (!sheet) return { success: false, message: 'シートが見つかりません' };

        if (newCategory !== undefined && newCategory !== null) {
            sheet.getRange(rowIndex, 3).setValue(newCategory);
        }
        if (newMemo !== undefined && newMemo !== null) {
            sheet.getRange(rowIndex, 4).setValue(newMemo);
        }

        return { success: true, message: '更新しました' };
    } catch (error) {
        console.error('レコード更新エラー:', error);
        return { success: false, message: error.message };
    }
}

/**
 * 月間予算を取得（設定シートから。なければデフォルト値）
 */
function getMonthlyBudget(ss) {
    try {
        const settingsSheet = ss.getSheetByName('設定');
        if (settingsSheet) {
            const budgetLabel = settingsSheet.getRange('F1').getValue();
            if (budgetLabel === 'Monthly_Budget') {
                const budget = Number(settingsSheet.getRange('F2').getValue());
                if (budget > 0) return budget;
            }
        }
    } catch (e) {
        // 無視してデフォルト値を返す
    }
    return DEFAULT_MONTHLY_BUDGET;
}

/**
 * 🗑️ 指定した年月のデータを家計簿シートから一括削除する
 * 使い方: GASエディタから手動で deleteDataByMonth(2026, 2) を実行
 */
function deleteDataByMonth(year, month) {
    if (!SPREADSHEET_ID) {
        console.warn('SPREADSHEET_ID が未設定です');
        return;
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');
    if (!sheet) {
        console.warn('家計簿シートが見つかりません');
        return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
        console.warn('データがありません（ヘッダー行のみ）');
        return;
    }

    const targetYear = Number(year);
    const targetMonth = Number(month);
    let deletedCount = 0;
    for (let row = lastRow; row >= 2; row--) {
        const cellValue = sheet.getRange(row, 1).getValue();
        const parsed = parseRowDate(cellValue);
        if (parsed && parsed.year === targetYear && parsed.month === targetMonth) {
            sheet.deleteRow(row);
            deletedCount++;
        }
    }
    console.log(`✅ 削除完了: ${datePrefix} のデータを ${deletedCount} 件削除しました`);
    return deletedCount;
}
