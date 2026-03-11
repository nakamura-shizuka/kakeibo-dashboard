/**
 * ⚙️ 設定管理
 * =========================================================
 * ユーザー設定（予算・カテゴリ・固定費・口座）の読み書き
 */

/**
 * ⚙️ 設定データを取得する（初期表示用）
 */
function getSettingsData() {
    if (!SPREADSHEET_ID) return { budget: DEFAULT_MONTHLY_BUDGET, categories: "", fixedExpenses: [] };
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName('設定');
        if (!sheet) return { budget: DEFAULT_MONTHLY_BUDGET, categories: "", fixedExpenses: [] };

        let budget = DEFAULT_MONTHLY_BUDGET;
        if (sheet.getRange('F1').getValue() === 'Monthly_Budget') {
            budget = Number(sheet.getRange('F2').getValue()) || DEFAULT_MONTHLY_BUDGET;
        }

        let categories = "食費,日用品,交通費,娯楽,医療,衣服,交際費,その他";
        if (sheet.getRange('F5').getValue() === 'Custom_Categories') {
            const savedCats = sheet.getRange('G5').getValue();
            if (savedCats) categories = savedCats;
        }

        let fixedExpenses = [];
        if (sheet.getRange('F6').getValue() === 'Fixed_Expenses') {
            const savedStr = sheet.getRange('G6').getValue();
            if (savedStr) {
                try {
                    fixedExpenses = JSON.parse(savedStr);
                } catch (e) { console.warn('固定費設定パース失敗:', e.message); }
            }
        }

        let accounts = [];
        if (sheet.getRange('F7').getValue() === 'Accounts_List') {
            const accStr = sheet.getRange('G7').getValue();
            if (accStr) {
                try {
                    accounts = JSON.parse(accStr);
                } catch (e) { console.warn('口座設定パース失敗:', e.message); }
            }
        }

        return { budget: budget, categories: categories, fixedExpenses: fixedExpenses, accounts: accounts };
    } catch (e) {
        return { budget: DEFAULT_MONTHLY_BUDGET, categories: "", fixedExpenses: [], accounts: [] };
    }
}

/**
 * ⚙️ ユーザーの設定データを保存する
 */
function saveSettingsData(budget, categoriesStr, fixedExpensesStr, accountsStr) {
    if (!SPREADSHEET_ID) return { success: false, error: 'DB未設定' };
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName('設定');
        if (!sheet) {
            sheet = ss.insertSheet('設定');
        }

        sheet.getRange('F1').setValue('Monthly_Budget');
        sheet.getRange('F2').setValue(Number(budget) || DEFAULT_MONTHLY_BUDGET);

        const cleanCats = categoriesStr.split(',')
            .map(c => c.trim())
            .filter(c => c.length > 0)
            .join(',');
        sheet.getRange('F5').setValue('Custom_Categories');
        sheet.getRange('G5').setValue(cleanCats);

        if (fixedExpensesStr !== undefined) {
            sheet.getRange('F6').setValue('Fixed_Expenses');
            sheet.getRange('G6').setValue(fixedExpensesStr);
        }

        if (accountsStr !== undefined) {
            sheet.getRange('F7').setValue('Accounts_List');
            sheet.getRange('G7').setValue(accountsStr);
        }

        return { success: true };
    } catch (e) {
        return { success: false, error: e.toString() };
    }
}
