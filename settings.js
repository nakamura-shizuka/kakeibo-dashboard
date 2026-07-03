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

        const warnings = [];

        let fixedExpenses = [];
        if (sheet.getRange('F6').getValue() === 'Fixed_Expenses') {
            const savedStr = sheet.getRange('G6').getValue();
            if (savedStr) {
                try {
                    fixedExpenses = JSON.parse(savedStr);
                } catch (e) {
                    logError('固定費設定パース失敗', e.message + ' / 保存値: ' + savedStr);
                    warnings.push('固定費設定の読み込みに失敗しました');
                }
            }
        }

        let accounts = [];
        if (sheet.getRange('F7').getValue() === 'Accounts_List') {
            const accStr = sheet.getRange('G7').getValue();
            if (accStr) {
                try {
                    accounts = JSON.parse(accStr);
                } catch (e) {
                    logError('口座設定パース失敗', e.message + ' / 保存値: ' + accStr);
                    warnings.push('口座設定の読み込みに失敗しました');
                }
            }
        }

        // ===== 計画層の設定（Layer 0: 家計モデル） =====
        let income = 0;
        if (sheet.getRange('F8').getValue() === 'Monthly_Income') {
            income = Number(sheet.getRange('G8').getValue()) || 0;
        }

        let nisaMonthly = 0;
        if (sheet.getRange('F9').getValue() === 'NISA_Monthly') {
            nisaMonthly = Number(sheet.getRange('G9').getValue()) || 0;
        }

        let annualEvents = [];
        if (sheet.getRange('F10').getValue() === 'Annual_Events') {
            const eventsStr = sheet.getRange('G10').getValue();
            if (eventsStr) {
                try {
                    annualEvents = JSON.parse(eventsStr);
                } catch (e) {
                    logError('年間イベント設定パース失敗', e.message + ' / 保存値: ' + eventsStr);
                    warnings.push('年間イベント設定の読み込みに失敗しました');
                }
            }
        }

        let reserveRate = DEFAULT_RESERVE_RATE;
        if (sheet.getRange('F11').getValue() === 'Reserve_Rate') {
            const savedRate = sheet.getRange('G11').getValue();
            if (savedRate !== '' && savedRate !== null) reserveRate = Number(savedRate) || 0;
        }

        const result = {
            budget: budget, categories: categories, fixedExpenses: fixedExpenses, accounts: accounts,
            income: income, nisaMonthly: nisaMonthly, annualEvents: annualEvents, reserveRate: reserveRate
        };
        if (warnings.length > 0) result.warning = warnings.join(' / ');
        return result;
    } catch (e) {
        logError('設定データ取得失敗', e.stack || e.message);
        return { budget: DEFAULT_MONTHLY_BUDGET, categories: "", fixedExpenses: [], accounts: [], warning: '設定データの取得に失敗しました' };
    }
}

/**
 * ⚙️ ユーザーの設定データを保存する
 */
function saveSettingsData(budget, categoriesStr, fixedExpensesStr, accountsStr, income, nisaMonthly, annualEventsStr, reserveRate) {
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

        if (income !== undefined) {
            sheet.getRange('F8').setValue('Monthly_Income');
            sheet.getRange('G8').setValue(Number(income) || 0);
        }

        if (nisaMonthly !== undefined) {
            sheet.getRange('F9').setValue('NISA_Monthly');
            sheet.getRange('G9').setValue(Number(nisaMonthly) || 0);
        }

        if (annualEventsStr !== undefined) {
            sheet.getRange('F10').setValue('Annual_Events');
            sheet.getRange('G10').setValue(annualEventsStr);
        }

        if (reserveRate !== undefined) {
            sheet.getRange('F11').setValue('Reserve_Rate');
            sheet.getRange('G11').setValue(Number(reserveRate) || 0);
        }

        // 設定変更は当月のKPI計算に影響するためキャッシュを無効化
        try {
            const now = new Date();
            invalidateDashboardCache(now.getFullYear(), now.getMonth());
        } catch (e) { /* キャッシュ無効化失敗は無視 */ }

        return { success: true };
    } catch (e) {
        return { success: false, error: e.toString() };
    }
}
