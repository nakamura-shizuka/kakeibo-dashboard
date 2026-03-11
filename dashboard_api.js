/**
 * 📊 ダッシュボード API
 * =========================================================
 * ダッシュボード表示用データ取得関数群
 */

/**
 * 📊 ダッシュボードデータを一括取得するAPI
 * @param {number} targetYear - 対象年 (例: 2026) 省略時は現在年
 * @param {number} targetMonth - 対象月 (0-11) 省略時は現在月
 */
function getDashboardData(targetYear, targetMonth) {
    if (!SPREADSHEET_ID) return { error: "SPREADSHEET_ID未設定" };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');

    const now = new Date();
    const currentYear = targetYear !== undefined ? targetYear : now.getFullYear();
    const currentMonth = targetMonth !== undefined ? targetMonth : now.getMonth();

    if (!sheet || sheet.getLastRow() <= 1) {
        return {
            totalSpending: 0,
            totalIncome: 0,
            carryOver: 0,
            budget: getMonthlyBudget(ss),
            categories: [],
            recentRecords: [],
            monthLabel: currentYear + "年" + (currentMonth + 1) + "月"
        };
    }

    const settings = getSettingsData();
    let accountBalances = {};
    if (settings.accounts && settings.accounts.length > 0) {
        settings.accounts.forEach(acc => {
            accountBalances[acc.name] = Number(acc.balance) || 0;
        });
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();

    const thisMonthData = [];
    let carryOverIncome = 0;
    let carryOverSpending = 0;

    data.forEach(function (row, idx) {
        if (!row[0]) return;
        const d = new Date(row[0]);
        const rYear = d.getFullYear();
        const rMonth = d.getMonth();
        const amount = Number(row[1]) || 0;
        const type = row[4] || '支出';

        let accountName = row[7] ? row[7].toString().trim() : '';
        if (!accountName) accountName = '未設定';

        if (accountBalances[accountName] !== undefined) {
            if (type === '収入') accountBalances[accountName] += amount;
            else accountBalances[accountName] -= amount;
        } else {
            accountBalances[accountName] = 0;
            if (type === '収入') accountBalances[accountName] += amount;
            else accountBalances[accountName] -= amount;
        }

        if (rYear < currentYear || (rYear === currentYear && rMonth < currentMonth)) {
            if (type === '収入') carryOverIncome += amount;
            else carryOverSpending += amount;
        } else if (rYear === currentYear && rMonth === currentMonth) {
            row._rowIndex = idx + 2;
            thisMonthData.push(row);
        }
    });

    const carryOver = carryOverIncome - carryOverSpending;

    let totalSpending = 0;
    let totalIncome = 0;
    const categoryMap = {};

    thisMonthData.forEach(function (row) {
        const amount = Number(row[1]) || 0;
        const category = row[2] || '未分類';
        const type = row[4] || '支出';

        if (type === '収入') {
            totalIncome += amount;
        } else {
            totalSpending += amount;
            categoryMap[category] = (categoryMap[category] || 0) + amount;
        }
    });

    const categories = Object.keys(categoryMap).map(function (key) {
        return { name: key, amount: categoryMap[key] };
    }).sort(function (a, b) { return b.amount - a.amount; });

    const recentRecords = thisMonthData
        .map(function (row) {
            return {
                rowIndex: row._rowIndex || 0,
                _ts: new Date(row[0]).getTime(),
                date: Utilities.formatDate(new Date(row[0]), "Asia/Tokyo", "M/d"),
                amount: Number(row[1]) || 0,
                category: row[2] || '未分類',
                memo: row[3] || '',
                type: row[4] || '支出',
                method: row[5] || ''
            };
        })
        .sort(function (a, b) { return b._ts - a._ts; })
        .slice(0, 10);

    recentRecords.forEach(function (r) { delete r._ts; });

    let customCategories = null;
    try {
        const settingsSheet = ss.getSheetByName('設定');
        if (settingsSheet) {
            const label = settingsSheet.getRange('F5').getValue();
            if (label === 'Custom_Categories') {
                const catsRaw = settingsSheet.getRange('G5').getValue();
                if (catsRaw) {
                    customCategories = catsRaw.toString().split(',').map(c => c.trim()).filter(c => c);
                }
            }
        }
    } catch (e) { console.warn('カスタムカテゴリ取得失敗:', e.message); }

    if (customCategories && customCategories.length > 0) {
        customCategories.forEach(function (catName) {
            if (!categoryMap[catName]) {
                categories.push({ name: catName, amount: 0 });
            }
        });
    }

    let aiMessage = "";
    try {
        const settingsSheet = ss.getSheetByName('設定');
        if (settingsSheet) {
            const label = settingsSheet.getRange('F4').getValue();
            if (label === 'AI_Message') {
                aiMessage = settingsSheet.getRange('G4').getValue();
            }
        }
    } catch (e) { console.warn('AIメッセージ取得失敗:', e.message); }

    return {
        totalSpending: totalSpending,
        totalIncome: totalIncome,
        carryOver: carryOver,
        budget: getMonthlyBudget(ss),
        categories: categories,
        recentRecords: recentRecords,
        aiMessage: aiMessage,
        accountBalances: accountBalances,
        monthLabel: currentYear + "年" + (currentMonth + 1) + "月"
    };
}

/**
 * 🌊 サンキーダイアグラム用データを取得
 */
function getSankeyData(targetYear, targetMonth) {
    if (!SPREADSHEET_ID) return { flows: [] };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');
    if (!sheet || sheet.getLastRow() <= 1) return { flows: [] };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    const now = new Date();
    const currentYear = targetYear !== undefined ? targetYear : now.getFullYear();
    const currentMonth = targetMonth !== undefined ? targetMonth : now.getMonth();

    const thisMonthData = data.filter(function (row) {
        if (!row[0]) return false;
        const d = new Date(row[0]);
        return d.getFullYear() === currentYear && d.getMonth() === currentMonth;
    });

    let totalIncome = 0;
    let totalSpending = 0;
    const categoryMap = {};

    thisMonthData.forEach(function (row) {
        const amount = Number(row[1]) || 0;
        const type = row[4] || '支出';
        const category = row[2] || '未分類';

        if (type === '収入') {
            totalIncome += amount;
        } else {
            totalSpending += amount;
            categoryMap[category] = (categoryMap[category] || 0) + amount;
        }
    });

    const flows = [];
    const sourceAmount = totalIncome > 0 ? totalIncome : getMonthlyBudget(ss);
    const sourceLabel = totalIncome > 0 ? '収入' : '予算';

    Object.keys(categoryMap).forEach(function (category) {
        flows.push([sourceLabel, category, categoryMap[category]]);
    });

    const remaining = sourceAmount - totalSpending;
    if (remaining > 0) {
        flows.push([sourceLabel, '残高', remaining]);
    }

    return {
        flows: flows,
        totalIncome: totalIncome,
        totalSpending: totalSpending,
        sourceLabel: sourceLabel,
        sourceAmount: sourceAmount
    };
}

/**
 * 📈 年間レポート用データを取得
 */
function getYearlyReportData(targetYear) {
    if (!SPREADSHEET_ID) return { error: "SPREADSHEET_ID未設定" };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');
    const now = new Date();
    const currentYear = targetYear !== undefined ? targetYear : now.getFullYear();

    if (!sheet || sheet.getLastRow() <= 1) {
        return { year: currentYear, monthlyData: [] };
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

    const monthlyData = [];
    for (let i = 0; i < 12; i++) {
        monthlyData.push({ month: i + 1, income: 0, expense: 0, savings: 0, cumulativeSavings: 0 });
    }

    let carryOverIncome = 0;
    let carryOverSpending = 0;

    data.forEach(function (row) {
        if (!row[0]) return;
        const d = new Date(row[0]);
        const rYear = d.getFullYear();
        const rMonth = d.getMonth();
        const amount = Number(row[1]) || 0;
        const type = row[4] || '支出';

        if (rYear < currentYear) {
            if (type === '収入') carryOverIncome += amount;
            else carryOverSpending += amount;
        } else if (rYear === currentYear) {
            if (type === '収入') {
                monthlyData[rMonth].income += amount;
            } else {
                monthlyData[rMonth].expense += amount;
            }
        }
    });

    let cumulative = carryOverIncome - carryOverSpending;
    monthlyData.forEach(function (m) {
        m.savings = m.income - m.expense;
        cumulative += m.savings;
        m.cumulativeSavings = cumulative;
    });

    return { year: currentYear, monthlyData: monthlyData };
}
