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
/**
 * 🗑️ ダッシュボードキャッシュを無効化する
 * データ更新後（addExpense, updateRecord）に呼び出す
 */
function invalidateDashboardCache(year, month) {
    const cache = CacheService.getScriptCache();
    cache.remove('dashboard_' + year + '_' + month);
    cache.remove('sankey_' + year + '_' + month);
    cache.remove('yearly_' + year);
}

/**
 * 💴 確定ベースの月間支出を集計する（Layer 1）
 * 確定支出 = 月次確定シートのカード請求額 + 家計簿シートの非カード明細（現金・固定費・手入力）
 * 家計簿シートのカード明細（CARD_METHODS）は請求確定額と二重になるため総額に算入しない。
 * @param {number} year - 対象年
 * @param {number} month - 対象月 (0-11)
 * @returns {Object} { total, cardTotal, nonCardSpending, cardBreakdown: [{card, amount, confirmed}] }
 */
function calcConfirmedSpending_(year, month) {
    const result = { total: 0, cardTotal: 0, nonCardSpending: 0, cardBreakdown: [] };
    if (!SPREADSHEET_ID) return result;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const billingMonth = year + '/' + String(month + 1).padStart(2, '0');

    // カード請求確定額（月次確定シート）
    const cardAmounts = {}; // card -> { amount, confirmed }
    const confirmedSheet = ss.getSheetByName(SHEET_MONTHLY_CONFIRMED);
    if (confirmedSheet && confirmedSheet.getLastRow() > 1) {
        const data = confirmedSheet.getRange(2, 1, confirmedSheet.getLastRow() - 1, 6).getValues();
        data.forEach(function (row) {
            const rowMonth = row[0] instanceof Date
                ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM')
                : String(row[0]).trim();
            if (rowMonth !== billingMonth) return;
            const card = String(row[1]).trim();
            const amount = Number(row[2]) || 0;
            cardAmounts[card] = { amount: amount, confirmed: amount > 0 };
        });
    }

    BILLING_CARDS.forEach(function (card) {
        const entry = cardAmounts[card];
        result.cardBreakdown.push({
            card: card,
            amount: entry ? entry.amount : 0,
            confirmed: entry ? entry.confirmed : false
        });
        if (entry) result.cardTotal += entry.amount;
    });

    // 非カード明細（家計簿シートのうち Method が CARD_METHODS 以外の支出）
    const kakeiboSheet = ss.getSheetByName('家計簿');
    if (kakeiboSheet && kakeiboSheet.getLastRow() > 1) {
        const data = kakeiboSheet.getRange(2, 1, kakeiboSheet.getLastRow() - 1, 8).getValues();
        data.forEach(function (row) {
            if (!row[0]) return;
            const d = new Date(row[0]);
            if (d.getFullYear() !== year || d.getMonth() !== month) return;
            if ((row[4] || '支出') !== '支出') return;
            const method = String(row[5] || '').trim();
            if (CARD_METHODS.indexOf(method) >= 0) return; // カード明細は請求確定額側で計上
            result.nonCardSpending += Number(row[1]) || 0;
        });
    }

    result.total = result.cardTotal + result.nonCardSpending;
    return result;
}

/**
 * 💰 貯蓄サマリー（計画層 + Layer 1）を構築する
 * 家計モデル: 使っていい額 = 月収 − 固定費 − NISA積立 − 年間イベント月割り − 予備費
 * @param {Object} settings - getSettingsData() の結果
 * @param {number} year - 対象年
 * @param {number} month - 対象月 (0-11)
 * @param {Array} thisMonthData - 家計簿シートの当月行（速報消化の算出用）
 */
function buildSavingsSummary_(settings, year, month, thisMonthData) {
    const income = Number(settings.income) || 0;
    const nisaMonthly = Number(settings.nisaMonthly) || 0;
    const reserveRate = settings.reserveRate !== undefined ? Number(settings.reserveRate) : DEFAULT_RESERVE_RATE;

    let fixedSum = 0;
    (settings.fixedExpenses || []).forEach(function (item) {
        fixedSum += Number(item.amount) || 0;
    });

    let eventsTotal = 0;
    (settings.annualEvents || []).forEach(function (ev) {
        eventsTotal += Number(ev.amount) || 0;
    });
    const eventsMonthly = Math.round(eventsTotal / 12);

    const reserve = Math.round(income * reserveRate / 100);
    const safeToSpend = income - fixedSum - nisaMonthly - eventsMonthly - reserve;

    // 変動費の速報値（当月明細のうち固定費以外の支出。カード明細も含む=リアルタイムの参考値）
    let provisionalVariable = 0;
    (thisMonthData || []).forEach(function (row) {
        if ((row[4] || '支出') !== '支出') return;
        if (String(row[5] || '').trim() === '自動(固定費)') return;
        provisionalVariable += Number(row[1]) || 0;
    });

    const confirmed = calcConfirmedSpending_(year, month);

    // 実測貯蓄: 当月末時点の残高合計 − 前月末時点の残高合計（残高シートから）
    const monthEnd = new Date(year, month + 1, 0, 23, 59, 59);
    const prevMonthEnd = new Date(year, month, 0, 23, 59, 59);
    const balances = getBalancesAsOf_(monthEnd);
    const prevBalances = getBalancesAsOf_(prevMonthEnd);
    const actualSavings = (balances.latestDate && prevBalances.latestDate)
        ? balances.total - prevBalances.total
        : null;

    return {
        income: income,
        nisaMonthly: nisaMonthly,
        fixedSum: fixedSum,
        eventsMonthly: eventsMonthly,
        reserve: reserve,
        reserveRate: reserveRate,
        safeToSpend: safeToSpend,
        provisionalVariable: provisionalVariable,
        confirmedSpending: confirmed.total,
        cardBreakdown: confirmed.cardBreakdown,
        nonCardSpending: confirmed.nonCardSpending,
        // 現金ベースの貯蓄見込み（NISA積立は別枠の資産形成として除外）
        projectedSavings: income - nisaMonthly - confirmed.total,
        actualSavings: actualSavings,
        balanceTotal: balances.latestDate ? balances.total : null,
        balanceDate: balances.latestDate
    };
}

function getDashboardData(targetYear, targetMonth) {
    if (!SPREADSHEET_ID) return { error: "SPREADSHEET_ID未設定" };

    const now = new Date();
    const currentYear = targetYear !== undefined ? targetYear : now.getFullYear();
    const currentMonth = targetMonth !== undefined ? targetMonth : now.getMonth();

    // キャッシュチェック（5分間）
    const cache = CacheService.getScriptCache();
    const cacheKey = 'dashboard_' + currentYear + '_' + currentMonth;
    const cached = cache.get(cacheKey);
    if (cached) {
        try { return JSON.parse(cached); } catch (e) { /* キャッシュ破損時は再取得 */ }
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');

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

    let savings = null;
    try {
        savings = buildSavingsSummary_(settings, currentYear, currentMonth, thisMonthData);
    } catch (e) {
        console.warn('貯蓄サマリー構築失敗:', e.message);
    }

    const result = {
        totalSpending: totalSpending,
        totalIncome: totalIncome,
        carryOver: carryOver,
        budget: getMonthlyBudget(ss),
        categories: categories,
        recentRecords: recentRecords,
        aiMessage: aiMessage,
        savings: savings,
        monthLabel: currentYear + "年" + (currentMonth + 1) + "月"
    };
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) { /* データが大きすぎる場合はスキップ */ }
    return result;
}

/**
 * 🌊 サンキーダイアグラム用データを取得
 */
function getSankeyData(targetYear, targetMonth) {
    if (!SPREADSHEET_ID) return { flows: [] };

    const now = new Date();
    const currentYear = targetYear !== undefined ? targetYear : now.getFullYear();
    const currentMonth = targetMonth !== undefined ? targetMonth : now.getMonth();

    const cache = CacheService.getScriptCache();
    const cacheKey = 'sankey_' + currentYear + '_' + currentMonth;
    const cached = cache.get(cacheKey);
    if (cached) {
        try { return JSON.parse(cached); } catch (e) { }
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');
    if (!sheet || sheet.getLastRow() <= 1) return { flows: [] };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

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

    const sankeyResult = {
        flows: flows,
        totalIncome: totalIncome,
        totalSpending: totalSpending,
        sourceLabel: sourceLabel,
        sourceAmount: sourceAmount
    };
    try { cache.put(cacheKey, JSON.stringify(sankeyResult), 600); } catch (e) { }
    return sankeyResult;
}

/**
 * 📈 年間レポート用データを取得
 */
function getYearlyReportData(targetYear) {
    if (!SPREADSHEET_ID) return { error: "SPREADSHEET_ID未設定" };

    const now = new Date();
    const currentYear = targetYear !== undefined ? targetYear : now.getFullYear();

    const cache = CacheService.getScriptCache();
    const cacheKey = 'yearly_' + currentYear;
    const cached = cache.get(cacheKey);
    if (cached) {
        try { return JSON.parse(cached); } catch (e) { }
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');

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

    const yearlyResult = { year: currentYear, monthlyData: monthlyData };
    try { cache.put(cacheKey, JSON.stringify(yearlyResult), 600); } catch (e) { }
    return yearlyResult;
}
