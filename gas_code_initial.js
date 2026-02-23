/**
 * 魔法の家計簿 - メインスクリプト (Phase 2-A)
 * =========================================================
 * LINE Bot Webhook + ダッシュボードAPI + データ配信
 *
 * 【バージョン履歴】
 * - v1.0: LINE Bot基本機能（Phase 1）
 * - v1.1: doPost安定性改善・署名検証堅牢化
 * - v2.0: ダッシュボードAPI・HTML配信追加（Phase 2-A）
 */

// ===== 定数・設定 =====
const PROPERTIES = PropertiesService.getScriptProperties();
const LINE_ACCESS_TOKEN = (PROPERTIES.getProperty('LINE_ACCESS_TOKEN') || "").trim();
const LINE_CHANNEL_SECRET = (PROPERTIES.getProperty('LINE_CHANNEL_SECRET') || "").trim();
const SPREADSHEET_ID = (PROPERTIES.getProperty('SPREADSHEET_ID') || "").trim();
const GEMINI_API_KEY = (PROPERTIES.getProperty('GEMINI_API_KEY') || "").trim();

// 月間予算（デフォルト値。設定シートで上書き可能）
const DEFAULT_MONTHLY_BUDGET = 120000;

// ===== Phase 1: 初期セットアップ =====

/**
 * 📝 スプレッドシート（DB）の自動作成と紐付け
 */
function createDatabase() {
    const ss = SpreadsheetApp.create("みえる化家計簿DB");
    const ssId = ss.getId();

    const sheetKakeibo = ss.getSheets()[0];
    sheetKakeibo.setName('家計簿');
    const kakeiboHeaders = [["Date", "Amount", "Category", "Memo", "Type", "Method", "IsFixed"]];
    sheetKakeibo.getRange(1, 1, 1, 7).setValues(kakeiboHeaders);
    sheetKakeibo.getRange(1, 1, 1, 7).setBackground('#e0f7fa').setFontWeight('bold');

    const sheetSettings = ss.insertSheet('設定');
    const settingsHeaders = [["Fixed_Date", "Fixed_Name", "Fixed_Amount", "Fixed_Category", "Fixed_Method"]];
    sheetSettings.getRange(1, 1, 1, 5).setValues(settingsHeaders);
    sheetSettings.getRange(1, 1, 1, 5).setBackground('#fff9c4').setFontWeight('bold');

    PROPERTIES.setProperty('SPREADSHEET_ID', ssId);

    console.log('✨ 完 成 ✨');
    console.log(`DB作成完了: ${ss.getUrl()}`);
}

// ===== Phase 1: LINE Bot Webhook =====

/**
 * LINEからのWebhookを受け取るエンドポイント
 */
function doPost(e) {
    // ダッシュボードAPIリクエスト判定（actionパラメータまたはJSON bodyにaction含む）
    const contentType = e.postData ? e.postData.type : '';
    let bodyJson = null;

    try {
        bodyJson = e.postData ? JSON.parse(e.postData.contents) : {};
    } catch (_) {
        bodyJson = {};
    }

    // ダッシュボード API（actionフィールドがある場合）
    if (bodyJson.action) {
        let result;
        try {
            switch (bodyJson.action) {
                case 'updateRecord':
                    result = updateRecord(
                        Number(bodyJson.rowIndex),
                        bodyJson.category !== undefined ? bodyJson.category : null,
                        bodyJson.memo !== undefined ? bodyJson.memo : null
                    );
                    break;
                case 'addExpense':
                    result = addExpenseFromDashboard(
                        bodyJson.memo,
                        bodyJson.amount,
                        bodyJson.category,
                        bodyJson.date,
                        bodyJson.account,
                        bodyJson.type
                    );
                    break;
                case 'saveSettings':
                    result = saveSettingsData(
                        bodyJson.budget,
                        bodyJson.categories,
                        bodyJson.fixedCosts,
                        bodyJson.accounts
                    );
                    break;
                default:
                    result = { success: false, message: '不明なaction: ' + bodyJson.action };
            }
        } catch (err) {
            result = { success: false, message: 'APIエラー: ' + err.message };
        }
        return ContentService.createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);
    }

    // ===== LINE Webhook 処理（従来のロジック） =====
    try {
        // 署名検証
        if (LINE_CHANNEL_SECRET) {
            const signature = e.headers ? (e.headers['X-Line-Signature'] || e.headers['x-line-signature']) : null;
            if (!signature) {
                return createJsonResponse({ 'content': 'ok' });
            }
            const expectedSignature = Utilities.base64Encode(
                Utilities.computeHmacSha256Signature(
                    Utilities.newBlob(e.postData.contents).getBytes(),
                    Utilities.newBlob(LINE_CHANNEL_SECRET).getBytes()
                )
            );
            if (signature !== expectedSignature) {
                console.error("署名検証失敗: 不正アクセス");
                return createJsonResponse({ 'content': 'invalid signature' });
            }
        }

        const json = JSON.parse(e.postData.contents);
        const events = json.events;

        if (!events || events.length === 0) {
            return createJsonResponse({ 'content': 'ok' });
        }

        const replyToken = events[0].replyToken;
        if (isTestToken(replyToken)) {
            return createJsonResponse({ 'content': 'ok' });
        }

        const userId = events[0].source ? events[0].source.userId : null;
        if (userId) {
            saveUserIdToSettings(userId);
        }

        const userMessage = events[0].message ? events[0].message.text : "";
        if (!userMessage) {
            return createJsonResponse({ 'content': 'ok' });
        }

        const replyMessage = processMessage(userMessage);
        sendLineReply(replyToken, replyMessage);

    } catch (error) {
        console.error("【致命的エラー】doPost:", error);
    }

    return createJsonResponse({ 'content': 'ok' });
}

/**
 * メッセージ解析 → スプレッドシート記録
 */
function processMessage(userMessage) {
    const match = userMessage.match(/^(.+?)[\s　]+([0-9０-９,，]+)円?$/);

    if (!match) {
        return "📝 使い方ガイド\n\n「品名 金額」の形式で送ってね！\n\n✅ 例：\n・ランチ 1200\n・コンビニ 350\n・電車代 500";
    }

    const memo = match[1].trim();
    const amountStr = match[2]
        .replace(/[,，]/g, "")
        .replace(/[０-９]/g, function (s) {
            return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
        });
    const amount = parseInt(amountStr, 10);

    if (isNaN(amount) || amount <= 0) {
        return "❌ 金額を正しく読み取れませんでした。";
    }

    try {
        writeToSpreadsheet(memo, amount);
        return `✅ 記録完了！\n📦 ${memo}: ${amount.toLocaleString()}円\n家計簿にバッチリ追記しました🧾`;
    } catch (error) {
        console.error("記録エラー:", error);
        return `❌ 記録失敗: ${error.message}`;
    }
}

/**
 * スプレッドシートに1行追加
 */
function writeToSpreadsheet(memo, amount, category, method, dateStrInput, account) {
    if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID未設定");

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('家計簿');
    if (!sheet) {
        sheet = ss.insertSheet('家計簿');
        sheet.appendRow(["Date", "Amount", "Category", "Memo", "Type", "Method", "IsFixed", "Account"]);
        sheet.getRange(1, 1, 1, 8).setBackground('#e0f7fa').setFontWeight('bold');
    }

    const dateStr = dateStrInput || Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
    const accountName = account || '未設定';
    const entryType = typeLabel || '支出';
    sheet.appendRow([dateStr, amount, category || '未分類', memo, entryType, method || 'LINE手入力', false, accountName]);
}

/**
 * 📱 ダッシュボードから支出を記録するAPI
 * クライアント側から google.script.run.addExpenseFromDashboard() で呼ばれる
 */
function addExpenseFromDashboard(memo, amount, category, dateStr, account, typeLabel) {
    try {
        if (!memo || !amount) {
            return { success: false, message: "品名と金額を入力してください" };
        }
        const numAmount = parseInt(String(amount).replace(/[,，]/g, ""), 10);
        if (isNaN(numAmount) || numAmount <= 0) {
            return { success: false, message: "金額は正の数値で入力してください" };
        }
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

        console.log(`getMonthlyRecords: year=${targetYear}, month=${targetMonth}, データ行数=${data.length}`);

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            let dateStr = '';
            let rowYear = 0;
            let rowMonth = 0;

            if (row[0] instanceof Date) {
                dateStr = Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd');
                rowYear = row[0].getFullYear();
                rowMonth = row[0].getMonth() + 1;
            } else {
                dateStr = String(row[0]);
                // "2026/02/21" or "2026-02-21" 形式をパース
                const parts = dateStr.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
                if (parts) {
                    rowYear = parseInt(parts[1]);
                    rowMonth = parseInt(parts[2]);
                    dateStr = `${parts[1]}/${String(parts[2]).padStart(2, '0')}/${String(parts[3]).padStart(2, '0')}`;
                }
            }

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

        console.log(`getMonthlyRecords: ${records.length}件の記録が見つかりました`);
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
        if (!rowIndex || rowIndex < 2) return { success: false, message: '行番号が不正です（rowIndex=' + rowIndex + '）。全件表示してから再度お試しください。' };
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('家計簿');
        if (!sheet) return { success: false, message: 'シートが見つかりません' };

        // カテゴリ（C列=3）とメモ（D列=4）を更新
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
 * LINE返信送信
 */
function sendLineReply(replyToken, message) {
    if (!LINE_ACCESS_TOKEN) return;

    const url = 'https://api.line.me/v2/bot/message/reply';
    const payload = {
        'replyToken': replyToken,
        'messages': [{ 'type': 'text', 'text': message }]
    };

    try {
        const response = UrlFetchApp.fetch(url, {
            'method': 'POST',
            'headers': { "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
            'contentType': 'application/json',
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
        });
        if (response.getResponseCode() !== 200) {
            logError("LINE API返信エラー", response.getContentText());
        }
    } catch (err) {
        logError("LINE API例外", err.toString());
    }
}

/**
 * 送信元のLINEユーザーIDを設定シートに保存する（Push送信用）
 */
function saveUserIdToSettings(userId) {
    if (!SPREADSHEET_ID || !userId) return;
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName('設定');
        if (!sheet) return;

        // F3セルに LINE_USER_ID を保存
        const currentId = sheet.getRange('G3').getValue();
        if (currentId !== userId) {
            sheet.getRange('F3').setValue('LINE_USER_ID');
            sheet.getRange('G3').setValue(userId);
        }
    } catch (e) {
        // 例外は無視
    }
}

/**
 * 🚨 LINEへプッシュ通知を送る（アラート等用）
 */
function pushLineMessage(userId, message) {
    if (!LINE_ACCESS_TOKEN || !userId) return;

    const url = 'https://api.line.me/v2/bot/message/push';
    const payload = {
        'to': userId,
        'messages': [{ 'type': 'text', 'text': message }]
    };

    try {
        const response = UrlFetchApp.fetch(url, {
            'method': 'POST',
            'headers': { "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
            'contentType': 'application/json',
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
        });
        if (response.getResponseCode() !== 200) {
            logError("LINE PushAPI エラー", response.getContentText());
        }
    } catch (err) {
        logError("LINE PushAPI 例外", err.toString());
    }
}
// ===== Phase 2-A: ダッシュボード =====

/**
 * GETリクエストの処理（JSON API ルーター / フォールバック: ダッシュボードHTML配信）
 */
function doGet(e) {
    const action = e && e.parameter ? e.parameter.action : null;

    // JSON APIルーター
    if (action) {
        let result;
        try {
            switch (action) {
                case 'getDashboardData':
                    result = getDashboardData(
                        e.parameter.year ? Number(e.parameter.year) : undefined,
                        e.parameter.month ? Number(e.parameter.month) : undefined
                    );
                    break;
                case 'getSettingsData':
                    result = getSettingsData();
                    break;
                case 'getSankeyData':
                    result = getSankeyData(
                        Number(e.parameter.year),
                        Number(e.parameter.month)
                    );
                    break;
                case 'getYearlyReportData':
                    result = getYearlyReportData(Number(e.parameter.year));
                    break;
                case 'getAiAnalysis':
                    result = getAiAnalysis(e.parameter.isWeekly === 'true');
                    break;
                case 'getMonthlyRecords':
                    result = getMonthlyRecords(
                        Number(e.parameter.year),
                        Number(e.parameter.month)
                    );
                    break;
                case 'updateRecord':
                    result = updateRecord(
                        Number(e.parameter.rowIndex),
                        e.parameter.category !== undefined ? e.parameter.category : null,
                        e.parameter.memo !== undefined ? e.parameter.memo : null
                    );
                    break;
                case 'addExpense':
                    result = addExpenseFromDashboard(
                        e.parameter.memo,
                        e.parameter.amount,
                        e.parameter.category,
                        e.parameter.date,
                        e.parameter.account,
                        e.parameter.type
                    );
                    break;
                case 'saveSettings':
                    result = saveSettingsData(
                        e.parameter.budget,
                        e.parameter.categories,
                        e.parameter.fixedCosts,
                        e.parameter.accounts
                    );
                    break;
                default:
                    result = { success: false, message: '不明なaction: ' + action };
            }
        } catch (err) {
            result = { success: false, message: 'APIエラー: ' + err.message };
        }
        return ContentService.createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);
    }

    // ステータスページ
    const page = e && e.parameter && e.parameter.page ? e.parameter.page : null;
    if (page === 'status') {
        return renderStatusPage();
    }

    // フォールバック: ダッシュボードHTML配信（GAS直接アクセス時）
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('みえる化家計簿')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * 📊 ダッシュボードデータを一括取得するAPI
 * クライアント側から google.script.run.getDashboardData(year, month) で呼ばれる
 * @param {number} targetYear - 対象年 (例: 2023) 省略時は現在年
 * @param {number} targetMonth - 対象月 (0-11) 省略時は現在月
 */
function getDashboardData(targetYear, targetMonth) {
    if (!SPREADSHEET_ID) return { error: "SPREADSHEET_ID未設定" };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');

    const now = new Date();
    const currentYear = targetYear !== undefined ? targetYear : now.getFullYear();
    const currentMonth = targetMonth !== undefined ? targetMonth : now.getMonth(); // 0-indexed

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

    // 「設定」シートから口座リスト（初期残高）を取得する
    const settings = getSettingsData();
    let accountBalances = {};
    if (settings.accounts && settings.accounts.length > 0) {
        settings.accounts.forEach(acc => {
            // 初期残高をセット
            accountBalances[acc.name] = Number(acc.balance) || 0;
        });
    }

    // H列（8列目）まで取得: [日時, 金額, カテゴリ, 品名, 収支(収入/支出), 登録元, UID, 口座名]
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();

    // 今月のデータと、先月末までのデータに分ける
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
        if (!accountName) accountName = '未設定'; // H列が空の場合は未設定

        // 資産（口座）残高の計算 (全期間対象)
        if (accountBalances[accountName] !== undefined) {
            if (type === '収入') accountBalances[accountName] += amount;
            else accountBalances[accountName] -= amount;
        } else {
            // 設定リストに無い口座が出てきた場合、0から計算を開始する
            accountBalances[accountName] = 0;
            if (type === '収入') accountBalances[accountName] += amount;
            else accountBalances[accountName] -= amount;
        }

        // 指定月より前（繰越金の計算対象）
        if (rYear < currentYear || (rYear === currentYear && rMonth < currentMonth)) {
            if (type === '収入') carryOverIncome += amount;
            else carryOverSpending += amount;
        }
        // 指定月のデータ（行インデックスも保持）
        else if (rYear === currentYear && rMonth === currentMonth) {
            row._rowIndex = idx + 2; // スプレッドシートの実際の行番号
            thisMonthData.push(row);
        }
    });

    const carryOver = carryOverIncome - carryOverSpending;

    // 支出合計
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

    // カテゴリ別データ（配列に変換し、金額降順）
    const categories = Object.keys(categoryMap).map(function (key) {
        return { name: key, amount: categoryMap[key] };
    }).sort(function (a, b) { return b.amount - a.amount; });

    // 直近10件（日付の新しい順）
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

    // シリアライズ不要なプロパティを削除
    recentRecords.forEach(function (r) { delete r._ts; });

    // カスタムカテゴリ設定の取得（設定シート G5セル、カンマ区切り。F5に 'Custom_Categories'）
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
    } catch (e) { }

    // 既存のカテゴリ支出データ（categories）に、0円のカスタムカテゴリもマージして表示枠を確保する
    if (customCategories && customCategories.length > 0) {
        customCategories.forEach(function (catName) {
            if (!categoryMap[catName]) {
                categories.push({ name: catName, amount: 0 });
            }
        });
        // 指定された順序にある程度沿いつつ、金額降順にするならソートロジックを工夫、今回は金額降順を維持
    }

    // AIアドバイスの取得（設定シート G4セルを想定。F4に 'AI_Message'）
    let aiMessage = "";
    try {
        const settingsSheet = ss.getSheetByName('設定');
        if (settingsSheet) {
            const label = settingsSheet.getRange('F4').getValue();
            if (label === 'AI_Message') {
                aiMessage = settingsSheet.getRange('G4').getValue();
            }
        }
    } catch (e) { }

    return {
        totalSpending: totalSpending,
        totalIncome: totalIncome,
        carryOver: carryOver,
        budget: getMonthlyBudget(ss),
        categories: categories,
        recentRecords: recentRecords,
        aiMessage: aiMessage,
        accountBalances: accountBalances, // 口座別残高データ
        monthLabel: currentYear + "年" + (currentMonth + 1) + "月"
    };
}

/**
 * 🌊 サンキーダイアグラム用データを取得
 * クライアント側から google.script.run.getSankeyData(year, month) で呼ばれる
 * @param {number} targetYear - 対象年 (例: 2023) 省略時は現在年
 * @param {number} targetMonth - 対象月 (0-11) 省略時は現在月
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

    // サンキー用のフロー（[from, to, amount]）
    const flows = [];

    // 収入がない場合は予算ベースで表示
    const sourceAmount = totalIncome > 0 ? totalIncome : getMonthlyBudget(ss);
    const sourceLabel = totalIncome > 0 ? '収入' : '予算';

    // 収入/予算 → 各カテゴリ
    Object.keys(categoryMap).forEach(function (category) {
        flows.push([sourceLabel, category, categoryMap[category]]);
    });

    // 残高
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
 * クライアント側から google.script.run.getYearlyReportData(year) で呼ばれる
 * @param {number} targetYear - 対象年 (例: 2023) 省略時は現在年
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

    // 1月〜12月の初期化
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

        // 前年以前（初期繰越額の算出用）
        if (rYear < currentYear) {
            if (type === '収入') carryOverIncome += amount;
            else carryOverSpending += amount;
        }
        // 対象年
        else if (rYear === currentYear) {
            if (type === '収入') {
                monthlyData[rMonth].income += amount;
            } else {
                monthlyData[rMonth].expense += amount;
            }
        }
    });

    let cumulative = carryOverIncome - carryOverSpending;

    // 累計貯蓄を計算
    monthlyData.forEach(function (m) {
        m.savings = m.income - m.expense;
        cumulative += m.savings;
        m.cumulativeSavings = cumulative;
    });

    return {
        year: currentYear,
        monthlyData: monthlyData
    };
}

// ===== Phase 8: AIによる客観的家計分析と浪費検知 =====

/**
 * 🤖 Gemini APIを使用して家計データを分析し、冷静かつ客観的なフィードバックを生成する
 * @param {boolean} isWeekly - true=週次分析, false=月次分析
 */
function generateAiAnalysis(isWeekly) {
    if (!GEMINI_API_KEY) return "AI分析機能が有効ではありません（GEMINI_API_KEY未設定）。";
    if (!SPREADSHEET_ID) return "DBが設定されていません。";

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');
    if (!sheet || sheet.getLastRow() <= 1) return "分析するデータがありません。";

    // データの取得と集計準備
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();
    const currentDay = now.getDate();
    const monthlyBudget = getMonthlyBudget(ss);

    let currentPeriodExpense = 0;
    let previousPeriodExpense = 0;
    const currentCategoryMap = {};
    const previousCategoryMap = {};

    let periodLabel = isWeekly ? "今週" : "今月";
    let prevPeriodLabel = isWeekly ? "先週" : "先月";

    // 期間の判定ロジック
    // ※今回は簡易的に、isWeeklyの場合は直近7日間 vs その前の7日間を比較。
    // 月次の場合は今月 vs 先月を比較する。
    const MS_PER_DAY = 1000 * 60 * 60 * 24;

    data.forEach(row => {
        if (!row[0] || row[4] !== '支出') return;
        const d = new Date(row[0]);
        const amount = Number(row[1]) || 0;
        const category = row[2] || '未分類';

        if (isWeekly) {
            const diffDays = Math.floor((now.getTime() - d.getTime()) / MS_PER_DAY);
            if (diffDays >= 0 && diffDays < 7) {
                // 過去7日間 (今週相当)
                currentPeriodExpense += amount;
                currentCategoryMap[category] = (currentCategoryMap[category] || 0) + amount;
            } else if (diffDays >= 7 && diffDays < 14) {
                // 8日〜14日前 (先週相当)
                previousPeriodExpense += amount;
                previousCategoryMap[category] = (previousCategoryMap[category] || 0) + amount;
            }
        } else {
            // 月次
            const rYear = d.getFullYear();
            const rMonth = d.getMonth();
            if (rYear === currentYear && rMonth === currentMonth) {
                currentPeriodExpense += amount;
                currentCategoryMap[category] = (currentCategoryMap[category] || 0) + amount;
            } else if (rYear === currentYear && rMonth === currentMonth - 1 || (currentMonth === 0 && rYear === currentYear - 1 && rMonth === 11)) {
                previousPeriodExpense += amount;
                previousCategoryMap[category] = (previousCategoryMap[category] || 0) + amount;
            }
        }
    });

    // プロンプト用データの整形
    const currentCatStr = Object.keys(currentCategoryMap).map(k => `・${k}: ${currentCategoryMap[k]}円`).join('\n') || "記録なし";
    const prevCatStr = Object.keys(previousCategoryMap).map(k => `・${k}: ${previousCategoryMap[k]}円`).join('\n') || "記録なし";

    // 進行度（今月の場合）
    let budgetProgressStr = "";
    if (!isWeekly) {
        const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
        const progressPercent = Math.round((currentDay / daysInMonth) * 100);
        const expensePercent = Math.round((currentPeriodExpense / monthlyBudget) * 100);
        budgetProgressStr = `【月間予算】: ${monthlyBudget}円 (日数経過: ${progressPercent}%、予算消化: ${expensePercent}%)`;
    } else {
        // 週次の場合は、月間予算の1/4を簡易目標とする
        const weeklyBudget = Math.floor(monthlyBudget / 4);
        const expensePercent = Math.round((currentPeriodExpense / weeklyBudget) * 100);
        budgetProgressStr = `【週次目安予算（月予算の1/4）】: ${weeklyBudget}円 (消化率: ${expensePercent}%)`;
    }

    // Gemini へのシステムプロンプト（人格設定）
    const systemPrompt = `
あなたは優秀で冷徹な専属データアナリストです。
家計簿のデータに基づき、感情を排して客観的かつ冷静に分析を行ってください。
以下の要件に厳密に従って出力してください。

1. 感情的な言葉（「頑張りましょう」「残念ですね」等）や、絵文字、過剰な装飾は一切不要です。
2. 提示された「${periodLabel}の支出」と「${prevPeriodLabel}の支出」、および予算データに基づいて、明確な事実のみを端的に述べること。
3. 特に、先期からの出費の大幅な増加や、目安予算を超過している「浪費（使いすぎ）ポイント」があれば、カテゴリと金額を挙げて鋭く指摘すること。
4. 全体として計画通りのペースか、それとも赤字ペースかを最後に結語として1〜2文で断定すること。
5. 出力はMarkdownのリスト形式等を用い、スマホのLINEやWeb画面で読みやすく簡潔にまとめること（最大でも400文字程度）。
`;

    // ユーザープロンプト（データ入力）
    const userPrompt = `
以下の家計データから分析レポートを作成してください。

${budgetProgressStr}

■ ${prevPeriodLabel}の支出合計: ${previousPeriodExpense}円
${prevCatStr}

■ ${periodLabel}の支出合計: ${currentPeriodExpense}円
${currentCatStr}
`;

    // Gemini API リクエスト
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
    const payload = {
        "contents": [
            { "role": "user", "parts": [{ "text": systemPrompt + "\n\n" + userPrompt }] }
        ],
        "generationConfig": {
            "temperature": 0.2, // 冷静・客観的にするため低めに設定
            "maxOutputTokens": 800
        }
    };

    try {
        const response = UrlFetchApp.fetch(url, {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        const resJson = JSON.parse(response.getContentText());
        if (resJson.error) {
            console.error("Gemini API Error:", resJson.error);
            return "分析エラー: AIへの接続に失敗しました。";
        }

        const answer = resJson.candidates[0].content.parts[0].text;
        return answer.trim();

    } catch (e) {
        console.error("AI分析実行エラー:", e);
        return "分析エラー: システムエラーが発生しました。";
    }
}

/**
 * 📊 ダッシュボードの即時分析ボタンから呼ばれるAPI
 */
function getAiAnalysis(isWeekly) {
    try {
        const resultText = generateAiAnalysis(isWeekly);
        return { success: true, analysis: resultText };
    } catch (error) {
        return { success: false, message: error.message };
    }
}

/**
 * ⏰ 定期実行トリガー用：週次レポート送信（毎週日曜の夕方などを想定）
 */
function sendWeeklyReport() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName('設定');
    if (!settingsSheet) return;

    // F3cel='LINE_USER_ID', G3cel=UserID
    const userId = settingsSheet.getRange('G3').getValue();
    if (!userId) {
        console.warn("LINE_USER_IDが設定されていないため、通知をスキップしました。LINEからBotへ一度メッセージを送ってください。");
        return;
    }

    const analysisResult = generateAiAnalysis(true); // 週次
    const message = "📊 【みえる化家計簿】週次データ分析レポート\n\n" + analysisResult;

    pushLineMessage(userId, message);
}

/**
 * ⏰ 定期実行トリガー用：月次レポート送信（毎月1日の朝などを想定）
 */
function sendMonthlyReport() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName('設定');
    if (!settingsSheet) return;

    const userId = settingsSheet.getRange('G3').getValue();
    if (!userId) return;

    // 月次は前月分の振り返りをしたいケースが多いため、月初実行時（1日）は事実上、
    // currentMonthの「1日分のデータ」と前月の比較になってしまう。
    // そのため、1日〜3日の間に直近30日間として分析するなどロジックの調整が必要だが、
    // ここでは MVP として、そのまま generateAiAnalysis(false) を呼ぶ。
    // （※ generateAiAnalysis 内で、今月vs先月の比較を行っている）
    const analysisResult = generateAiAnalysis(false); // 月次
    const message = "📈 【みえる化家計簿】月次データ分析レポート\n\n" + analysisResult;

    pushLineMessage(userId, message);
}

/**
 * 月間予算を取得（設定シートから。なければデフォルト値）
 */
function getMonthlyBudget(ss) {
    try {
        const settingsSheet = ss.getSheetByName('設定');
        if (settingsSheet) {
            // F1に「Monthly_Budget」、F2に金額 があれば使う
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

// ===== Phase 6: 設定(カスタマイズ)機能 =====

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

        let categories = "食費,日用品,交通費,娯楽,医療,衣服,交際費,その他"; // デフォルト
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
                } catch (e) { }
            }
        }

        let accounts = [];
        if (sheet.getRange('F7').getValue() === 'Accounts_List') {
            const accStr = sheet.getRange('G7').getValue();
            if (accStr) {
                try {
                    accounts = JSON.parse(accStr);
                } catch (e) { }
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

        // 予算の保存 (F1, F2)
        sheet.getRange('F1').setValue('Monthly_Budget');
        sheet.getRange('F2').setValue(Number(budget) || DEFAULT_MONTHLY_BUDGET);

        // カテゴリの保存 (F5, G5)
        const cleanCats = categoriesStr.split(',')
            .map(c => c.trim())
            .filter(c => c.length > 0)
            .join(',');

        sheet.getRange('F5').setValue('Custom_Categories');
        sheet.getRange('G5').setValue(cleanCats);

        // 固定費の保存 (F6, G6)
        if (fixedExpensesStr !== undefined) {
            sheet.getRange('F6').setValue('Fixed_Expenses');
            sheet.getRange('G6').setValue(fixedExpensesStr);
        }

        // 口座情報の保存 (F7, G7)
        if (accountsStr !== undefined) {
            sheet.getRange('F7').setValue('Accounts_List');
            sheet.getRange('G7').setValue(accountsStr);
        }

        return { success: true };
    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

// ===== Phase 6: 固定費・アラート機能 =====

/**
 * 📅 毎日定期実行で呼び出す固定費自動記録関数
 */
function autoRecordFixedExpenses() {
    if (!SPREADSHEET_ID) return;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName('設定');
    const kakeiboSheet = ss.getSheetByName('家計簿');
    if (!settingsSheet || !kakeiboSheet) return;

    // F6, G6 から固定費JSONを読み取る
    let fixedExpenses = [];
    if (settingsSheet.getRange('F6').getValue() === 'Fixed_Expenses') {
        const savedStr = settingsSheet.getRange('G6').getValue();
        if (savedStr) {
            try {
                fixedExpenses = JSON.parse(savedStr);
            } catch (e) { }
        }
    }

    if (fixedExpenses.length === 0) return;

    const now = new Date();
    const todayDate = now.getDate();
    // 翌日の日付が1なら今日は月末
    const isEndOfMonth = (new Date(now.getFullYear(), now.getMonth(), todayDate + 1).getDate() === 1);

    // 今日記録すべき固定費を抽出
    const targets = fixedExpenses.filter(item => {
        const targetDate = parseInt(item.date, 10);
        // 設定日が今日
        if (targetDate === todayDate) return true;
        // 月末日の場合、設定日が今日より大きければ強制的に今日記録する（例: 31日設定で今月が30日までの場合）
        if (isEndOfMonth && targetDate > todayDate) return true;
        return false;
    });

    if (targets.length === 0) return;

    // 今月の既存家計簿データを取得（重複防止用）
    const lastRow = kakeiboSheet.getLastRow();
    const recordedKeys = new Set();
    const targetUserId = settingsSheet.getRange('G3').getValue() || "";

    if (lastRow > 1) {
        // [日時, 金額, カテゴリ, メモ] を取得
        const data = kakeiboSheet.getRange(2, 1, lastRow - 1, 4).getValues();
        data.forEach(row => {
            if (!row[0]) return;
            const d = new Date(row[0]);
            // 今月のデータのみ
            if (d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth()) {
                const amount = Number(row[1]) || 0;
                const cat = row[2] || '';
                const memo = row[3] || '';
                recordedKeys.add(`${cat}_${memo}_${amount}`);
            }
        });
    }

    const recordedItems = [];

    // 固定費の記帳処理
    targets.forEach(item => {
        const amount = Number(item.amount) || 0;
        const memo = item.memo || '固定費';
        const cat = item.category || '未分類';

        const key = `${cat}_${memo}_${amount}`; // 重複判定キー

        // すでに今月同額同名の記録があればスキップ
        if (recordedKeys.has(key)) return;

        const timeStamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
        kakeiboSheet.insertRowAfter(1);
        kakeiboSheet.getRange("A2:G2").setValues([[
            timeStamp,
            amount,
            cat,
            memo,
            "支出",
            "自動(固定費)",
            targetUserId
        ]]);

        recordedItems.push(`${memo} (${amount.toLocaleString()}円)`);
        recordedKeys.add(key); // このループ内の重複防止
    });

    // 通知処理
    if (recordedItems.length > 0 && targetUserId) {
        const msg = `🤖 【固定費の自動記録】\n\n本日設定されていた以下の固定費を記録しました！\n\n` +
            recordedItems.map(i => `・${i}`).join('\n') +
            `\n\n※すでに同じ記録がある場合はスキップされています。`;
        pushLineMessage(targetUserId, msg);
    }
}


/**
 * 毎日定期実行（タイムドリブン）で呼び出す予算監視関数
 * その月の支出合計が「予算の80%」や「100%」を超えたらPush通知を送る
 */
function checkBudgetAndAlert() {
    if (!SPREADSHEET_ID) return;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName('設定');
    if (!settingsSheet) return;

    // F3, G3 セルからLINE USER IDを取得
    const targetUserId = settingsSheet.getRange('G3').getValue();
    if (!targetUserId) {
        console.log("LINE USER ID が未設定のためアラートをスキップします。");
        return;
    }

    // 今月の総支出と予算を取得
    const now = new Date();
    const dashData = getDashboardData(now.getFullYear(), now.getMonth());
    const spent = dashData.totalSpending || 0;
    const budget = dashData.budget || DEFAULT_MONTHLY_BUDGET;

    if (budget <= 0) return;

    const percent = (spent / budget) * 100;
    const currentMonthStr = `${now.getFullYear()}-${now.getMonth() + 1}`;

    // 設定シートの K列/L列 にアラートフラグを保持すると仮定
    // K1: "Alert_Month", L1: (例: "2026-2")
    // K2: "Alert_80_Sent", L2: (TRUE/FALSE)
    // K3: "Alert_100_Sent", L3: (TRUE/FALSE)

    const flagMonth = settingsSheet.getRange('L1').getValue();

    // もし月が変わっていればフラグをリセット
    if (flagMonth !== currentMonthStr) {
        settingsSheet.getRange('K1').setValue('Alert_Month');
        settingsSheet.getRange('L1').setValue(currentMonthStr);
        settingsSheet.getRange('K2').setValue('Alert_80_Sent');
        settingsSheet.getRange('L2').setValue(false);
        settingsSheet.getRange('K3').setValue('Alert_100_Sent');
        settingsSheet.getRange('L3').setValue(false);
    }

    const is80Sent = settingsSheet.getRange('L2').getValue();
    const is100Sent = settingsSheet.getRange('L3').getValue();

    // 100% 超過チェック
    if (percent >= 100 && !is100Sent) {
        const msg = `🚨 【予算超過アラート】\n\n今月の支出が予算（${budget.toLocaleString()}円）を超えました！\n現在: ${spent.toLocaleString()}円（${Math.round(percent)}%）\n\n来月に向けて支出ペースを見直しましょう💦`;
        pushLineMessage(targetUserId, msg);
        settingsSheet.getRange('L3').setValue(true); // 送信済みフラグを立てる
        return; // 100%を送るなら80%は同時に送らない
    }

    // 80% 超過チェック
    if (percent >= 80 && percent < 100 && !is80Sent) {
        const msg = `⚠️ 【予算アラート】\n\n今月の支出が予算の80%を超えました。\n残り: ${(budget - spent).toLocaleString()}円\n\n月末まで少し節約を意識してみましょう👀`;
        pushLineMessage(targetUserId, msg);
        settingsSheet.getRange('L2').setValue(true); // 送信済みフラグを立てる
    }
}

/**
 * 🤖 Gemini APIを呼び出してテキストを生成する
 */
function callGeminiAPI(promptText) {
    if (!GEMINI_API_KEY) return "AIアドバイザーは現在お休み中です（APIキー未設定）";

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;

    // Gemini 1.5 用のペイロード構造
    const payload = {
        "contents": [{
            "parts": [{ "text": promptText }]
        }],
        "generationConfig": {
            "temperature": 0.7,
            "maxOutputTokens": 300
        }
    };

    try {
        const response = UrlFetchApp.fetch(url, {
            'method': 'POST',
            'headers': { 'Content-Type': 'application/json' },
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
        });

        if (response.getResponseCode() === 200) {
            const result = JSON.parse(response.getContentText());
            if (result.candidates && result.candidates[0].content.parts.length > 0) {
                return result.candidates[0].content.parts[0].text;
            }
        } else {
            logError("Gemini API エラー", response.getContentText());
        }
    } catch (e) {
        logError("Gemini API 例外", e.toString());
    }
    return "AIの分析中にエラーが発生しました💦 また後で試してね。";
}

/**
 * 📅 毎週/毎月実行される AI予算分析・通知関数
 */
function analyzeBudgetWithAI() {
    if (!SPREADSHEET_ID) return;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName('設定');
    if (!settingsSheet) return;

    const targetUserId = settingsSheet.getRange('G3').getValue(); // push通知用

    const now = new Date();
    const currentMonthLabel = `${now.getFullYear()}年${now.getMonth() + 1}月`;
    const data = getDashboardData(now.getFullYear(), now.getMonth());

    // カテゴリごとの支出をテキスト化
    const catText = data.categories.map(c => `・${c.name}: ${c.amount}円`).join('\n');
    const remaining = data.budget - data.totalSpending;

    // AIへの指示書き（プロンプト）
    const prompt = `
あなたは優秀で親しみやすい「魔法の家計簿」のAIアドバイザーです。
以下の今月（${currentMonthLabel}）の家計簿データをもとに、ユーザーにアドバイスを送ってください。

【データ】
・今月の予算: ${data.budget}円
・現在の支出合計: ${data.totalSpending}円
・現在の残額: ${remaining}円
・カテゴリ別支出:
${catText}

【厳守するルール】
1. トーンは親しみやすく、絵文字（✨や💡など）を適度に使ってください。
2. 長すぎる文章は避け、**3行〜4行程度**に簡潔にまとめてください。
3. まずはこれまでの頑張りを褒め、その後、一番支出が多いカテゴリについて「どうすれば無理なく節約できるか」の軽い提案を1つだけ入れてください。
4. Markdown記法（太字など）は使用不可で、プレーンテキストのみを出力してください。
`;

    const aiResponse = callGeminiAPI(prompt);
    const finalMessage = `🤖 【AI家計アドバイス】\n\n${aiResponse}\n\n※このメッセージはAIが作成しました✨`;

    // 1. 設定シート (F4, G4) に最新メッセージを保存（ダッシュボード表示用）
    settingsSheet.getRange('F4').setValue('AI_Message');
    settingsSheet.getRange('G4').setValue(finalMessage);

    // 2. ユーザーへLINE Push通知
    if (targetUserId) {
        pushLineMessage(targetUserId, finalMessage);
    }
}

// ===== ユーティリティ =====

function createJsonResponse(data) {
    return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}

function isTestToken(token) {
    return ['00000000000000000000000000000000', 'ffffffffffffffffffffffffffffffff', '11111111111111111111111111111111'].indexOf(token) !== -1;
}

function logError(errorType, details) {
    try {
        if (!SPREADSHEET_ID) return;
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let debugSheet = ss.getSheetByName('エラーログ');
        if (!debugSheet) {
            debugSheet = ss.insertSheet('エラーログ');
            debugSheet.appendRow(["日時", "エラー種別", "詳細"]);
        }
        debugSheet.appendRow([new Date(), errorType, details]);
    } catch (ignored) { }
}

/**
 * 旧ステータス確認画面（?page=status で表示）
 */
function renderStatusPage() {
    const token = PROPERTIES.getProperty('LINE_ACCESS_TOKEN') || "";
    const secret = PROPERTIES.getProperty('LINE_CHANNEL_SECRET') || "";
    const ssId = PROPERTIES.getProperty('SPREADSHEET_ID') || "";

    let dbStatus = "未確認";
    if (ssId) {
        try {
            SpreadsheetApp.openById(ssId.trim());
            dbStatus = "✅ アクセス成功";
        } catch (err) {
            dbStatus = "❌ アクセス失敗: " + err.message;
        }
    }

    let html = "<div style='font-family:sans-serif;padding:20px;line-height:1.6'>";
    html += "<h2 style='color:#009688'>接続ステータス確認</h2><ul style='list-style:none;padding:0'>";
    html += "<li style='margin:8px 0;padding:10px;background:" + (token ? "#e8f5e9" : "#ffebee") + ";border-radius:5px'><b>LINE Token:</b> " + (token ? "✅ 設定あり" : "❌ 未設定") + "</li>";
    html += "<li style='margin:8px 0;padding:10px;background:" + (secret ? "#e8f5e9" : "#ffebee") + ";border-radius:5px'><b>LINE Secret:</b> " + (secret ? "✅ 設定あり" : "❌ 未設定") + "</li>";
    html += "<li style='margin:8px 0;padding:10px;background:" + (ssId ? "#e8f5e9" : "#ffebee") + ";border-radius:5px'><b>DB ID:</b> " + (ssId ? "✅ " + ssId : "❌ 未設定") + "</li>";
    html += "<li style='margin:8px 0;padding:10px;background:" + (dbStatus.includes("✅") ? "#e8f5e9" : "#fff3e0") + ";border-radius:5px'><b>DB接続:</b> " + dbStatus + "</li>";
    html += "</ul></div>";

    return HtmlService.createHtmlOutput(html).setTitle("ステータス確認");
}

// ===== Phase 7: データクリーンアップ =====

/**
 * 🗑️ 指定した年月のデータを家計簿シートから一括削除する
 * 使い方: GASエディタから手動で deleteDataByMonth(2026, 2) を実行
 */
function deleteDataByMonth(year, month) {
    if (!SPREADSHEET_ID) {
        console.log('SPREADSHEET_ID が未設定です');
        return;
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');
    if (!sheet) {
        console.log('家計簿シートが見つかりません');
        return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
        console.log('データがありません（ヘッダー行のみ）');
        return;
    }

    // 削除対象行を後ろから検索して削除（行番号がずれないよう逆順で）
    const datePrefix = `${year}/${String(month).padStart(2, '0')}`;
    let deletedCount = 0;
    for (let row = lastRow; row >= 2; row--) {
        const cellValue = sheet.getRange(row, 1).getValue();
        const dateStr = cellValue instanceof Date
            ? Utilities.formatDate(cellValue, 'Asia/Tokyo', 'yyyy/MM')
            : String(cellValue).substring(0, 7);
        if (dateStr === datePrefix) {
            sheet.deleteRow(row);
            deletedCount++;
        }
    }
    console.log(`✅ 削除完了: ${datePrefix} のデータを ${deletedCount} 件削除しました`);
    return deletedCount;
}

// ===== Phase 7: Gmail自動連携（レイヤー1）=====

/**
 * 🏷️ 店名からカテゴリを自動推定する
 */
function guessCategory_(shopName) {
    if (!shopName) return '未分類';
    const s = shopName.toLowerCase();

    // 食費
    if (/スーパー|イオン|ウエルシア|セブン|ファミマ|ローソン|マクド|モス|ケンタッキー|くら寿司|すき家|吉野家|松屋|なか卯|王将|ココス|食品|ピザ|パン|ベーカリー|カフェ|スタバ|ドトール|コーヒー|レストラン|居酒屋|食堂|弁当|ガスト|デニーズ|バーガー|ランチ|うどん|そば|ラーメン|焼肉|定食|コンビニ|飲食|グルメ|ドンキ|はま寿司|アオキ|バロー|業務|ようげん|あまのや|ubereats|uber eats|出前館|ディナー|夕食|朝食|夜ごはん|昼ごはん|飲み会|飲み|外食|ご飯|食事/.test(s)) return '食費';

    // 日用品
    if (/ドラッグ|薬局|クスリ|マツモトキヨシ|サンドラッグ|コスモス|ダイソー|カインズ|ホームセンター|ニトリ|コーナン|ドン・キホーテ|無印良品|ロフト|シャンプー|赤ちゃん本舗|買い物|買物|ショッピング/.test(s)) return '日用品';

    // 交通費
    if (/jr|suica|pasmo|鉄道|タクシー|ガソリン|駅|電車|バス|型タク|航空|空港|gas|eneos|出光|shell|コスモ石油|駐車|給油|ドライブ/.test(s)) return '交通費';

    // 娯楽
    if (/映画|シネマ|カラオケ|ゲーム|ボウリング|テーマパーク|遊園地|アミューズ|スポーツ|ジム|美術館|博物館|netflix|spotify|amazon prime|youtube|disney|ネットフリックス|書籍|本屋|旅行|ホテル|温泉|観光|遊び|デート|イベント|ライブ|コンサート/.test(s)) return '娯楽';

    // 医療
    if (/病院|クリニック|歯科|歯医者|薬|医院|調剤|診療|健康|整形|美容外科|美容皮膚|内科|小児科|眼科|耳鼻|皮膚科|検診|健診|通院/.test(s)) return '医療';

    // 衣服
    if (/ユニクロ|gu|ザラ|h&m|シマムラ|アオキ|服|アパレル|ファッション|abcマート|靴|シューズ/.test(s)) return '衣服';

    // 通信費
    if (/ソフトバンク|docomo|au|softbank|ラインモバイル|ocn|nuro|ビッグローブ|wi-?fi|通信/.test(s)) return '通信費';

    // 美容
    if (/美容院|美容室|ヘアサロン|サロン|ネイル|エステ|マッサージ|整体|カット|パーマ|ヘアカラー/.test(s)) return '日用品';

    // 家電・ガジェット
    if (/ヤマダ電機|ビックカメラ|ヨドバシ|ケーズ電器|apple|アップル|アマゾン|amazon/.test(s)) return '日用品';

    return '未分類';
}

/**
 * 🔍 Googleカレンダー・同時刻メールから店名を推定する
 * カード利用先が「Mastercard加盟店」等の汎用名の場合に呼ばれる
 */
function guessShopFromContext_(txDate) {
    try {
        // ① Googleカレンダーから推定（利用時刻の前後2時間のイベントを検索）
        const calStart = new Date(txDate.getTime() - 2 * 60 * 60 * 1000);
        const calEnd = new Date(txDate.getTime() + 2 * 60 * 60 * 1000);
        const cal = CalendarApp.getDefaultCalendar();
        const events = cal.getEvents(calStart, calEnd);
        for (const event of events) {
            const title = event.getTitle();
            // 「ランチ」「美容院」「病院」などイベント名ならそれを使う
            if (title && title.length > 1 && !/^(予定|TODO|タスク|リマインダー)$/i.test(title)) {
                return title;
            }
        }

        // ② 同時間帯のメールから推定（前後30分の受信メールで購入系を探す）
        const mailDate = Utilities.formatDate(txDate, 'Asia/Tokyo', 'yyyy/MM/dd');
        const searchQuery = `after:${mailDate} before:${mailDate} (subject:ご注文 OR subject:ご購入 OR subject:お買い上げ OR subject:レシート OR subject:お支払い OR subject:receipt OR subject:order)`;
        const threads = GmailApp.search(searchQuery, 0, 5);
        for (const thread of threads) {
            const msgs = thread.getMessages();
            for (const msg of msgs) {
                const msgTime = msg.getDate().getTime();
                // 利用時刻の前後1時間以内のメール
                if (Math.abs(msgTime - txDate.getTime()) < 60 * 60 * 1000) {
                    // 送信元のメール名からヒントを取得（例: "Amazon.co.jp", "楽天市場"）
                    const from = msg.getFrom();
                    const nameMatch = from.match(/"?([^"<]+)"?\s*</);
                    if (nameMatch) {
                        const senderName = nameMatch[1].trim();
                        if (senderName.length > 1 && !/info|noreply|no-reply|support|mail/.test(senderName.toLowerCase())) {
                            return senderName;
                        }
                    }
                    // 件名からヒントを取得
                    const subj = msg.getSubject();
                    if (subj.length > 2) {
                        return subj.substring(0, 30);
                    }
                }
            }
        }
    } catch (e) {
        console.log('コンテキスト推定エラー（スキップ）: ' + e.message);
    }
    return null; // 推定できず
}

/**
 * 📧 メール本文からカード利用情報を解析する
 * 対応: 三井住友カード（statement@vpass.ne.jp からの利用通知）
 * ※ PayPayカードは利用毎の通知メールが存在しないため、月次請求からの取り込みは対象外
 */
function parseCardEmail_(subject, body, sender) {
    // --- 三井住友カード ---
    // 送信元: statement@vpass.ne.jp
    // 件名: 「ご利用のお知らせ【三井住友カード】」
    // 本文例:
    //   ◇利用日：2026/02/21 17:14
    //   ◇利用先：Mastercard加盟店
    //   ◇利用取引：買物
    //   ◇利用金額：9,350円
    if (sender.includes('vpass.ne.jp') || sender.includes('smbc-card.com')) {
        // 「ご利用のお知らせ」以外のメール（キャンペーン、特典等）はスキップ
        if (!subject.includes('ご利用のお知らせ')) {
            return null;
        }

        const dateMatch = body.match(/◇利用日[：:]\s*(\d{4})\/(\d{1,2})\/(\d{1,2})\s*(\d{1,2}):(\d{2})?/);
        const amountMatch = body.match(/◇利用金額[：:]\s*(-?[\d,]+)円/);
        const shopMatch = body.match(/◇利用先[：:]\s*([^\n\r]+)/);

        if (dateMatch && amountMatch) {
            const dateStr = `${dateMatch[1]}/${String(dateMatch[2]).padStart(2, '0')}/${String(dateMatch[3]).padStart(2, '0')}`;
            const rawAmount = parseInt(amountMatch[1].replace(/,/g, ''));
            const isRefund = rawAmount < 0;
            const amount = Math.abs(rawAmount);
            let shop = shopMatch ? shopMatch[1].trim() : '三井住友カード利用';
            let hintCategory = null;

            // 「Mastercard加盟店」等の汎用名の場合、カレンダー・メールから推定を試みる
            if (/加盟店|Mastercard|Visa|JCB/.test(shop)) {
                const txHour = dateMatch[4] ? parseInt(dateMatch[4]) : 12;
                const txDate = new Date(parseInt(dateMatch[1]), parseInt(dateMatch[2]) - 1, parseInt(dateMatch[3]), txHour, dateMatch[5] ? parseInt(dateMatch[5]) : 0);
                const hint = guessShopFromContext_(txDate);
                if (hint) {
                    shop = hint + '（推定）';
                    // 推定名からカテゴリも判定する
                    hintCategory = guessCategory_(hint);
                }
            }

            return {
                date: dateStr,
                amount: amount,
                memo: isRefund ? `【返金】${shop}` : shop,
                method: '三井住友カード',
                category: isRefund ? '返金' : (hintCategory && hintCategory !== '未分類') ? hintCategory : guessCategory_(shop),
                type: isRefund ? '収入' : '支出'
            };
        }
    }

    // --- PayPayカード 利用速報 ---
    // 送信元: paypaycard-info@mail.paypay-card.co.jp
    // 件名: 「PayPayカード ゴールド（Visa）利用速報」
    // 本文例: 「PayPayカード ゴールド（Visa）利用速報  ソフトバンク(B) 2026年2月5日 22:53 4,733円」
    if (sender.includes('paypay-card.co.jp') && subject.includes('利用速報')) {
        // 本文フォーマット: 「利用速報  店名 日付 時刻 金額円」
        const match = body.match(/利用速報\s+(.+?)\s+(\d{4})年(\d{1,2})月(\d{1,2})日\s+\d{1,2}:\d{2}\s+([\d,]+)円/);
        if (match) {
            const shop = match[1].trim();
            const dateStr = `${match[2]}/${String(match[3]).padStart(2, '0')}/${String(match[4]).padStart(2, '0')}`;
            const amount = parseInt(match[5].replace(/,/g, ''));
            return { date: dateStr, amount: amount, memo: shop, method: 'PayPayカード', category: guessCategory_(shop), type: '支出' };
        }
    }

    return null; // 解析失敗（対象外メール）
}

/**
 * ✅ 解析済みレコードをスプレッドシートへ書き込む（重複チェック付き）
 */
function writeCardRecord_(sheet, record) {
    // 重複チェック: 同じ日付+金額+摘要の組み合わせが既に存在する場合はスキップ
    const lastRow = Math.max(sheet.getLastRow(), 1);
    if (lastRow > 1) {
        const existingData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
        for (const row of existingData) {
            const existingDate = row[0] instanceof Date
                ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd')
                : String(row[0]).substring(0, 10);
            const existingAmount = Number(row[1]);
            const existingMemo = String(row[3]);
            if (existingDate === record.date && existingAmount === record.amount && existingMemo === record.memo) {
                return false; // 重複のためスキップ
            }
        }
    }
    sheet.appendRow([
        record.date,
        record.amount,
        record.category,
        record.memo,
        record.type,
        record.method,
        '', // IsFixed
        ''  // Account (口座は後から設定可)
    ]);
    return true; // 書き込み成功
}

/**
 * 🔄 最新のカード利用メールを処理する（定期実行トリガー用）
 * GASのタイマーに設定: dailyFetchCardEmails を「毎日1回」などで実行する
 */
function dailyFetchCardEmails() {
    fetchCardEmailsByQuery_('newer_than:2d'); // 直近2日以内のメールを処理
}

/**
 * 📅 指定期間の過去メールを一括取り込み（初回のみ手動実行）
 * 使い方: GASエディタから fetchPastCardEmails() を実行してください
 */
function fetchPastCardEmails() {
    // 2026年1月1日以降のメールを取り込む
    fetchCardEmailsByQuery_('after:2026/01/01');
}

/**
 * 内部処理: Gmailクエリを実行してカードメールを取得・解析する
 */
function fetchCardEmailsByQuery_(query) {
    if (!SPREADSHEET_ID) {
        console.log('SPREADSHEET_ID が未設定です');
        return;
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('家計簿');
    if (!sheet) {
        console.log('家計簿シートが見つかりません');
        return;
    }

    // 三井住友カードの利用通知 + PayPayカードの利用速報メールを検索
    // ※2つのクエリを分けて検索し、結果を合算する（GmailのOR構文の問題を回避）
    const smbcQuery = `from:statement@vpass.ne.jp subject:"ご利用のお知らせ" ${query}`;
    const paypayQuery = `from:paypay-card.co.jp subject:"利用速報" ${query}`;
    const smbcThreads = GmailApp.search(smbcQuery, 0, 200);
    const paypayThreads = GmailApp.search(paypayQuery, 0, 200);
    const threads = smbcThreads.concat(paypayThreads);

    let writtenCount = 0;
    let skippedCount = 0;

    for (const thread of threads) {
        const messages = thread.getMessages();
        for (const message of messages) {
            const sender = message.getFrom();
            const subject = message.getSubject();
            const body = message.getPlainBody();
            const record = parseCardEmail_(subject, body, sender);
            if (record) {
                const written = writeCardRecord_(sheet, record);
                if (written) {
                    writtenCount++;
                } else {
                    skippedCount++;
                }
            }
        }
    }

    console.log(`✅ Gmail連携完了: ${writtenCount}件追記, ${skippedCount}件スキップ（重複）`);
    return { written: writtenCount, skipped: skippedCount };
}

/**
 * 🔍 【診断用】Gmailのカード関連メールを幅広く検索して情報を表示する
 * GASエディタから実行して、実行ログで送信元・件名・本文冒頭を確認してください
 */
function debugSearchCardEmails() {
    // 広い条件で検索（三井住友 or PayPay に関連しそうなメール）
    var queries = [
        'from:smbc subject:カード after:2026/01/01',
        'from:vpass after:2026/01/01',
        'from:paypay subject:カード after:2026/01/01',
        'subject:三井住友 subject:利用 after:2026/01/01',
        'subject:PayPayカード after:2026/01/01',
        'subject:ご利用 (三井住友 OR paypay OR smbc) after:2026/01/01',
        'from:smbc-card after:2026/01/01',
        'from:paypay-card after:2026/01/01',
        'subject:利用速報 after:2026/01/01'
    ];

    var found = 0;
    for (var q = 0; q < queries.length; q++) {
        var threads = GmailApp.search(queries[q], 0, 5);
        if (threads.length > 0) {
            console.log('=== クエリ: ' + queries[q] + ' → ' + threads.length + '件 ===');
            for (var t = 0; t < threads.length; t++) {
                var msgs = threads[t].getMessages();
                var msg = msgs[0];
                console.log('  📧 件名: ' + msg.getSubject());
                console.log('  📬 送信元: ' + msg.getFrom());
                console.log('  📅 日付: ' + msg.getDate());
                var bodySnippet = msg.getPlainBody().substring(0, 300).replace(/\n/g, ' ');
                console.log('  📝 本文(先頭300文字): ' + bodySnippet);
                console.log('  ---');
                found++;
            }
        }
    }

    if (found === 0) {
        console.log('⚠️ どのクエリでもカード関連メールが見つかりませんでした。');
        console.log('💡 Gmailで「三井住友」「PayPay」で検索して、実際のメールがあるか確認してください。');
        console.log('💡 GASが紐づいているGmailアカウントが、カード通知を受信しているアカウントと同じか確認してください。');
    } else {
        console.log('✅ 合計 ' + found + ' 件のメールが見つかりました。上記の送信元と件名をもとにパーサーを調整します。');
    }
}
