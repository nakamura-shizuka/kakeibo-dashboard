/**
 * 🛠️ ユーティリティ
 * =========================================================
 * 共通ヘルパー関数群
 */

/**
 * JSON レスポンスを生成する
 */
function createJsonResponse(data) {
    return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}

/**
 * LINEのテスト用トークンかどうかを判定する
 */
function isTestToken(token) {
    return ['00000000000000000000000000000000', 'ffffffffffffffffffffffffffffffff', '11111111111111111111111111111111'].indexOf(token) !== -1;
}

/**
 * エラーをスプレッドシートに記録する
 */
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
    } catch (ignored) { /* ログ書き込み失敗は無視（循環呼び出し防止） */ }
}

// ===== 入力バリデーション =====

/**
 * 金額のバリデーション
 * @param {*} amount - 入力値
 * @returns {{ valid: boolean, value: number, message?: string }}
 */
function validateAmount(amount) {
    const num = parseInt(String(amount).replace(/[,，]/g, ''), 10);
    if (isNaN(num) || num <= 0) {
        return { valid: false, value: 0, message: '金額は1以上の整数で入力してください' };
    }
    return { valid: true, value: num };
}

/**
 * 日付文字列のバリデーション（yyyy/MM/dd 形式）
 * @param {string} dateStr - 日付文字列
 * @returns {{ valid: boolean, message?: string }}
 */
function validateDate(dateStr) {
    if (!dateStr) return { valid: false, message: '日付を入力してください' };
    if (!/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(String(dateStr).trim())) {
        return { valid: false, message: '日付形式が不正です（yyyy/MM/dd）' };
    }
    return { valid: true };
}

/**
 * 行番号のバリデーション
 * @param {*} rowIndex - スプレッドシートの行番号（2以上）
 * @returns {{ valid: boolean, message?: string }}
 */
function validateRowIndex(rowIndex) {
    const n = Number(rowIndex);
    if (!Number.isInteger(n) || n < 2) {
        return { valid: false, message: '行番号が不正です（rowIndex=' + rowIndex + '）' };
    }
    return { valid: true };
}

/**
 * HtmlService テンプレートのインクルードヘルパー
 * index.html 内で <?!= include('styles'); ?> のように使用する
 * createTemplateFromFile を使うことで、インクルード先の <?= ?> 式も評価される
 */
function include(filename) {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

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
