/**
 * 🚪 エントリポイント
 * =========================================================
 * doGet / doPost - GAS Web App のリクエストハンドラ
 */

/**
 * LINEからのWebhookを受け取るエンドポイント
 */
function doPost(e) {
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

    // ===== LINE Webhook 処理 =====
    try {
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

    // ダッシュボードHTML配信（テンプレートエンジン経由）
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('みえる化家計簿')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}
