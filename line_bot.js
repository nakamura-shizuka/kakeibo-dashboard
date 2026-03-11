/**
 * 💬 LINE Bot
 * =========================================================
 * LINE Webhook 処理・メッセージ解析・返信・プッシュ通知
 */

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
 * 送信元のLINEユーザーIDを ScriptProperties に保存する（Push送信用）
 */
function saveUserIdToSettings(userId) {
    if (!userId) return;
    try {
        const currentId = (PROPERTIES.getProperty('LINE_USER_ID') || "").trim();
        if (currentId !== userId) {
            PROPERTIES.setProperty('LINE_USER_ID', userId);
        }
    } catch (e) {
        console.warn('LINE_USER_ID保存失敗:', e.message);
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

/**
 * 🔑 LINE_USER_IDを取得するヘルパー（ScriptProperties を参照）
 */
function getLineUserId_() {
    const propId = (PROPERTIES.getProperty('LINE_USER_ID') || "").trim();
    if (propId) return propId;

    // フォールバック: 旧バージョンで設定シートG3に保存されていた場合
    try {
        if (SPREADSHEET_ID) {
            const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
            const settingsSheet = ss.getSheetByName('設定');
            if (settingsSheet) {
                const userId = settingsSheet.getRange('G3').getValue();
                if (userId) {
                    const id = userId.toString().trim();
                    PROPERTIES.setProperty('LINE_USER_ID', id);
                    return id;
                }
            }
        }
    } catch (e) {
        console.warn("設定シートからLINE_USER_ID取得失敗:", e.message);
    }

    return null;
}
