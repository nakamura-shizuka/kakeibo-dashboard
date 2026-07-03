/**
 * 💬 LINE Bot
 * =========================================================
 * LINE Webhook 処理・メッセージ解析・返信・プッシュ通知
 */

/**
 * 全角数字・カンマを除去して整数に変換する
 */
function parseAmountText_(text) {
    const normalized = text
        .replace(/[,，]/g, "")
        .replace(/[０-９]/g, function (s) {
            return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
        });
    return parseInt(normalized, 10);
}

/**
 * メッセージ解析 → スプレッドシート記録
 */
function processMessage(userMessage) {
    // --- 残高コマンド: 「残高 ゆうちょ 1234567」 ---
    const balanceMatch = userMessage.match(/^残高[\s　]+(\S+)[\s　]+([0-9０-９,，]+)円?$/);
    if (balanceMatch) {
        const account = balanceMatch[1].trim();
        const balance = parseAmountText_(balanceMatch[2]);
        if (isNaN(balance) || balance < 0) return "❌ 残高を正しく読み取れませんでした。";
        try {
            const result = recordBalance_(account, balance, 'LINE');
            let reply = `🏦 ${account} の残高 ¥${balance.toLocaleString()} を記録しました！`;
            if (result.diff !== null) {
                const sign = result.diff >= 0 ? '+' : '';
                reply += `\n（前月末比 ${sign}${result.diff.toLocaleString()}円）`;
            }
            return reply;
        } catch (error) {
            console.error("残高記録エラー:", error);
            return `❌ 残高の記録に失敗しました: ${error.message}`;
        }
    }

    // --- 「残高」単体: 使い方ガイド ---
    if (/^残高$/.test(userMessage.trim())) {
        let accountNames = [];
        try {
            const settings = getSettingsData();
            if (settings.accounts && settings.accounts.length > 0) {
                accountNames = settings.accounts.map(a => a.name);
            }
        } catch (e) { /* 設定取得失敗時はガイドのみ */ }
        let guide = "🏦 残高の記録方法\n\n「残高 口座名 金額」の形式で送ってね！\n\n✅ 例：\n・残高 ゆうちょ 1234567";
        if (accountNames.length > 0) {
            guide += "\n\n📋 登録済みの口座:\n" + accountNames.map(n => `・${n}`).join('\n');
        }
        return guide;
    }

    // --- 確定コマンド: 「確定 イオン 45000」（請求確定額の手入力フォールバック） ---
    const confirmedMatch = userMessage.match(/^確定[\s　]+(\S+?)[\s　]+(?:(\d{1,2})月[\s　]+)?([0-9０-９,，]+)円?$/);
    if (confirmedMatch) {
        const cardInput = confirmedMatch[1].trim();
        const card = BILLING_CARDS.find(c => cardInput.includes(c) || c.includes(cardInput));
        if (!card) {
            return "❌ カード名を認識できませんでした。\n「確定 イオン 45000」のように、" + BILLING_CARDS.join('・') + " のいずれかで送ってね。";
        }
        const amount = parseAmountText_(confirmedMatch[3]);
        if (isNaN(amount) || amount < 0) return "❌ 金額を正しく読み取れませんでした。";
        const now = new Date();
        const month = confirmedMatch[2] ? parseInt(confirmedMatch[2], 10) : now.getMonth() + 1;
        let year = now.getFullYear();
        if (confirmedMatch[2] && month > now.getMonth() + 1 + 6) year -= 1; // 半年以上先の月指定は前年扱い
        const billingMonth = year + '/' + String(month).padStart(2, '0');
        try {
            const result = writeBillingRecord_({ billingMonth: billingMonth, card: card, amount: amount, source: 'manual', note: 'LINE手入力' });
            const action = result === 'updated' ? '更新' : '記録';
            return `💳 ${card}カードの ${billingMonth} 支払分 ¥${amount.toLocaleString()} を${action}しました！`;
        } catch (error) {
            console.error("確定額記録エラー:", error);
            return `❌ 確定額の記録に失敗しました: ${error.message}`;
        }
    }

    const match = userMessage.match(/^(.+?)[\s　]+([0-9０-９,，]+)円?$/);

    if (!match) {
        return "📝 使い方ガイド\n\n「品名 金額」の形式で送ってね！\n\n✅ 例：\n・ランチ 1200\n・コンビニ 350\n・電車代 500\n\n🏦 残高の記録: 「残高 ゆうちょ 1234567」\n💳 請求確定額: 「確定 イオン 45000」";
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
        try {
            const now = new Date();
            invalidateDashboardCache(now.getFullYear(), now.getMonth());
        } catch (e) { /* キャッシュ無効化失敗は無視 */ }
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
