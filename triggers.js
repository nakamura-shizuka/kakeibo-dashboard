// ===== トリガー・自動処理モジュール =====
// タイムドリブントリガーで定期実行される関数群

/**
 * ⏰ 定期実行トリガー用：週次レポート送信（毎週土曜日の夕方などを想定）
 */
function sendWeeklyReport() {
    const userId = getLineUserId_();
    if (!userId) {
        console.warn("LINE_USER_IDが設定されていないため、通知をスキップしました。\n対処法: (1) LINEからBotへ一度メッセージを送る、または (2) GASのスクリプトプロパティに LINE_USER_ID を手動設定してください。");
        return;
    }

    const analysisResult = generateAiAnalysis(true); // 週次
    const message = "📊 【みえる化家計簿】週次データ分析レポート\n\n" + analysisResult;

    pushLineMessage(userId, message);
    console.log("週次レポートをLINEに送信しました (userId: " + userId.substring(0, 8) + "...)");
}

/**
 * ⏰ 定期実行トリガー用：月次レポート送信（毎月1日の朝などを想定）
 */
function sendMonthlyReport() {
    const userId = getLineUserId_();
    if (!userId) {
        console.warn("LINE_USER_IDが設定されていないため、月次通知をスキップしました。");
        return;
    }

    const analysisResult = generateAiAnalysis(false); // 月次
    const message = "📈 【みえる化家計簿】月次データ分析レポート\n\n" + analysisResult;

    pushLineMessage(userId, message);
    console.log("月次レポートをLINEに送信しました");
}

/**
 * ⏰ 週次レポート用のトリガーを作成する (毎週土曜日の夕方 18:00頃)
 */
function setupWeeklyTrigger() {
    // 既存の同名トリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'sendWeeklyReport') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // 毎週土曜日の18時頃に実行
    ScriptApp.newTrigger('sendWeeklyReport')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.SATURDAY)
        .atHour(18)
        .create();

    console.log("週次レポート(sendWeeklyReport)のトリガーを土曜日18時台に設定しました。");
}

/**
 * ⏰ 月次レポート用のトリガーを作成する (毎月1日の朝 08:00頃)
 */
function setupMonthlyTrigger() {
    // 既存の同名トリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'sendMonthlyReport') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // 毎月1日の8時頃に実行
    ScriptApp.newTrigger('sendMonthlyReport')
        .timeBased()
        .onMonthDay(1)
        .atHour(8)
        .create();

    console.log("月次レポート(sendMonthlyReport)のトリガーを毎月1日の8時台に設定しました。");
}

/**
 * 🚀 AI分析用の定期トリガーを一括で設定するマスター関数
 * 初回セットアップ時や、トリガーを作り直したい時にエディタから手動実行してください。
 */
function setupAITriggers() {
    setupWeeklyTrigger();
    setupMonthlyTrigger();
    console.log("AI分析用の定期トリガー(週次/月次)のセットアップが完了しました。");
}

/**
 * ⏰ 日次トリガーを設定する
 * - autoRecordFixedExpenses: 毎日06時（固定費自動記録）
 * - checkBudgetAndAlert: 毎日08時（予算アラート）
 * - dailyFetchCardEmails: 1時間ごと（Gmail カード明細取込）
 */
function setupDailyTrigger() {
    const dailyFunctions = ['autoRecordFixedExpenses', 'checkBudgetAndAlert', 'dailyFetchCardEmails'];

    // 既存の同名トリガーを先に削除
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (dailyFunctions.indexOf(trigger.getHandlerFunction()) >= 0) {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    ScriptApp.newTrigger('autoRecordFixedExpenses').timeBased().everyDays(1).atHour(6).create();
    ScriptApp.newTrigger('dailyFetchCardEmails').timeBased().everyHours(1).create(); // 毎時取込（当日中に反映）
    ScriptApp.newTrigger('checkBudgetAndAlert').timeBased().everyDays(1).atHour(8).create();

    console.log("トリガー設定完了（固定費記録: 毎日6時, Gmail取込: 毎時, 予算アラート: 毎日8時）");
}

/**
 * 🚀 全トリガーを一括セットアップするマスター関数
 * 初回セットアップ時にGASエディタから手動実行してください。
 */
function setupAllTriggers() {
    setupDailyTrigger();
    setupAITriggers();
    console.log("✅ 全トリガーのセットアップが完了しました（日次/週次/月次）。");
}

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
            } catch (e) { console.warn('固定費JSONパース失敗:', e.message); }
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
        if (targetDate === todayDate) return true;
        // 月末日の場合、設定日が今日より大きければ強制的に今日記録する（例: 31日設定で今月が30日までの場合）
        if (isEndOfMonth && targetDate > todayDate) return true;
        return false;
    });

    if (targets.length === 0) return;

    // 今月の既存家計簿データを取得（重複防止用）
    const lastRow = kakeiboSheet.getLastRow();
    const recordedKeys = new Set();
    const targetUserId = getLineUserId_() || "";

    if (lastRow > 1) {
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

    const targetUserId = getLineUserId_();
    if (!targetUserId) {
        console.warn("LINE USER ID が未設定のためアラートをスキップします。");
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

    // 設定シートの K列/L列 にアラートフラグを保持
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
        settingsSheet.getRange('L3').setValue(true);
        return; // 100%を送るなら80%は同時に送らない
    }

    // 80% 超過チェック
    if (percent >= 80 && percent < 100 && !is80Sent) {
        const msg = `⚠️ 【予算アラート】\n\n今月の支出が予算の80%を超えました。\n残り: ${(budget - spent).toLocaleString()}円\n\n月末まで少し節約を意識してみましょう👀`;
        pushLineMessage(targetUserId, msg);
        settingsSheet.getRange('L2').setValue(true);
    }
}
