// ===== Layer 1: 月次確定モジュール =====
// カード請求確定メールの取込と口座残高の記録を担当する。
// 支出総額のKPIは常に「月次確定シート＋家計簿シートの非カード明細」から算出し、
// カード利用明細（リアルタイム速報）はカテゴリ内訳の参考情報に留める。
//
// 確認済みの実メール形式（2026年7月時点）:
// - PayPayカード: paypaycard-info@mail.paypay-card.co.jp
//   件名「6月の請求金額のお知らせ」/ 本文「請求金額：34,027円（確定）」（毎月22日頃）
//   ※「請求予定金額のお知らせ」（仮確定・13日頃）は取込対象外
// - イオンカード: statement@email.aeon.co.jp
//   件名「3月ご請求額のお知らせ」/ 本文「ご請求額 ： 207円」（毎月17日頃・翌月支払分）
//   ※イオンウォレット(AEON Pay)利用者にはメールが配信されないため、
//     LINE「確定 イオン 45000」での手入力フォールバックあり
// - 三井住友カード: statement@vpass.ne.jp
//   件名「お支払い金額のお知らせ」（毎月26日頃・翌月10日支払分）
//   ※Vpass設定で金額表示をオンにしていない場合、本文に金額が含まれない。
//     その場合は金額0円のプレースホルダ行を作成してLINEで手入力を促す。

/**
 * 🗂️ Layer 1 用シート（月次確定・残高）を初期化して返す
 */
function ensureLayer1Sheets_() {
    if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID未設定');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let confirmedSheet = ss.getSheetByName(SHEET_MONTHLY_CONFIRMED);
    if (!confirmedSheet) {
        confirmedSheet = ss.insertSheet(SHEET_MONTHLY_CONFIRMED);
        confirmedSheet.appendRow(['BillingMonth', 'Card', 'Amount', 'ConfirmedAt', 'Source', 'Note']);
        confirmedSheet.getRange(1, 1, 1, 6).setBackground('#e0f7fa').setFontWeight('bold');
    }

    let balanceSheet = ss.getSheetByName(SHEET_BALANCE);
    if (!balanceSheet) {
        balanceSheet = ss.insertSheet(SHEET_BALANCE);
        balanceSheet.appendRow(['Date', 'Account', 'Balance', 'Source']);
        balanceSheet.getRange(1, 1, 1, 4).setBackground('#e0f7fa').setFontWeight('bold');
    }

    return { confirmedSheet: confirmedSheet, balanceSheet: balanceSheet };
}

/**
 * 📧 請求確定メールを解析する
 * @param {string} subject - 件名
 * @param {string} body - 本文（プレーンテキスト）
 * @param {string} sender - 送信元
 * @param {Date} msgDate - 受信日時
 * @returns {Object|null} { billingMonth: 'yyyy/MM', card, amount, note } 解析対象外なら null
 *                        金額が本文にない場合は amount: null を返す
 */
function parseBillingEmail_(subject, body, sender, msgDate) {
    // --- PayPayカード（確定のみ。「請求予定金額」=仮確定は対象外） ---
    if (sender.includes('paypay-card.co.jp') && subject.includes('請求金額のお知らせ') && !subject.includes('予定')) {
        const monthMatch = body.match(/(\d{4})年(\d{1,2})月の請求金額のお知らせ/);
        const amountMatch = body.match(/請求金額[：:]\s*([\d,]+)円（確定）/) || body.match(/請求金額[：:]\s*([\d,]+)円/);
        if (monthMatch && amountMatch) {
            return {
                billingMonth: monthMatch[1] + '/' + String(monthMatch[2]).padStart(2, '0'),
                card: 'PayPay',
                amount: parseInt(amountMatch[1].replace(/,/g, ''), 10),
                note: ''
            };
        }
        return null;
    }

    // --- イオンカード ---
    if (sender.includes('aeon.co.jp') && /(\d{1,2})月ご請求額のお知らせ/.test(subject)) {
        const monthMatch = subject.match(/(\d{1,2})月ご請求額のお知らせ/);
        const amountMatch = body.match(/ご請求額\s*[：:]\s*([\d,]+)円/);
        // 支払月の年は「お支払日 ： 2026年3月2日」から取得（12月受信→1月請求の年跨ぎ対応）
        const payDateMatch = body.match(/お支払日\s*[：:]\s*(\d{4})年/);
        if (monthMatch && amountMatch) {
            const month = parseInt(monthMatch[1], 10);
            let year = payDateMatch ? parseInt(payDateMatch[1], 10) : msgDate.getFullYear();
            if (!payDateMatch && month < msgDate.getMonth() + 1) year += 1; // 12月受信で1月請求など
            return {
                billingMonth: year + '/' + String(month).padStart(2, '0'),
                card: 'イオン',
                amount: parseInt(amountMatch[1].replace(/,/g, ''), 10),
                note: ''
            };
        }
        return null;
    }

    // --- 三井住友カード（お支払い金額のお知らせ） ---
    if (sender.includes('vpass.ne.jp') && subject.includes('お支払い金額のお知らせ')) {
        // 支払月: 毎月26日頃に届き翌月10日支払のため、受信月+1を支払月とする
        const payDate = new Date(msgDate.getFullYear(), msgDate.getMonth() + 1, 1);
        const billingMonth = Utilities.formatDate(payDate, 'Asia/Tokyo', 'yyyy/MM');
        // Vpass設定で金額表示がオンの場合に備えて複数パターンを試す
        const amountMatch = body.match(/お支払い合計額\s*[：:]?\s*([\d,]+)\s*円/) ||
            body.match(/お支払い金額\s*[：:]\s*([\d,]+)\s*円/) ||
            body.match(/ご請求金額\s*[：:]\s*([\d,]+)\s*円/);
        return {
            billingMonth: billingMonth,
            card: '三井住友',
            amount: amountMatch ? parseInt(amountMatch[1].replace(/,/g, ''), 10) : null,
            note: amountMatch ? '' : '⚠️金額未取得（メール本文に金額なし）'
        };
    }

    return null;
}

/**
 * ✅ 請求確定レコードを月次確定シートへ書き込む
 * 重複キー: BillingMonth + Card。
 * - 既存と同額: スキップ
 * - 金額が異なる: 上書き（請求額の訂正メール・手入力修正に対応）
 * - ただし手入力（manual）の値をメール（mail）が上書きすることはしない
 * @param {Object} record - { billingMonth, card, amount, source, note }
 * @returns {string} 'written' | 'updated' | 'skipped'
 */
function writeBillingRecord_(record) {
    const sheets = ensureLayer1Sheets_();
    const sheet = sheets.confirmedSheet;
    const amount = record.amount === null || record.amount === undefined ? 0 : Number(record.amount);
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
        for (let i = 0; i < data.length; i++) {
            const rowMonth = data[i][0] instanceof Date
                ? Utilities.formatDate(data[i][0], 'Asia/Tokyo', 'yyyy/MM')
                : String(data[i][0]).trim();
            if (rowMonth === record.billingMonth && String(data[i][1]).trim() === record.card) {
                const existingAmount = Number(data[i][2]) || 0;
                const existingSource = String(data[i][4]).trim();
                if (existingAmount === amount) return 'skipped';
                // 手入力値はメール取込で上書きしない（金額未取得プレースホルダの0円は除く）
                if (existingSource === 'manual' && record.source === 'mail' && existingAmount > 0) return 'skipped';
                // メールに金額がないプレースホルダで既存行を潰さない
                if (record.source === 'mail' && amount === 0 && existingAmount > 0) return 'skipped';
                sheet.getRange(i + 2, 3, 1, 4).setValues([[amount, now, record.source, record.note || '']]);
                invalidateConfirmedCache_(record.billingMonth);
                return 'updated';
            }
        }
    }

    sheet.appendRow([record.billingMonth, record.card, amount, now, record.source, record.note || '']);
    invalidateConfirmedCache_(record.billingMonth);
    return 'written';
}

/**
 * 対象支払月のダッシュボードキャッシュを無効化する
 * @param {string} billingMonth - 'yyyy/MM'
 */
function invalidateConfirmedCache_(billingMonth) {
    try {
        const parts = billingMonth.split('/');
        invalidateDashboardCache(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1);
    } catch (e) { /* キャッシュ無効化失敗は無視 */ }
}

/**
 * 🔄 請求確定メールの日次取込（トリガー用）
 */
function fetchBillingEmails() {
    fetchBillingEmailsByQuery_('newer_than:7d');
}

/**
 * 📅 過去の請求確定メールを一括取込（初回のみGASエディタから手動実行）
 */
function fetchPastBillingEmails() {
    return fetchBillingEmailsByQuery_('after:2025/07/01');
}

/**
 * 内部処理: Gmailクエリで請求確定メールを取得・解析・記録する
 * 取込結果と金額未取得はLINEで通知する（沈黙して欠測させない）
 * @param {string} query - Gmail検索クエリ（期間指定等）
 */
function fetchBillingEmailsByQuery_(query) {
    if (!SPREADSHEET_ID) {
        console.warn('SPREADSHEET_ID が未設定です');
        return;
    }

    const searchQueries = [
        'from:paypay-card.co.jp subject:"請求金額のお知らせ" ' + query,
        'from:statement@email.aeon.co.jp subject:"ご請求額のお知らせ" ' + query,
        'from:statement@vpass.ne.jp subject:"お支払い金額のお知らせ" ' + query
    ];

    let writtenItems = [];
    let needsAmountItems = [];
    let skipped = 0;

    for (const q of searchQueries) {
        const threads = GmailApp.search(q, 0, 100);
        for (const thread of threads) {
            for (const message of thread.getMessages()) {
                const record = parseBillingEmail_(
                    message.getSubject(),
                    message.getPlainBody(),
                    message.getFrom(),
                    message.getDate()
                );
                if (!record) continue;
                record.source = 'mail';
                const result = writeBillingRecord_(record);
                if (result === 'skipped') {
                    skipped++;
                } else if (record.amount === null) {
                    needsAmountItems.push(record);
                } else {
                    writtenItems.push(record);
                }
            }
        }
    }

    console.log('✅ 請求確定メール取込: ' + writtenItems.length + '件記録, ' +
        needsAmountItems.length + '件金額未取得, ' + skipped + '件スキップ（重複）');

    // LINE通知
    const userId = getLineUserId_();
    if (userId) {
        if (writtenItems.length > 0) {
            const lines = writtenItems.map(r => '・' + r.card + 'カード ' + r.billingMonth + '支払分: ' +
                Number(r.amount).toLocaleString() + '円');
            pushLineMessage(userId, '💳 【請求確定額を記録しました】\n\n' + lines.join('\n'));
        }
        for (const r of needsAmountItems) {
            pushLineMessage(userId,
                '⚠️ 三井住友カードの「お支払い金額のお知らせ」が届きましたが、メール本文に金額が含まれていません。\n\n' +
                '対応方法（どちらか）:\n' +
                '① Vpassの設定で「お支払い金額の確定メール表示内容」をオンにする（次回から自動記録されます）\n' +
                'https://www.smbc-card.com/mem/update/vp_henkou.jsp\n\n' +
                '② このLINEに「確定 三井住友 45000」の形式で金額を送る（' + r.billingMonth + '支払分として記録）');
        }
    }

    return { written: writtenItems.length, needsAmount: needsAmountItems.length, skipped: skipped };
}

/**
 * 💰 口座残高を記録する（LINE Bot・ダッシュボードから呼ばれる）
 * 同日・同口座の既存記録は上書きする
 * @param {string} account - 口座名
 * @param {number} balance - 残高
 * @param {string} source - 'LINE' | 'dashboard'
 * @returns {Object} { balance, prevMonthTotal, currentTotal, diff } diffは前月末比（データ不足ならnull）
 */
function recordBalance_(account, balance, source) {
    const sheets = ensureLayer1Sheets_();
    const sheet = sheets.balanceSheet;
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

    let updated = false;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
        for (let i = 0; i < data.length; i++) {
            const rowDate = data[i][0] instanceof Date
                ? Utilities.formatDate(data[i][0], 'Asia/Tokyo', 'yyyy/MM/dd')
                : String(data[i][0]).substring(0, 10);
            if (rowDate === today && String(data[i][1]).trim() === account) {
                sheet.getRange(i + 2, 3, 1, 2).setValues([[balance, source]]);
                updated = true;
                break;
            }
        }
    }
    if (!updated) {
        sheet.appendRow([today, account, balance, source]);
    }

    // 前月末時点の同口座残高との差分を計算（入力の動機付け用）
    const now = new Date();
    const prevBalances = getBalancesAsOf_(new Date(now.getFullYear(), now.getMonth(), 0)); // 前月末
    const prev = prevBalances.byAccount[account];
    const diff = (prev !== undefined) ? balance - prev : null;

    try { invalidateDashboardCache(now.getFullYear(), now.getMonth()); } catch (e) { /* 無視 */ }
    return { balance: balance, diff: diff };
}

/**
 * 📊 指定日時点の各口座の最新残高を取得する
 * 各口座について asOf 以前の最新記録を採用する
 * @param {Date} asOf - 基準日（この日以前の記録を対象）
 * @returns {Object} { byAccount: {口座名: 残高}, total, latestDate: 'yyyy/MM/dd'|null }
 */
function getBalancesAsOf_(asOf) {
    const result = { byAccount: {}, total: 0, latestDate: null };
    if (!SPREADSHEET_ID) return result;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_BALANCE);
    if (!sheet || sheet.getLastRow() <= 1) return result;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    const latestByAccount = {}; // 口座名 -> { time, balance }

    data.forEach(function (row) {
        if (!row[0] || !row[1]) return;
        const d = row[0] instanceof Date ? row[0] : new Date(String(row[0]));
        if (isNaN(d.getTime()) || d.getTime() > asOf.getTime()) return;
        const account = String(row[1]).trim();
        if (!latestByAccount[account] || d.getTime() >= latestByAccount[account].time) {
            latestByAccount[account] = { time: d.getTime(), balance: Number(row[2]) || 0 };
        }
    });

    let latestTime = 0;
    Object.keys(latestByAccount).forEach(function (account) {
        result.byAccount[account] = latestByAccount[account].balance;
        result.total += latestByAccount[account].balance;
        if (latestByAccount[account].time > latestTime) latestTime = latestByAccount[account].time;
    });
    if (latestTime > 0) {
        result.latestDate = Utilities.formatDate(new Date(latestTime), 'Asia/Tokyo', 'yyyy/MM/dd');
    }
    return result;
}

/**
 * 🔍 【診断用】請求関連メールを幅広く検索して実形式をログ表示する
 * メール形式が変わった疑いがある時にGASエディタから実行する
 */
function debugSearchBillingEmails() {
    const queries = [
        'from:paypay-card.co.jp subject:請求 newer_than:120d',
        'from:statement@email.aeon.co.jp newer_than:365d',
        'from:aeon.co.jp subject:請求 newer_than:365d',
        'from:statement@vpass.ne.jp subject:お支払い newer_than:120d'
    ];
    for (const q of queries) {
        const threads = GmailApp.search(q, 0, 3);
        console.log('=== クエリ: ' + q + ' → ' + threads.length + '件 ===');
        for (const thread of threads) {
            const msg = thread.getMessages()[0];
            console.log('  📧 件名: ' + msg.getSubject());
            console.log('  📬 送信元: ' + msg.getFrom());
            console.log('  📅 日付: ' + msg.getDate());
            console.log('  📝 本文(先頭500文字): ' + msg.getPlainBody().substring(0, 500).replace(/\n/g, ' '));
            console.log('  ---');
        }
    }
}
