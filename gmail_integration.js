// ===== Gmail自動連携モジュール =====
// クレジットカード利用メールを解析してスプレッドシートに自動記録する

/**
 * 🏷️ 店名からカテゴリを自動推定する
 * @param {string} shopName - 店名
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
 * @param {Date} txDate - 利用日時
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
        console.warn('コンテキスト推定エラー（スキップ）: ' + e.message);
    }
    return null; // 推定できず
}

/**
 * 📧 メール本文からカード利用情報を解析する
 * 対応: 三井住友カード（statement@vpass.ne.jp からの利用通知）
 *       PayPayカード（paypay-card.co.jp からの利用速報）
 * @param {string} subject - メール件名
 * @param {string} body - メール本文
 * @param {string} sender - 送信元アドレス
 */
function parseCardEmail_(subject, body, sender) {
    // --- 三井住友カード ---
    if (sender.includes('vpass.ne.jp') || sender.includes('smbc-card.com')) {
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
    if (sender.includes('paypay-card.co.jp') && subject.includes('利用速報')) {
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
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 家計簿シート
 * @param {Object} record - 書き込むレコードオブジェクト
 * @returns {boolean} 書き込み成功なら true、重複スキップなら false
 */
function writeCardRecord_(sheet, record) {
    // 重複チェック: 同じ日付+金額+Method(F列)の組み合わせが既に存在する場合はスキップ
    const lastRow = Math.max(sheet.getLastRow(), 1);
    if (lastRow > 1) {
        const existingData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
        for (const row of existingData) {
            const existingDate = row[0] instanceof Date
                ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd')
                : String(row[0]).substring(0, 10);
            const existingAmount = Number(row[1]);
            const existingMethod = String(row[5]); // F列 = Method
            if (existingDate === record.date && existingAmount === record.amount && existingMethod === record.method) {
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
 * @param {string} query - Gmailの検索クエリ（日付範囲等）
 */
function fetchCardEmailsByQuery_(query) {
    if (!SPREADSHEET_ID) {
        console.warn('SPREADSHEET_ID が未設定です');
        return;
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('家計簿');
    if (!sheet) {
        console.warn('家計簿シートが見つかりません');
        return;
    }

    // 三井住友カードの利用通知 + PayPayカードの利用速報メールを検索
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
    const queries = [
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

    let found = 0;
    for (const q of queries) {
        const threads = GmailApp.search(q, 0, 5);
        if (threads.length > 0) {
            console.log('=== クエリ: ' + q + ' → ' + threads.length + '件 ===');
            for (const thread of threads) {
                const msgs = thread.getMessages();
                const msg = msgs[0];
                console.log('  📧 件名: ' + msg.getSubject());
                console.log('  📬 送信元: ' + msg.getFrom());
                console.log('  📅 日付: ' + msg.getDate());
                const bodySnippet = msg.getPlainBody().substring(0, 300).replace(/\n/g, ' ');
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
