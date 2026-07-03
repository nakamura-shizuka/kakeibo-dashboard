// ===== AI分析モジュール =====
// Gemini APIを使った月次家計レポート生成
// 対象は「先月」の確定データ（月次確定シート＋残高実測）。
// 明細ベースの週次分析は廃止済み（不正確なデータの頻回分析は読まれないため）。

/**
 * 千円単位の読みやすい金額表記にする（例: 105000 → 10.5万円, 800 → 800円）
 */
function formatYenForPrompt_(amount) {
    const n = Math.round(Number(amount) || 0);
    if (Math.abs(n) >= 10000) {
        const man = Math.round(n / 1000) / 10;
        return man + '万円';
    }
    return n.toLocaleString() + '円';
}

/**
 * 🤖 先月の確定データをもとに「家計プランの見守りレポート」を生成する
 * 視点: ①先月は赤字か ②使っていい額の答え合わせ ③NISA継続可否 ④イベント準備 ⑤今月の一手
 */
function generateAiAnalysis() {
    if (!GEMINI_API_KEY) return "AI分析機能が有効ではありません（GEMINI_API_KEY未設定）。";
    if (!SPREADSHEET_ID) return "DBが設定されていません。";

    const now = new Date();
    // 対象は先月（確定データが揃っている直近の月）
    const target = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const y = target.getFullYear();
    const m = target.getMonth(); // 0-11
    const monthLabel = y + '年' + (m + 1) + '月';

    const settings = getSettingsData();
    const income = Number(settings.income) || 0;
    if (income <= 0) {
        return "分析には月収の設定が必要です。ダッシュボードの⚙️設定から「月収（手取り）」を入力してください。";
    }

    // 家計モデルと先月の確定実績（計算ロジックはダッシュボードと共通）
    const s = buildSavingsSummary_(settings, y, m, []);
    const hasCardData = s.cardBreakdown.some(function (c) { return c.amount > 0; });
    if (s.confirmedSpending <= 0 && !hasCardData) {
        return monthLabel + "の確定データがまだありません。請求確定メールの取込（fetchBillingEmails）後にもう一度お試しください。";
    }

    const variableActual = s.confirmedSpending - s.fixedSum; // 先月の変動費実績
    const variableDiff = variableActual - s.safeToSpend;     // 計画との差（プラス=使いすぎ）

    // 直近3ヶ月の変動費平均（NISA継続判定の材料。確定データがある月のみ）
    let variableSum = 0, variableMonths = 0;
    for (let i = 1; i <= 3; i++) {
        const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
        const c = calcConfirmedSpending_(d.getFullYear(), d.getMonth());
        if (c.total > 0) {
            variableSum += (c.total - s.fixedSum);
            variableMonths++;
        }
    }
    const avgVariable = variableMonths > 0 ? Math.round(variableSum / variableMonths) : variableActual;

    // NISA継続の判定材料: 変動費の実測平均を前提に、NISA+予備費を払える余力があるか
    const cashFlowForNisa = income - s.fixedSum - avgVariable - s.eventsMonthly;
    const nisaHeadroom = cashFlowForNisa - s.nisaMonthly - s.reserve;

    // カード別内訳（未確定があれば明示）
    const cardLines = s.cardBreakdown.map(function (c) {
        return '・' + c.card + ': ' + (c.confirmed ? formatYenForPrompt_(c.amount) : '未確定（要確認）');
    }).join('\n');

    // 実測貯蓄（残高入力ベース）
    const actualSavingsStr = (s.actualSavings !== null && s.actualSavings !== undefined)
        ? formatYenForPrompt_(s.actualSavings)
        : 'データなし（残高が未入力の月がある）';

    // 今後の年間イベント
    const currentMonth1 = now.getMonth() + 1;
    const upcomingEvents = (settings.annualEvents || []).filter(function (ev) {
        return !ev.month || Number(ev.month) >= currentMonth1;
    });
    const eventsStr = upcomingEvents.length > 0
        ? upcomingEvents.map(function (ev) {
            return '・' + (ev.month ? ev.month + '月 ' : '時期未定 ') + ev.name + ' ' + formatYenForPrompt_(ev.amount);
        }).join('\n')
        : '登録なし';

    // カテゴリ内訳（参考・明細ベース。傾向把握のみに使用）
    let categoryStr = '記録なし';
    let uncategorizedRate = 0;
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('家計簿');
        if (sheet && sheet.getLastRow() > 1) {
            const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
            const catMap = {};
            let catTotal = 0, uncatTotal = 0;
            rows.forEach(function (row) {
                if (!row[0] || (row[4] || '支出') !== '支出') return;
                const d = new Date(row[0]);
                if (d.getFullYear() !== y || d.getMonth() !== m) return;
                const amount = Number(row[1]) || 0;
                const cat = row[2] || '未分類';
                catMap[cat] = (catMap[cat] || 0) + amount;
                catTotal += amount;
                if (cat === '未分類') uncatTotal += amount;
            });
            const sorted = Object.entries(catMap).sort(function (a, b) { return b[1] - a[1]; });
            if (sorted.length > 0) {
                categoryStr = sorted.slice(0, 5).map(function (c) {
                    return '・' + c[0] + ': ' + formatYenForPrompt_(c[1]);
                }).join('\n');
                uncategorizedRate = catTotal > 0 ? Math.round(uncatTotal / catTotal * 100) : 0;
            }
        }
    } catch (e) { console.warn('カテゴリ集計失敗:', e.message); }

    // Gemini へのシステムプロンプト
    const systemPrompt = `あなたは家計プランの伴走者（信頼できるファイナンシャルプランナー）です。
ユーザーの目的は「貯蓄と消費のバランスが保たれ、年間でやりたいことができ、子どもの進学や老後に心理的な安心を持つこと」。
渡されたデータはすべて計算済みです。あなたの仕事は計算ではなく、「結局どうなのか」を分かりやすく伝えることです。

## 出力フォーマット（この5ブロックのみ。LINEで読むため Markdown記法は禁止、絵文字見出し＋改行だけで構成）

🌡️ 先月の結論
黒字か赤字かを最初の一文で断定し、貯蓄額を添える。

💸 使っていい額の答え合わせ
変動費の計画と実績の差を一言で。超過していたら最大の原因カテゴリを1つだけ挙げる。

📈 NISAは続けて大丈夫？
判定材料の数字を使って「余裕あり／ぎりぎり／見直し推奨」のどれかを金額つきで断定。

🎉 イベント準備
直近の年間イベントに向けた一言（イベント登録がなければこのブロックごと省略）。

✅ 今月やること
1つだけ。具体的な金額つきで。

## ルール
- 全体で300〜500文字。短いほど良い。
- 数字は「約10.5万円」のような読みやすい表記のまま使う。細かい端数を書かない。
- 責めない・煽らない・説教しない。ただし赤字や不足は曖昧にせずはっきり書く。
- 「データなし」「未確定」の項目は、無理に分析せず「まだ分からない」と正直に書く。
- 専門用語・横文字を使わない。「キャッシュフロー」ではなく「お金の出入り」。`;

    // ユーザープロンプト（計算済みデータ）
    const userPrompt = `■ ${monthLabel}の実績（確定ベース）
・確定支出の合計: ${formatYenForPrompt_(s.confirmedSpending)}（うち固定費 ${formatYenForPrompt_(s.fixedSum)}）
・カード請求の内訳:
${cardLines}
・現金の実測貯蓄（残高の前月差分）: ${actualSavingsStr}
・貯蓄見込み（月収 − NISA − 確定支出）: ${formatYenForPrompt_(s.projectedSavings)}

■ 家計プラン（設定値）
・月収 ${formatYenForPrompt_(income)} − 固定費 ${formatYenForPrompt_(s.fixedSum)} − NISA ${formatYenForPrompt_(s.nisaMonthly)} − イベント積立 ${formatYenForPrompt_(s.eventsMonthly)} − 予備費 ${formatYenForPrompt_(s.reserve)}
　→ 使っていい額（変動費の予算）: ${formatYenForPrompt_(s.safeToSpend)}
・${monthLabel}の変動費実績: ${formatYenForPrompt_(variableActual)}（予算比 ${variableDiff >= 0 ? '+' : ''}${formatYenForPrompt_(variableDiff)}）

■ NISA継続の判定材料（直近${variableMonths || 1}ヶ月の実績ベース）
・変動費の平均: ${formatYenForPrompt_(avgVariable)}
・月収 − 固定費 − 変動費平均 − イベント積立 = ${formatYenForPrompt_(cashFlowForNisa)}
・そこから NISA ${formatYenForPrompt_(s.nisaMonthly)} と予備費 ${formatYenForPrompt_(s.reserve)} を払うと ${nisaHeadroom >= 0 ? '余裕' : '不足'} ${formatYenForPrompt_(Math.abs(nisaHeadroom))}

■ 今後の年間イベント
${eventsStr}

■ カテゴリ内訳（参考値・明細ベース・未分類率${uncategorizedRate}%）
${categoryStr}`;

    // Gemini API リクエスト (gemini-2.5-flash を使用)
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
    const payload = {
        "contents": [
            { "role": "user", "parts": [{ "text": systemPrompt + "\n\n" + userPrompt }] }
        ],
        "generationConfig": {
            "temperature": 0.3,
            "maxOutputTokens": 4000
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
            return `分析エラー: AIへの接続に失敗しました(${resJson.error.message || '詳細不明'})`;
        }

        if (!resJson.candidates || resJson.candidates.length === 0 || !resJson.candidates[0].content) {
            console.error("Gemini API Parse Error:", resJson);
            return "分析エラー: AIからの応答形式が想定外でした。データが少なすぎるか、APIキーの設定不備の可能性があります。";
        }

        const answer = resJson.candidates[0].content.parts[0].text;
        return answer ? answer.trim() : "分析エラー: AIから空の回答が返されました。";

    } catch (e) {
        console.error("AI分析実行エラー:", e);
        return "分析エラー: ネットワークまたはシステムエラーが発生しました。 (" + e.message + ")";
    }
}

/**
 * 📊 ダッシュボードの即時分析ボタンから呼ばれるAPI
 * 週次分析は廃止済み（不正確な明細ベースの頻回分析は読まれないため）。常に月次で実行する。
 */
function getAiAnalysis() {
    try {
        const resultText = generateAiAnalysis();
        // generateAiAnalysis はエラー時も文字列を返すため、エラープレフィックスで判定
        if (resultText && resultText.startsWith('分析エラー:')) {
            logError('getAiAnalysis', resultText);
            return { success: false, message: resultText };
        }
        return { success: true, analysis: resultText };
    } catch (error) {
        logError('getAiAnalysis 例外', error.message);
        return { success: false, message: error.message };
    }
}

/**
 * 🤖 Gemini APIを呼び出してテキストを生成する（短文アドバイス用）
 * @param {string} promptText - Geminiへのプロンプト
 */
function callGeminiAPI(promptText) {
    if (!GEMINI_API_KEY) return "AIアドバイザーは現在お休み中です（APIキー未設定）";

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
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

    const targetUserId = getLineUserId_(); // push通知用

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
