// ===== AI分析モジュール =====
// Gemini APIを使ったAI家計分析・アドバイス生成

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

    const periodLabel = isWeekly ? "今週" : "今月";
    const prevPeriodLabel = isWeekly ? "先週" : "先月";

    // 期間の判定ロジック
    // isWeeklyの場合は直近7日間 vs その前の7日間を比較。
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
                currentPeriodExpense += amount;
                currentCategoryMap[category] = (currentCategoryMap[category] || 0) + amount;
            } else if (diffDays >= 7 && diffDays < 14) {
                previousPeriodExpense += amount;
                previousCategoryMap[category] = (previousCategoryMap[category] || 0) + amount;
            }
        } else {
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

    // プロンプト用データの整形: カテゴリ別 前期比diff付き
    const allCategories = new Set([...Object.keys(currentCategoryMap), ...Object.keys(previousCategoryMap)]);
    const categoryDiffLines = [];
    allCategories.forEach(cat => {
        const curr = currentCategoryMap[cat] || 0;
        const prev = previousCategoryMap[cat] || 0;
        const diff = curr - prev;
        const diffStr = diff > 0 ? `+${diff}円(↑)` : diff < 0 ? `${diff}円(↓)` : '±0';
        const pctChange = prev > 0 ? Math.round((diff / prev) * 100) : (curr > 0 ? '+∞' : '0');
        categoryDiffLines.push(`・${cat}: ${curr}円 (${prevPeriodLabel}: ${prev}円, 変動: ${diffStr}, ${pctChange}%)`);
    });
    const categoryAnalysisStr = categoryDiffLines.join('\n') || "記録なし";

    // 日別支出推移データの構築
    const dailyExpenses = {};
    data.forEach(row => {
        if (!row[0] || row[4] !== '支出') return;
        const d = new Date(row[0]);
        const amount = Number(row[1]) || 0;
        if (isWeekly) {
            const diffDays = Math.floor((now.getTime() - d.getTime()) / MS_PER_DAY);
            if (diffDays >= 0 && diffDays < 7) {
                const dayLabel = Utilities.formatDate(d, 'Asia/Tokyo', 'M/d(E)');
                dailyExpenses[dayLabel] = (dailyExpenses[dayLabel] || 0) + amount;
            }
        } else {
            const rYear = d.getFullYear();
            const rMonth = d.getMonth();
            if (rYear === currentYear && rMonth === currentMonth) {
                const dayLabel = Utilities.formatDate(d, 'Asia/Tokyo', 'M/d');
                dailyExpenses[dayLabel] = (dailyExpenses[dayLabel] || 0) + amount;
            }
        }
    });
    const dailyStr = Object.keys(dailyExpenses).sort().map(k => `  ${k}: ${dailyExpenses[k]}円`).join('\n') || "  記録なし";

    // 進行度（今月の場合）
    let budgetProgressStr = "";
    let dailyAvgStr = "";
    if (!isWeekly) {
        const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
        const progressPercent = Math.round((currentDay / daysInMonth) * 100);
        const expensePercent = Math.round((currentPeriodExpense / monthlyBudget) * 100);
        const remainingDays = daysInMonth - currentDay;
        const remainingBudget = monthlyBudget - currentPeriodExpense;
        const dailyAvailable = remainingDays > 0 ? Math.round(remainingBudget / remainingDays) : 0;
        budgetProgressStr = `【月間予算】: ${monthlyBudget}円\n  日数経過: ${currentDay}/${daysInMonth}日 (${progressPercent}%)\n  予算消化: ${currentPeriodExpense}/${monthlyBudget}円 (${expensePercent}%)\n  残り予算: ${remainingBudget}円 (残${remainingDays}日)`;
        dailyAvgStr = `1日あたりの許容上限: ${dailyAvailable}円/日`;
        if (currentDay > 0) {
            const dailyPace = Math.round(currentPeriodExpense / currentDay);
            const projectedTotal = dailyPace * daysInMonth;
            dailyAvgStr += `\n  現在の日平均: ${dailyPace}円/日\n  このペースの月末予測: ${projectedTotal}円 (予算比 ${Math.round(projectedTotal / monthlyBudget * 100)}%)`;
        }
    } else {
        const weeklyBudget = Math.floor(monthlyBudget / 4);
        const expensePercent = Math.round((currentPeriodExpense / weeklyBudget) * 100);
        budgetProgressStr = `【週次目安予算（月予算の1/4）】: ${weeklyBudget}円 (消化率: ${expensePercent}%)`;
        if (Object.keys(dailyExpenses).length > 0) {
            const dailyPace = Math.round(currentPeriodExpense / Object.keys(dailyExpenses).length);
            dailyAvgStr = `日平均支出: ${dailyPace}円/日`;
        }
    }

    // 浪費ランキング（支出額上位3カテゴリ）
    const sortedCats = Object.entries(currentCategoryMap).sort((a, b) => b[1] - a[1]);
    const top3Str = sortedCats.slice(0, 3).map((c, i) => `  ${i + 1}位: ${c[0]} ${c[1]}円 (全体の${Math.round(c[1] / (currentPeriodExpense || 1) * 100)}%)`).join('\n');

    // Gemini へのシステムプロンプト（人格設定）
    const systemPrompt = `あなたは10年以上の経験を持つ冷徹なファイナンシャルアナリストです。
家計簿データに基づき、感情を排して鋭く客観的な分析レポートを作成してください。

## 出力フォーマット（厳守）

### 📊 概況
予算に対する進捗と、前期比較の要約を2〜3文で。

### 🔍 カテゴリ別診断
前期比で増加が顕著なカテゴリを**金額と増加率つき**で指摘。
減少したカテゴリがあればそれも記載。

### ⚠️ 浪費アラート
支出上位カテゴリの中で「削減余地がある」ものを特定し、
**具体的にいくら削れば予算内に収まるか**を金額で提示。

### 📈 ペース診断
日次の支出ペースから月末の着地予測を計算し、
予算内に収まるかどうかを断定。

### 💡 アクション提案
残りの期間で予算内に着地するための**具体的な行動**を2〜3個、箇条書きで。

## ルール
- 絵文字はセクション見出しのみ使用可。本文には不要。
- 「頑張りましょう」等の応援は不要。事実と数字のみ。
- 全体で600〜800文字程度。`;

    // ユーザープロンプト（データ入力）
    const userPrompt = `以下の家計データから分析レポートを作成してください。

${budgetProgressStr}
${dailyAvgStr}

■ カテゴリ別支出（${periodLabel} vs ${prevPeriodLabel}）
${categoryAnalysisStr}

■ 支出額ランキング（${periodLabel}）
${top3Str || "  データなし"}

■ 日別支出推移（${periodLabel}）
${dailyStr}

■ 合計
  ${periodLabel}: ${currentPeriodExpense}円
  ${prevPeriodLabel}: ${previousPeriodExpense}円
  増減: ${currentPeriodExpense - previousPeriodExpense >= 0 ? '+' : ''}${currentPeriodExpense - previousPeriodExpense}円`;

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
 */
function getAiAnalysis(isWeekly) {
    try {
        const resultText = generateAiAnalysis(isWeekly);
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
