// Gemini API連携モジュール - AI分析機能

/**
 * Gemini APIを使用して日報データを分析
 * @param {Object} summaryData 日次集計データ
 * @returns {Promise<Object>} 分析結果
 */
export async function analyzeWithGemini(summaryData) {
    const { geminiApiKey } = await chrome.storage.sync.get(['geminiApiKey']);

    if (!geminiApiKey) {
        throw new Error('Gemini API Keyが設定されていません。設定画面から設定してください。');
    }

    const prompt = buildAnalysisPrompt(summaryData);

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiApiKey}`;

    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            contents: [{
                parts: [{ text: prompt }]
            }],
            generationConfig: {
                temperature: 0.7,
                maxOutputTokens: 1024
            }
        })
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'Gemini API呼び出しに失敗しました');
    }

    const data = await response.json();
    const analysisText = data.candidates?.[0]?.content?.parts?.[0]?.text || '';

    return {
        success: true,
        analysis: analysisText,
        timestamp: new Date().toISOString()
    };
}

/**
 * 分析用のプロンプトを構築
 */
function buildAnalysisPrompt(summaryData) {
    return `あなたは業務効率化コンサルタントです。以下の日報データを分析し、AI活用の改善提案を行ってください。

## 日報データ
- 日付: ${summaryData.date}
- ユーザー: ${summaryData.userEmail}
- 合計予定数: ${summaryData.totalEvents}件
- 合計稼働時間: ${summaryData.totalMinutes}分（${Math.round(summaryData.totalMinutes / 60 * 10) / 10}時間）

## AI活用状況
- AI活用中: ${summaryData.aiUsingCount}件（${summaryData.aiUsingMinutes}分）
- AI未活用: ${summaryData.aiNotUsingCount}件（${summaryData.aiNotUsingMinutes}分）
- AI活用余地あり: ${summaryData.aiPotentialCount}件（${summaryData.aiPotentialMinutes}分）
- 平均活用率: ${summaryData.avgRate}%

## 予定の詳細
${summaryData.entries.map(e =>
        `- ${e.title}（${e.start}-${e.end}）: ${e.aiFlag === 'yes-using' ? '✅活用中' : e.aiFlag === 'yes-potential' ? '💡余地あり' : '❌未活用'} ${e.aiRate}%${e.note ? ` / ${e.note}` : ''}`
    ).join('\n')}

## 分析してほしいこと
1. **本日のAI活用評価**（100点満点で採点、一言コメント）
2. **改善ポイント**（具体的に2-3点）
3. **AI活用できそうな業務の提案**（未活用・余地ありの予定について具体的なAIツール名と使い方を提案）
4. **明日への一言アドバイス**

回答は日本語で、簡潔にまとめてください。マークダウン形式で、見出しを使って構造化してください。`;
}

/**
 * 週次分析を実行
 */
export async function analyzeWeeklySummary(weeklyData) {
    const { geminiApiKey } = await chrome.storage.sync.get(['geminiApiKey']);

    if (!geminiApiKey) {
        throw new Error('Gemini API Keyが設定されていません');
    }

    const prompt = buildWeeklyPrompt(weeklyData);

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiApiKey}`;

    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: {
                temperature: 0.7,
                maxOutputTokens: 2048
            }
        })
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'Gemini API呼び出しに失敗しました');
    }

    const data = await response.json();
    return {
        success: true,
        analysis: data.candidates?.[0]?.content?.parts?.[0]?.text || '',
        timestamp: new Date().toISOString()
    };
}

function buildWeeklyPrompt(weeklyData) {
    return `あなたは業務効率化コンサルタントです。以下の週次データを分析し、AI活用トレンドと改善提案を行ってください。

## 週次サマリー
${JSON.stringify(weeklyData, null, 2)}

## 分析してほしいこと
1. **週間AI活用トレンド**（グラフが描けるような数値データと解説）
2. **最も改善余地のある業務カテゴリ**
3. **今週のベストプラクティス**（うまく活用できた事例）
4. **来週への具体的アクションプラン**（3つ）

回答は日本語で、マークダウン形式でお願いします。`;
}

/**
 * Gemini API Key設定の保存
 */
export async function saveGeminiConfig(apiKey) {
    await chrome.storage.sync.set({ geminiApiKey: apiKey });
}

/**
 * Gemini API Key設定の取得
 */
export async function getGeminiConfig() {
    return chrome.storage.sync.get(['geminiApiKey']);
}

/**
 * Gemini API接続テスト
 */
export async function testGeminiConnection() {
    const { geminiApiKey } = await chrome.storage.sync.get(['geminiApiKey']);

    if (!geminiApiKey) {
        return { success: false, error: 'API Keyが設定されていません' };
    }

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiApiKey}`;

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{ parts: [{ text: 'こんにちは' }] }],
                generationConfig: { maxOutputTokens: 10 }
            })
        });

        if (!response.ok) {
            const error = await response.json();
            return { success: false, error: error.error?.message || 'API呼び出しエラー' };
        }

        return { success: true, message: 'Gemini APIに接続しました' };
    } catch (error) {
        return { success: false, error: error.message };
    }
}
