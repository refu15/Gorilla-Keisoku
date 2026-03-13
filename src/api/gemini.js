// Gemini API連携モジュール - AI分析機能

import { sanitizeForPrompt } from '../utils/sanitize.js';

/**
 * Gemini APIを使用して日報データを分析
 * @param {Object} summaryData 日次集計データ
 * @returns {Promise<Object>} 分析結果
 */
export async function analyzeWithGemini(summaryData) {
    const { geminiApiKey } = await chrome.storage.local.get(['geminiApiKey']);

    if (!geminiApiKey) {
        throw new Error('Gemini API Keyが設定されていません。設定画面から設定してください。');
    }

    // 期間に応じたプロンプト生成
    let prompt;
    if (summaryData.period === 'daily') {
        prompt = `
あなたはプロの生産性向上コンサルタントです。
以下の日次作業ログを分析し、ユーザーに対するフィードバックレポートを作成してください。

## ユーザー情報
Email: ${summaryData.userEmail}
日付: ${summaryData.date}

## 作業統計
- 総予定数: ${summaryData.totalEvents}件
- 総作業時間: ${summaryData.totalTime || 'N/A'} (内訳から推定)

## 作業詳細 (Time, Title, AI Flag, AI Rate, AI Score)
${summaryData.entries.map(e => `- ${e.start}-${e.end}: ${sanitizeForPrompt(e.title)} [AI: ${e.aiFlag}, Rate: ${e.aiRate}%, Score: ${e.aiScore}pt]`).join('\n')}

## 指示
以下の構成で、簡潔かつ具体的なMarkdown形式のレポートを作成してください。

1. **全体サマリー** (100文字以内): 本日の働き方の特徴とAI活用の概要。
2. **AI活用の評価**:
   - AIスコアが高い作業（効果的に活用できた点）への称賛。
   - AIスコアが低いが、活用余地があった作業（Potential）への具体的な改善提案。
3. **明日に向けたアクション**: 具体的に取り組むべきこと1つ。

トーン＆マナー:
- ポジティブで励ましのあるトーン
- 具体的で実行可能なアドバイス
- 箇条書きを活用して読みやすく
`;
    } else {
        // 週次・月次 (既存ロジック)
        prompt = `
あなたはプロの生産性向上コンサルタントです。
以下の作業集計データを分析し、ユーザーに対するフィードバックレポートを作成してください。

## ユーザー情報
Email: ${summaryData.userEmail}
期間: ${summaryData.date}

## 統計データ
- 総作業時間: ${summaryData.totalMinutes}分
- AI活用時間: ${summaryData.aiUsingMinutes}分
- AI活用余地あり時間: ${summaryData.aiPotentialMinutes}分
- AI活用なし時間: ${summaryData.aiNotUsingMinutes}分
- AI平均活用率: ${summaryData.avgRate}% (AI使用中の作業における平均)

## 指示
以下の構成で、Markdown形式のレポートを作成してください。
1. **総評**: 今回の期間の働き方とAI活用の特徴（200文字程度）
2. **AI活用の成果**: AIによってどれくらいの時間が効率化されたか、または質が向上したかの推定
3. **改善の機会**: 「AI活用余地あり」の時間を減らし、「AI活用時間」を増やすための具体的なアドバイス
4. **次の期間の目標**: 具体的なアクションプラン

トーン＆マナー:
- ポジティブで建設的
- データに基づいた客観的な分析
- 簡潔で読みやすい文章
`;
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`;

    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'x-goog-api-key': geminiApiKey
        },
        body: JSON.stringify({
            contents: [{
                parts: [{ text: prompt }]
            }],
            generationConfig: {
                temperature: 0.7,
                maxOutputTokens: 8192
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

// ※ buildAnalysisPrompt は未使用のため削除しました（デッドコード整理）

/**
 * 週次分析を実行
 */
export async function analyzeWeeklySummary(weeklyData) {
    const { geminiApiKey } = await chrome.storage.local.get(['geminiApiKey']);

    if (!geminiApiKey) {
        throw new Error('Gemini API Keyが設定されていません');
    }

    const prompt = buildWeeklyPrompt(weeklyData);

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`;

    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': geminiApiKey },
        body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: {
                temperature: 0.7,
                maxOutputTokens: 8192
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
    // プロンプトインジェクション対策: weeklyData全体をJSON文字列化せず、必要フィールドのみ明示的に埋め込む
    const totalMinutes = Number(weeklyData.totalMinutes) || 0;
    const aiUsingMinutes = Number(weeklyData.aiUsingMinutes) || 0;
    const aiPotentialMinutes = Number(weeklyData.aiPotentialMinutes) || 0;
    const aiNotUsingMinutes = Number(weeklyData.aiNotUsingMinutes) || 0;
    const avgRate = Number(weeklyData.avgRate) || 0;
    const period = sanitizeForPrompt(String(weeklyData.period || ''), 50);

    return `あなたは業務効率化コンサルタントです。以下の週次データを分析し、AI活用トレンドと改善提案を行ってください。

## 週次サマリー
- 期間: ${period}
- 総作業時間: ${totalMinutes}分
- AI活用時間: ${aiUsingMinutes}分
- AI活用余地あり時間: ${aiPotentialMinutes}分
- AI活用なし時間: ${aiNotUsingMinutes}分
- AI平均活用率: ${avgRate}%

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
    await chrome.storage.local.set({ geminiApiKey: apiKey });
}

/**
 * Gemini API Key設定の取得
 */
export async function getGeminiConfig() {
    return chrome.storage.local.get(['geminiApiKey']);
}

/**
 * Gemini API接続テスト
 */
export async function testGeminiConnection() {
    const { geminiApiKey } = await chrome.storage.local.get(['geminiApiKey']);

    if (!geminiApiKey) {
        return { success: false, error: 'API Keyが設定されていません' };
    }

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`;

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'x-goog-api-key': geminiApiKey },
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
        console.error('Gemini接続テストエラー詳細:', error);
        return { success: false, error: 'Gemini APIへの接続に失敗しました。ネットワークを確認してください。' };
    }
}
