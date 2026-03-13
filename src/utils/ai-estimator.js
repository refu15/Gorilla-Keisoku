// AI活用余地の自動推定モジュール

/**
 * AI活用余地を推定するためのキーワードルール
 */
const AI_KEYWORDS = {
    // 高い可能性（既に使用中の可能性）
    highProbability: [
        'ChatGPT', 'GPT', 'Claude', 'Gemini', 'AI', 'LLM',
        '生成AI', '翻訳', 'コパイロット', 'Copilot'
    ],

    // AI活用の余地あり
    potential: [
        // 文書作成系
        '報告書', 'レポート', '提案書', '企画書', '議事録', 'ドキュメント',
        '資料作成', '文書作成', '下書き', 'ドラフト', '文章',

        // 分析・調査系
        '分析', '調査', 'リサーチ', '市場調査', '競合調査',

        // プログラミング系
        'コーディング', 'プログラミング', '開発', 'コードレビュー',
        'デバッグ', 'テスト作成', 'API',

        // デザイン系
        'デザイン', 'UI', 'UX', 'ワイヤーフレーム', 'モックアップ',

        // コミュニケーション系
        'メール作成', 'メール', '返信', 'チャット', 'Slack',

        // 翻訳・言語
        '翻訳', '英訳', '和訳', '英語', '多言語',

        // データ処理
        'Excel', 'スプレッドシート', 'データ整理', 'データ入力',
        '集計', 'グラフ作成'
    ],

    // AI活用が難しい可能性
    lowProbability: [
        '1on1', 'ミーティング', '会議', '定例', '朝会', '昼会',
        '顧客訪問', '商談', '面談', '面接', '研修',
        '移動', '出張', '休憩', 'ランチ', '昼食'
    ]
};

/**
 * イベントタイトルからAI活用余地を推定
 * @param {string} title イベントタイトル
 * @returns {Object} 推定結果
 */
export function estimateAIUsage(title) {
    const lowerTitle = title.toLowerCase();

    // 高い可能性のキーワードをチェック
    for (const keyword of AI_KEYWORDS.highProbability) {
        if (title.includes(keyword) || lowerTitle.includes(keyword.toLowerCase())) {
            return {
                flag: 'yes-using',
                confidence: 'high',
                matchedKeyword: keyword,
                suggestedRate: 50
            };
        }
    }

    // 活用余地ありのキーワードをチェック
    for (const keyword of AI_KEYWORDS.potential) {
        if (title.includes(keyword) || lowerTitle.includes(keyword.toLowerCase())) {
            return {
                flag: 'yes-potential',
                confidence: 'medium',
                matchedKeyword: keyword,
                suggestedRate: 0
            };
        }
    }

    // AI活用が難しいキーワードをチェック
    for (const keyword of AI_KEYWORDS.lowProbability) {
        if (title.includes(keyword) || lowerTitle.includes(keyword.toLowerCase())) {
            return {
                flag: 'no',
                confidence: 'medium',
                matchedKeyword: keyword,
                suggestedRate: 0
            };
        }
    }

    // デフォルト: 判定できない
    return {
        flag: 'unknown',
        confidence: 'low',
        matchedKeyword: null,
        suggestedRate: 0
    };
}

/**
 * 複数のイベントに対して一括推定
 * @param {Array} events イベント一覧
 * @returns {Array} 推定結果付きイベント一覧
 */
export function estimateAllEvents(events) {
    return events.map(event => ({
        ...event,
        aiEstimate: estimateAIUsage(event.summary || event.title || '')
    }));
}

/**
 * AI活用フラグの選択肢
 */
export const AI_FLAG_OPTIONS = [
    { value: 'no', label: 'No - 活用なし' },
    { value: 'yes-using', label: 'Yes - 既に使用中' },
    { value: 'yes-potential', label: 'Yes - 活用余地あり' },
    { value: 'yes-failed', label: 'Yes - 失敗・課題あり' }
];

/**
 * フラグ値からラベルを取得
 * @param {string} value フラグ値
 * @returns {string} ラベル
 */
export function getFlagLabel(value) {
    const option = AI_FLAG_OPTIONS.find(opt => opt.value === value);
    return option ? option.label : 'No - 活用なし';
}
