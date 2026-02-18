// Notion API連携モジュール (バックグラウンド経由)

/**
 * バックグラウンドサービスワーカー経由でNotion APIにリクエストを送信
 */
async function notionRequest(url, method, headers, body) {
    return new Promise((resolve) => {
        chrome.runtime.sendMessage(
            { action: 'notionRequest', url, method, headers, body },
            (response) => resolve(response)
        );
    });
}

/**
 * Notionデータベースに日報を保存
 */
export async function saveToNotion(reportData) {
    const { notionToken, notionDatabaseId } = await chrome.storage.sync.get(['notionToken', 'notionDatabaseId']);

    if (!notionToken || !notionDatabaseId) {
        throw new Error('Notion設定が完了していません。設定画面から設定してください。');
    }

    const result = await createNotionPage(notionToken, notionDatabaseId, reportData);

    if (!result.success) {
        throw new Error(result.error);
    }

    return { success: true, pageId: result.data?.id };
}

/**
 * Notionページを作成（コンパクト版 - 100ブロック制限対応）
 */
async function createNotionPage(token, databaseId, reportData) {
    const url = 'https://api.notion.com/v1/pages';

    const headers = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
    };

    const pageTitle = `${reportData.date} 日報 - ${reportData.userEmail}`;
    const contentBlocks = generateCompactBlocks(reportData);

    const body = {
        parent: { database_id: databaseId },
        properties: {
            '名前': {
                title: [{ text: { content: pageTitle } }]
            }
        },
        children: contentBlocks
    };

    return notionRequest(url, 'POST', headers, body);
}

/**
 * コンパクトな日報ブロックを生成（100ブロック以内）
 */
function generateCompactBlocks(reportData) {
    const blocks = [];

    // サマリー（1ブロック）
    blocks.push({
        object: 'block',
        type: 'callout',
        callout: {
            rich_text: [{
                type: 'text',
                text: {
                    content: `📊 サマリー\n日付: ${reportData.date}\nユーザー: ${reportData.userEmail}\n合計時間: ${reportData.totalWorkHours}\n予定数: ${reportData.scheduleEntries.length}件`
                }
            }],
            icon: { type: 'emoji', emoji: '📋' }
        }
    });

    // 区切り線
    blocks.push({ object: 'block', type: 'divider', divider: {} });

    // 予定一覧見出し
    blocks.push({
        object: 'block',
        type: 'heading_2',
        heading_2: {
            rich_text: [{ type: 'text', text: { content: '📅 予定一覧' } }]
        }
    });

    // 各予定（1予定 = 1ブロックにまとめる）
    // 最大90件まで（100ブロック制限のため）
    const maxEntries = Math.min(reportData.scheduleEntries.length, 90);

    for (let i = 0; i < maxEntries; i++) {
        const entry = reportData.scheduleEntries[i];
        const flagEmoji = entry.aiFlag === 'no' ? '❌' : entry.aiFlag === 'yes-using' ? '✅' : '💡';
        const flagLabel = getAIFlagLabel(entry.aiFlag);

        let content = `${entry.title}\n⏰ ${entry.start} - ${entry.end} | ${flagEmoji} ${flagLabel} | 📈 活用率: ${entry.aiRate}%`;

        if (entry.note) {
            content += `\n📝 ${entry.note}`;
        }

        blocks.push({
            object: 'block',
            type: 'quote',
            quote: {
                rich_text: [{ type: 'text', text: { content } }]
            }
        });
    }

    // 省略された予定がある場合
    if (reportData.scheduleEntries.length > maxEntries) {
        blocks.push({
            object: 'block',
            type: 'paragraph',
            paragraph: {
                rich_text: [{
                    type: 'text',
                    text: { content: `... 他 ${reportData.scheduleEntries.length - maxEntries}件の予定` }
                }]
            }
        });
    }

    return blocks;
}

/**
 * AI活用フラグのラベルを取得
 */
function getAIFlagLabel(flag) {
    const labels = {
        'no': 'No',
        'yes-using': 'Yes-使用中',
        'yes-potential': 'Yes-余地あり'
    };
    return labels[flag] || 'No';
}

/**
 * Notion接続テスト
 */
export async function testNotionConnection() {
    const { notionToken, notionDatabaseId } = await chrome.storage.sync.get(['notionToken', 'notionDatabaseId']);

    if (!notionToken || !notionDatabaseId) {
        return { success: false, error: 'トークンまたはデータベースIDが設定されていません' };
    }

    const url = `https://api.notion.com/v1/databases/${notionDatabaseId}`;
    const headers = {
        'Authorization': `Bearer ${notionToken}`,
        'Notion-Version': '2022-06-28'
    };

    const result = await notionRequest(url, 'GET', headers, null);

    if (!result.success) {
        return { success: false, error: result.error };
    }

    return {
        success: true,
        databaseTitle: result.data.title?.[0]?.plain_text || 'Untitled',
        properties: Object.keys(result.data.properties || {})
    };
}

/**
 * Notion設定の保存
 */
export async function saveNotionConfig(token, databaseId) {
    await chrome.storage.sync.set({ notionToken: token, notionDatabaseId: databaseId });
}

/**
 * Notion設定の取得
 */
export async function getNotionConfig() {
    return chrome.storage.sync.get(['notionToken', 'notionDatabaseId']);
}

/**
 * 送信先設定の保存
 */
export async function saveDestinationConfig(enableSheets, enableNotion) {
    await chrome.storage.sync.set({ enableSheets, enableNotion });
}

/**
 * 送信先設定の取得
 */
export async function getDestinationConfig() {
    const config = await chrome.storage.sync.get(['enableSheets', 'enableNotion']);
    return {
        enableSheets: config.enableSheets !== false,
        enableNotion: config.enableNotion === true
    };
}
