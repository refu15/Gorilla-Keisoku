// Chrome Storage ユーティリティモジュール

/**
 * 日報下書きを保存
 * @param {string} date 日付（YYYY-MM-DD）
 * @param {Object} draft 下書きデータ
 * @returns {Promise<void>}
 */
export async function saveDraft(date, draft) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        console.error('saveDraft: 不正な日付形式:', date);
        return;
    }
    const key = `draft_${date}`;
    await chrome.storage.local.set({ [key]: draft });

    // 保存のついでに古い下書き（7日以上前のもの）をクリーンアップ
    cleanupOldDrafts().catch(e => console.warn('下書きクリーンアップエラー:', e));
}

/**
 * 7日以上前の古い下書きを自動削除する
 * ストレージ容量の逼迫を防ぐための安全装置
 */
async function cleanupOldDrafts() {
    const data = await chrome.storage.local.get(null);
    const keysToRemove = [];
    const now = Date.now();
    const SEVEN_DAYS = 7 * 24 * 60 * 60 * 1000;

    for (const key of Object.keys(data)) {
        if (key.startsWith('draft_')) {
            const dateStr = key.replace('draft_', '');
            const draftDate = new Date(dateStr).getTime();
            if (!isNaN(draftDate) && (now - draftDate > SEVEN_DAYS)) {
                keysToRemove.push(key);
            }
        }
    }

    if (keysToRemove.length > 0) {
        await chrome.storage.local.remove(keysToRemove);
    }
}

/**
 * 日報下書きを取得
 * @param {string} date 日付（YYYY-MM-DD）
 * @returns {Promise<Object|null>} 下書きデータ
 */
export async function getDraft(date) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        console.error('getDraft: 不正な日付形式:', date);
        return null;
    }
    const key = `draft_${date}`;
    const result = await chrome.storage.local.get(key);
    return result[key] || null;
}

/**
 * 日報下書きを削除
 * @param {string} date 日付（YYYY-MM-DD）
 * @returns {Promise<void>}
 */
export async function deleteDraft(date) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        console.error('deleteDraft: 不正な日付形式:', date);
        return;
    }
    const key = `draft_${date}`;
    await chrome.storage.local.remove(key);
}

/**
 * ユーザー情報を保存
 * @param {Object} userInfo ユーザー情報
 * @returns {Promise<void>}
 */
export async function saveUserInfo(userInfo) {
    await chrome.storage.local.set({ userInfo });
}

/**
 * 日付をYYYY-MM-DD形式にフォーマット
 * @param {Date} date 日付
 * @returns {string} YYYY-MM-DD形式
 */
export function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

/**
 * 日付を日本語表示用にフォーマット
 * @param {Date} date 日付
 * @returns {string} YYYY年M月D日（曜日）形式
 */
export function formatDateJapanese(date) {
    const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const weekday = weekdays[date.getDay()];
    return `${year}年${month}月${day}日（${weekday}）`;
}
