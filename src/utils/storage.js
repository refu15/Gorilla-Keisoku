// Chrome Storage ユーティリティモジュール

/**
 * ユーザー設定を保存
 * @param {Object} settings 設定オブジェクト
 * @returns {Promise<void>}
 */
export async function saveSettings(settings) {
    await chrome.storage.sync.set({ userSettings: settings });
}

/**
 * ユーザー設定を取得
 * @returns {Promise<Object>} 設定オブジェクト
 */
export async function getSettings() {
    const { userSettings = {} } = await chrome.storage.sync.get('userSettings');
    return {
        // デフォルト設定
        autoEstimate: true,
        showConfidence: true,
        defaultIncludeInReport: true,
        ...userSettings
    };
}

/**
 * 日報下書きを保存
 * @param {string} date 日付（YYYY-MM-DD）
 * @param {Object} draft 下書きデータ
 * @returns {Promise<void>}
 */
export async function saveDraft(date, draft) {
    const key = `draft_${date}`;
    await chrome.storage.local.set({ [key]: draft });
}

/**
 * 日報下書きを取得
 * @param {string} date 日付（YYYY-MM-DD）
 * @returns {Promise<Object|null>} 下書きデータ
 */
export async function getDraft(date) {
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
 * ユーザー情報を取得
 * @returns {Promise<Object|null>} ユーザー情報
 */
export async function getUserInfo() {
    const { userInfo } = await chrome.storage.local.get('userInfo');
    return userInfo || null;
}

/**
 * 最後に送信した日付を保存
 * @param {string} date 日付（YYYY-MM-DD）
 * @returns {Promise<void>}
 */
export async function saveLastSubmittedDate(date) {
    await chrome.storage.local.set({ lastSubmittedDate: date });
}

/**
 * 最後に送信した日付を取得
 * @returns {Promise<string|null>} 日付
 */
export async function getLastSubmittedDate() {
    const { lastSubmittedDate } = await chrome.storage.local.get('lastSubmittedDate');
    return lastSubmittedDate || null;
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
