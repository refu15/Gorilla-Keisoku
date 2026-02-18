// オフラインキュー管理モジュール

/**
 * オフラインキューにアイテムを追加
 * @param {Object} data 保存するデータ
 * @returns {Promise<string>} アイテムID
 */
export async function addToQueue(data) {
    const { offlineQueue = [] } = await chrome.storage.local.get('offlineQueue');

    const item = {
        id: generateId(),
        data,
        createdAt: new Date().toISOString(),
        retryCount: 0
    };

    offlineQueue.push(item);
    await chrome.storage.local.set({ offlineQueue });

    return item.id;
}

/**
 * オフラインキューからアイテムを削除
 * @param {string} id アイテムID
 * @returns {Promise<void>}
 */
export async function removeFromQueue(id) {
    const { offlineQueue = [] } = await chrome.storage.local.get('offlineQueue');
    const filtered = offlineQueue.filter(item => item.id !== id);
    await chrome.storage.local.set({ offlineQueue: filtered });
}

/**
 * オフラインキューの内容を取得
 * @returns {Promise<Array>} キュー内のアイテム
 */
export async function getQueue() {
    const { offlineQueue = [] } = await chrome.storage.local.get('offlineQueue');
    return offlineQueue;
}

/**
 * オフラインキューをクリア
 * @returns {Promise<void>}
 */
export async function clearQueue() {
    await chrome.storage.local.set({ offlineQueue: [] });
}

/**
 * オフラインキューの再送を試行
 * @returns {Promise<Object>} 再送結果
 */
export async function retryQueue() {
    return new Promise((resolve) => {
        chrome.runtime.sendMessage({ action: 'retryOfflineQueue' }, resolve);
    });
}

/**
 * ユニークIDを生成
 * @returns {string} ID
 */
function generateId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * オフラインかどうかを判定
 * @returns {boolean}
 */
export function isOffline() {
    return !navigator.onLine;
}

/**
 * オンライン/オフライン状態の変更を監視
 * @param {Function} callback コールバック関数
 */
export function watchOnlineStatus(callback) {
    window.addEventListener('online', () => callback(true));
    window.addEventListener('offline', () => callback(false));
}
