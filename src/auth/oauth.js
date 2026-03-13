// OAuth認証モジュール

/**
 * Google OAuth認証を実行してトークンを取得
 * @param {boolean} interactive インタラクティブなログインプロンプトを表示するか
 * @returns {Promise<string>} アクセストークン
 */
export async function getAuthToken(interactive = true) {
    return new Promise((resolve, reject) => {
        chrome.runtime.sendMessage({ action: 'getAuthToken', interactive }, (response) => {
            if (chrome.runtime.lastError) {
                return reject(new Error('通信エラーが発生しました。時間を置いて再度お試しください (' + chrome.runtime.lastError.message + ')'));
            }
            if (response && response.success) {
                resolve(response.token);
            } else {
                reject(new Error(response?.error || '認証メッセージの送信に失敗しました'));
            }
        });
    });
}

/**
 * 認証トークンを失効させてログアウト
 * @returns {Promise<void>}
 */
export async function revokeToken() {
    return new Promise((resolve, reject) => {
        chrome.runtime.sendMessage({ action: 'revokeToken' }, (response) => {
            if (chrome.runtime.lastError) {
                return reject(new Error('通信エラーが発生しました。時間を置いて再度お試しください (' + chrome.runtime.lastError.message + ')'));
            }
            if (response && response.success) {
                resolve();
            } else {
                reject(new Error(response?.error || '認証解除メッセージの送信に失敗しました'));
            }
        });
    });
}

/**
 * 現在のユーザー情報を取得
 * @param {string} token アクセストークン
 * @returns {Promise<Object>} ユーザー情報
 */
export async function getUserInfo(token) {
    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });

    if (!response.ok) {
        throw new Error('ユーザー情報の取得に失敗しました');
    }

    return response.json();
}

/**
 * トークンが有効かどうかをチェック
 * @param {string} token アクセストークン
 * @returns {Promise<boolean>}
 */
export async function validateToken(token) {
    try {
        const response = await fetch('https://www.googleapis.com/oauth2/v1/tokeninfo', {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        return response.ok;
    } catch {
        return false;
    }
}
