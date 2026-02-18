// OAuth認証モジュール

/**
 * Google OAuth認証を実行してトークンを取得
 * @returns {Promise<string>} アクセストークン
 */
export async function getAuthToken() {
    return new Promise((resolve, reject) => {
        chrome.runtime.sendMessage({ action: 'getAuthToken' }, (response) => {
            if (response.success) {
                resolve(response.token);
            } else {
                reject(new Error(response.error));
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
            if (response.success) {
                resolve();
            } else {
                reject(new Error(response.error));
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
        const response = await fetch(`https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=${token}`);
        return response.ok;
    } catch {
        return false;
    }
}
