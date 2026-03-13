// HTMLサニタイズユーティリティ

/**
 * HTMLの特殊文字をエスケープ（XSS対策）
 * @param {string} str エスケープ対象の文字列
 * @returns {string} エスケープ済みの文字列
 */
export function escapeHtml(str) {
    if (str == null) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

/**
 * マークダウンテキストを安全なHTMLに変換（XSS対策済み）
 * Gemini APIなど外部ソースからのマークダウンを表示する際に使用
 * @param {string} markdown マークダウンテキスト
 * @returns {string} サニタイズ済みHTML
 */
export function markdownToSafeHtml(markdown) {
    if (!markdown) return '';

    // まず全体をエスケープ
    let html = escapeHtml(markdown);

    // 安全なマークダウン変換のみ適用
    html = html
        .replace(/^### (.+)$/gm, '<h4>$1</h4>')
        .replace(/^## (.+)$/gm, '<h3>$1</h3>')
        .replace(/^# (.+)$/gm, '<h2>$1</h2>')
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
        .replace(/^\d+\. /gm, (match) => '<br>' + match)
        .replace(/^- /gm, '• ')
        .replace(/\n- /g, '<br>• ')
        .replace(/\n\n/g, '<br><br>')
        .replace(/\n/g, '<br>');

    return html;
}

/**
 * プロンプトインジェクション対策（Gemini API等へ送信するユーザー入力の安全検証）
 * @param {string} text 入力テキスト
 * @param {number} maxLen 最大文字数
 * @returns {string} サニタイズ済みの文字列
 */
export function sanitizeForPrompt(text, maxLen = 200) {
    if (!text || typeof text !== 'string') return '';
    return text.replace(/[\r\n]+/g, ' ').replace(/[\x00-\x1F\x7F]/g, '').replace(/[\[\]\{\}"\\]/g, '').slice(0, maxLen);
}
