// Google Calendar API連携モジュール

/**
 * 指定した日のカレンダーイベントを取得
 * @param {string} token アクセストークン
 * @param {Date} date 対象日
 * @returns {Promise<Array>} イベント一覧
 */
export async function getEventsForDate(token, date) {
    const startOfDay = new Date(date);
    startOfDay.setHours(0, 0, 0, 0);

    const endOfDay = new Date(date);
    endOfDay.setHours(23, 59, 59, 999);

    const params = new URLSearchParams({
        timeMin: startOfDay.toISOString(),
        timeMax: endOfDay.toISOString(),
        singleEvents: 'true',
        orderBy: 'startTime',
        maxResults: '50'
    });

    // 選択されたカレンダーを取得
    const { selectedCalendars } = await chrome.storage.sync.get(['selectedCalendars']);

    // カレンダー一覧を取得
    const calendars = await getCalendarList(token);
    const allEvents = [];

    // 選択されたカレンダー（未設定の場合は全て）からイベントを取得
    for (const calendar of calendars) {
        // freeBusyReaderは予定詳細を取得できないのでスキップ
        if (calendar.accessRole === 'freeBusyReader') continue;

        // 選択されたカレンダーのみ取得（未設定なら全て取得）
        if (selectedCalendars && selectedCalendars.length > 0) {
            if (!selectedCalendars.includes(calendar.id)) continue;
        }

        try {
            const calendarEvents = await fetchCalendarEvents(token, calendar.id, params);
            // カレンダー情報を各イベントに追加
            calendarEvents.forEach(event => {
                event.calendarName = calendar.summary;
                event.calendarColor = calendar.backgroundColor;
                event.calendarId = calendar.id;
            });
            allEvents.push(...calendarEvents);
        } catch (error) {
            console.warn(`カレンダー ${calendar.summary} の取得に失敗:`, error);
        }
    }

    // 開始時刻でソート
    allEvents.sort((a, b) => {
        const aStart = new Date(a.start.dateTime || a.start.date);
        const bStart = new Date(b.start.dateTime || b.start.date);
        return aStart - bStart;
    });

    return allEvents;
}

/**
 * カレンダー一覧を取得
 * @param {string} token アクセストークン
 * @returns {Promise<Array>} カレンダー一覧
 */
export async function getCalendarList(token) {
    const response = await fetch('https://www.googleapis.com/calendar/v3/users/me/calendarList', {
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'カレンダー一覧の取得に失敗しました');
    }

    const data = await response.json();
    return data.items || [];
}

/**
 * カレンダー選択設定を保存
 * @param {Array<string>} calendarIds 選択されたカレンダーID
 */
export async function saveSelectedCalendars(calendarIds) {
    await chrome.storage.sync.set({ selectedCalendars: calendarIds });
}

/**
 * カレンダー選択設定を取得
 * @returns {Promise<Array<string>>} 選択されたカレンダーID
 */
export async function getSelectedCalendars() {
    const { selectedCalendars } = await chrome.storage.sync.get(['selectedCalendars']);
    return selectedCalendars || [];
}

/**
 * 指定カレンダーのイベントを取得
 * @param {string} token アクセストークン
 * @param {string} calendarId カレンダーID
 * @param {URLSearchParams} params クエリパラメータ
 * @returns {Promise<Array>} イベント一覧
 */
async function fetchCalendarEvents(token, calendarId, params) {
    const url = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events?${params}`;

    const response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'イベントの取得に失敗しました');
    }

    const data = await response.json();
    return data.items || [];
}

/**
 * イベントの所要時間を計算（分）
 * @param {Object} event イベント
 * @returns {number} 所要時間（分）
 */
export function getEventDuration(event) {
    const start = new Date(event.start.dateTime || event.start.date);
    const end = new Date(event.end.dateTime || event.end.date);
    return Math.round((end - start) / (1000 * 60));
}

/**
 * 時刻をフォーマット
 * @param {string} dateTimeString ISO日時文字列
 * @returns {string} HH:MM形式
 */
export function formatTime(dateTimeString) {
    const date = new Date(dateTimeString);
    return date.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' });
}

/**
 * 合計稼働時間を計算
 * @param {Array} events イベント一覧
 * @returns {number} 合計時間（分）
 */
export function calculateTotalWorkTime(events) {
    return events.reduce((total, event) => total + getEventDuration(event), 0);
}

/**
 * 分を時間:分形式に変換
 * @param {number} minutes 分
 * @returns {string} X時間Y分形式
 */
export function formatDuration(minutes) {
    const hours = Math.floor(minutes / 60);
    const mins = minutes % 60;
    if (hours === 0) {
        return `${mins}分`;
    }
    if (mins === 0) {
        return `${hours}時間`;
    }
    return `${hours}時間${mins}分`;
}
