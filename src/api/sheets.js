// Google Sheets API連携モジュール

/**
 * シート名のバリデーション
 * Google Sheets APIで使用不可の文字やURL操作可能な文字をブロック
 * @param {string} name シート名
 * @returns {string} 安全なシート名（不正な場合は'Sheet1'を返す）
 */
function sanitizeSheetName(name) {
    if (!name || typeof name !== 'string') return 'Sheet1';
    // Google Sheets API: シート名にはある程度の記号も使えるが、単一引用符等のエスケープが必要になる場合がある。
    // 今回のエラーは全角文字を弾いているわけではないが、[\\/*?\[\]':!] 等の記号が含まれるとSheet1になる。
    // 日本語のシート名（例：シート１）がそのまま通るように、サニタイズ処理を緩めるか、URLエンコードに任せる。
    if (name.length > 100) return 'Sheet1';
    // 不正な文字を置換して返す（Sheet1で上書きしない）
    return name.replace(/[\\/*?\[\]':!]/g, '_');
}
/**
 * 日報データをスプレッドシートに保存
 * @param {string} token アクセストークン
 * @param {Object} reportData 日報データ
 * @returns {Promise<Object>} 保存結果
 */
export async function saveReport(token, reportData) {
    const { spreadsheetId, sheetName } = await chrome.storage.local.get(['spreadsheetId', 'sheetName']);

    if (!spreadsheetId) {
        throw new Error('スプレッドシートIDが設定されていません。設定画面から設定してください。');
    }

    const range = `'${sanitizeSheetName(sheetName)}'!A:J`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;

    // 日報データをスプレッドシート用の行に変換
    const rows = formatReportForSheet(reportData);

    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            values: rows
        })
    });

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'スプレッドシートへの保存に失敗しました');
    }

    return response.json();
}

/**
 * 日報データをスプレッドシートの行形式に変換
 * @param {Object} reportData 日報データ
 * @returns {Array<Array>} 行データ
 */
function formatReportForSheet(reportData) {
    const rows = [];
    const timestamp = new Date().toISOString();

    for (const entry of reportData.scheduleEntries) {
        rows.push([
            reportData.userEmail,           // A: ユーザーメール
            reportData.date,                // B: 日付
            entry.title,                    // C: 予定タイトル
            entry.start,                    // D: 開始時刻
            entry.end,                      // E: 終了時刻
            entry.calendar || 'primary',    // F: カレンダー名
            entry.aiFlag,                   // G: AI活用フラグ
            entry.aiRate,                   // H: 活用率
            entry.note || '',               // I: メモ
            timestamp                       // J: 送信日時
        ]);
    }

    return rows;
}

/**
 * スプレッドシートのヘッダー行を確認・作成
 * @param {string} token アクセストークン
 * @returns {Promise<void>}
 */
export async function ensureHeaders(token) {
    const { spreadsheetId, sheetName } = await chrome.storage.local.get(['spreadsheetId', 'sheetName']);

    if (!spreadsheetId) {
        throw new Error('スプレッドシートIDが設定されていません');
    }

    const safeSheetName = sanitizeSheetName(sheetName);
    const range = `'${safeSheetName}'!A1:J1`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`;

    try {
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!response.ok) {
            await createSheet(token, spreadsheetId, safeSheetName);
        }

        let data = { values: [] };
        if (response.ok) {
            data = await response.json();
        }

        if (!data.values || data.values.length === 0) {
            const headers = [
                ['ユーザーメール', '日付', '予定タイトル', '開始時刻', '終了時刻', 'カレンダー', 'AI活用フラグ', '活用率(%)', 'メモ', '送信日時']
            ];

            await fetch(`${url}?valueInputOption=RAW`, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: headers })
            });
        }
    } catch (error) {
        await createSheet(token, spreadsheetId, safeSheetName);

        const headers = [
            ['ユーザーメール', '日付', '予定タイトル', '開始時刻', '終了時刻', 'カレンダー', 'AI活用フラグ', '活用率(%)', 'メモ', '送信日時']
        ];

        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}?valueInputOption=RAW`, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: headers })
        });
    }
}

/**
 * スプレッドシートID設定の保存
 * @param {string} spreadsheetId スプレッドシートID
 * @param {string} sheetName シート名
 * @returns {Promise<void>}
 */
export async function saveSpreadsheetConfig(spreadsheetId, sheetName = 'Sheet1', directorySpreadsheetId = '') {
    await chrome.storage.local.set({ spreadsheetId, sheetName, directorySpreadsheetId });
}

/**
 * スプレッドシート設定の取得
 * @returns {Promise<Object>} 設定オブジェクト
 */
export async function getSpreadsheetConfig() {
    const result = await chrome.storage.local.get(['spreadsheetId', 'sheetName', 'directorySpreadsheetId']);
    return {
        spreadsheetId: result.spreadsheetId || '',
        sheetName: result.sheetName || 'Sheet1',
        directorySpreadsheetId: result.directorySpreadsheetId || ''
    };
}

/**
 * スプレッドシート接続テスト
 * @param {string} token アクセストークン
 * @returns {Promise<Object>} テスト結果
 */
export async function testSpreadsheetConnection(token) {
    const { spreadsheetId, sheetName } = await chrome.storage.local.get(['spreadsheetId', 'sheetName']);

    if (!spreadsheetId) {
        return { success: false, error: 'スプレッドシートIDが設定されていません' };
    }

    try {
        // スプレッドシートのメタデータを取得してテスト
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties.title`;

        const response = await fetch(url, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!response.ok) {
            const error = await response.json();
            const errorMsg = error.error?.message || `HTTP ${response.status}`;
            return { success: false, error: errorMsg };
        }

        const data = await response.json();
        const spreadsheetTitle = data.properties?.title || 'Untitled';
        const sheets = data.sheets?.map(s => s.properties?.title) || [];
        const targetSheet = sanitizeSheetName(sheetName);
        const sheetExists = sheets.includes(targetSheet);

        return {
            success: true,
            spreadsheetTitle,
            sheets,
            targetSheet,
            sheetExists,
            message: sheetExists
                ? `「${spreadsheetTitle}」に接続しました`
                : `「${spreadsheetTitle}」に接続しました（シート「${targetSheet}」は自動作成されます）`
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

/**
 * 日次集計シートを更新（AI分析用）
 * @param {string} token アクセストークン
 * @param {Object} reportData 日報データ
 * @returns {Promise<Object>} 更新結果
 */
export async function updateDailySummary(token, reportData) {
    const { spreadsheetId } = await chrome.storage.local.get(['spreadsheetId']);

    if (!spreadsheetId) {
        throw new Error('スプレッドシートIDが設定されていません');
    }

    const summarySheetName = 'DailySummary';

    // サマリーシートのヘッダーを確認・作成
    await ensureSummaryHeaders(token, spreadsheetId, summarySheetName);

    // 日次集計データを計算
    const summaryRow = calculateDailySummary(reportData);

    // 既存の同日・同ユーザーの行を更新または新規追加
    await upsertSummaryRow(token, spreadsheetId, summarySheetName, summaryRow);

    return { success: true };
}

/**
 * サマリーシートのヘッダーを確認・作成
 */
async function ensureSummaryHeaders(token, spreadsheetId, sheetName) {
    const range = `'${sheetName}'!A1:L1`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`;

    try {
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        // シートが存在しない場合は作成
        if (!response.ok) {
            await createSheet(token, spreadsheetId, sheetName);
        }

        const data = await response.json();

        // ヘッダーが存在しない場合は作成
        if (!data.values || data.values.length === 0) {
            const headers = [[
                '日付',
                'ユーザー',
                '予定数',
                '合計時間(分)',
                'AI活用件数',
                'AI未活用件数',
                'AI活用可能件数',
                '平均活用率(%)',
                'AI活用時間(分)',
                'AI未活用時間(分)',
                'AI活用可能時間(分)',
                '更新日時'
            ]];

            await fetch(`${url}?valueInputOption=RAW`, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: headers })
            });
        }
    } catch (error) {
        // シート作成を試行
        await createSheet(token, spreadsheetId, sheetName);

        // ヘッダー作成
        const headers = [[
            '日付',
            'ユーザー',
            '予定数',
            '合計時間(分)',
            'AI活用件数',
            'AI未活用件数',
            'AI活用可能件数',
            '平均活用率(%)',
            'AI活用時間(分)',
            'AI未活用時間(分)',
            'AI活用可能時間(分)',
            '更新日時'
        ]];

        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${sheetName}'!A1:L1?valueInputOption=RAW`, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: headers })
        });
    }
}

/**
 * シートを作成
 */
async function createSheet(token, spreadsheetId, sheetName) {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;

    await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            requests: [{
                addSheet: {
                    properties: { title: sheetName }
                }
            }]
        })
    });
}

/**
 * 日次集計を計算
 */
function calculateDailySummary(reportData) {
    const entries = reportData.scheduleEntries;

    let totalMinutes = 0;
    let aiUsingCount = 0;
    let aiNotUsingCount = 0;
    let aiPotentialCount = 0;
    let aiUsingMinutes = 0;
    let aiNotUsingMinutes = 0;
    let aiPotentialMinutes = 0;
    let totalRate = 0;
    let rateCount = 0;

    for (const entry of entries) {
        const duration = calculateDurationMinutes(entry.start, entry.end);
        totalMinutes += duration;

        if (entry.aiFlag === 'yes-using' || entry.aiFlag === 'yes-failed') {
            aiUsingCount++;
            aiUsingMinutes += duration;
            totalRate += entry.aiRate;
            rateCount++;
        } else if (entry.aiFlag === 'yes-potential') {
            aiPotentialCount++;
            aiPotentialMinutes += duration;
            totalRate += entry.aiRate;
            rateCount++;
        } else {
            aiNotUsingCount++;
            aiNotUsingMinutes += duration;
        }
    }

    const avgRate = rateCount > 0 ? Math.round(totalRate / rateCount) : 0;

    return [
        reportData.date,                    // 日付
        reportData.userEmail,               // ユーザー
        entries.length,                     // 予定数
        totalMinutes,                       // 合計時間(分)
        aiUsingCount,                       // AI活用件数
        aiNotUsingCount,                    // AI未活用件数
        aiPotentialCount,                   // AI活用可能件数
        avgRate,                            // 平均活用率(%)
        aiUsingMinutes,                     // AI活用時間(分)
        aiNotUsingMinutes,                  // AI未活用時間(分)
        aiPotentialMinutes,                 // AI活用可能時間(分)
        new Date().toISOString()            // 更新日時
    ];
}

/**
 * 時刻から所要時間（分）を計算
 */
function calculateDurationMinutes(startStr, endStr) {
    try {
        const [startH, startM] = startStr.split(':').map(Number);
        const [endH, endM] = endStr.split(':').map(Number);
        return (endH * 60 + endM) - (startH * 60 + startM);
    } catch {
        return 0;
    }
}

/**
 * サマリー行を更新または挿入（同日・同ユーザーがあれば更新）
 */
async function upsertSummaryRow(token, spreadsheetId, sheetName, newRow) {
    const range = `'${sheetName}'!A:B`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`;

    const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${token}` }
    });

    if (!response.ok) {
        // 新規追加
        await appendSummaryRow(token, spreadsheetId, sheetName, newRow);
        return;
    }

    const data = await response.json();
    const values = data.values || [];

    // 同日・同ユーザーの行を検索
    const targetDate = newRow[0];
    const targetUser = newRow[1];
    let rowIndex = -1;

    for (let i = 1; i < values.length; i++) { // ヘッダーをスキップ
        if (values[i][0] === targetDate && values[i][1] === targetUser) {
            rowIndex = i + 1; // 1-indexed
            break;
        }
    }

    if (rowIndex > 0) {
        // 既存行を更新
        const updateRange = `'${sheetName}'!A${rowIndex}:L${rowIndex}`;
        const encodedUpdateRange = encodeURIComponent(updateRange);
        const updateUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedUpdateRange}?valueInputOption=USER_ENTERED`;

        await fetch(updateUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [newRow] })
        });
    } else {
        // 新規追加
        await appendSummaryRow(token, spreadsheetId, sheetName, newRow);
    }
}

/**
 * サマリー行を追加
 */
async function appendSummaryRow(token, spreadsheetId, sheetName, row) {
    const range = `'${sheetName}'!A:L`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;

    await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [row] })
    });
}

// ========================================
// CalendarSummary シート管理
// ========================================

/**
 * カレンダー別の日次集計シートを更新
 * @param {string} token アクセストークン
 * @param {Object} reportData 日報データ
 * @returns {Promise<Object>} 更新結果
 */
export async function updateCalendarSummary(token, reportData) {
    const { spreadsheetId } = await chrome.storage.local.get(['spreadsheetId']);

    if (!spreadsheetId) {
        throw new Error('スプレッドシートIDが設定されていません');
    }

    const sheetName = 'CalendarSummary';

    // ヘッダー確認・作成
    await ensureCalendarSummaryHeaders(token, spreadsheetId, sheetName);

    // カレンダー別にグルーピング
    const calendarGroups = {};
    for (const entry of reportData.scheduleEntries) {
        const calName = entry.calendar || 'primary';
        if (!calendarGroups[calName]) {
            calendarGroups[calName] = [];
        }
        calendarGroups[calName].push(entry);
    }

    // 各カレンダーの集計データを書き込み
    for (const [calName, entries] of Object.entries(calendarGroups)) {
        const row = calculateCalendarSummaryRow(reportData.date, reportData.userEmail, calName, entries);
        await upsertCalendarSummaryRow(token, spreadsheetId, sheetName, row);
    }

    return { success: true };
}

/**
 * CalendarSummaryシートのヘッダーを確認・作成
 */
async function ensureCalendarSummaryHeaders(token, spreadsheetId, sheetName) {
    const range = `'${sheetName}'!A1:M1`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`;

    try {
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!response.ok) {
            await createSheet(token, spreadsheetId, sheetName);
        }

        const data = await response.json();

        if (!data.values || data.values.length === 0) {
            const headers = [[
                '日付',
                'ユーザー',
                'カレンダー名',
                '予定数',
                '合計時間(分)',
                'AI活用件数',
                'AI未活用件数',
                'AI活用可能件数',
                '平均活用率(%)',
                'AI活用時間(分)',
                'AI未活用時間(分)',
                'AI活用可能時間(分)',
                '更新日時'
            ]];

            await fetch(`${url}?valueInputOption=RAW`, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: headers })
            });
        }
    } catch (error) {
        await createSheet(token, spreadsheetId, sheetName);

        const headers = [[
            '日付', 'ユーザー', 'カレンダー名', '予定数', '合計時間(分)',
            'AI活用件数', 'AI未活用件数', 'AI活用可能件数', '平均活用率(%)',
            'AI活用時間(分)', 'AI未活用時間(分)', 'AI活用可能時間(分)', '更新日時'
        ]];

        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${sheetName}'!A1:M1?valueInputOption=RAW`, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: headers })
        });
    }
}

/**
 * カレンダー別の集計行を計算
 */
function calculateCalendarSummaryRow(date, userEmail, calendarName, entries) {
    let totalMinutes = 0;
    let aiUsingCount = 0;
    let aiNotUsingCount = 0;
    let aiPotentialCount = 0;
    let aiUsingMinutes = 0;
    let aiNotUsingMinutes = 0;
    let aiPotentialMinutes = 0;
    let totalRate = 0;
    let rateCount = 0;

    for (const entry of entries) {
        const duration = calculateDurationMinutes(entry.start, entry.end);
        totalMinutes += duration;

        if (entry.aiFlag === 'yes-using' || entry.aiFlag === 'yes-failed') {
            aiUsingCount++;
            aiUsingMinutes += duration;
            totalRate += entry.aiRate;
            rateCount++;
        } else if (entry.aiFlag === 'yes-potential') {
            aiPotentialCount++;
            aiPotentialMinutes += duration;
            totalRate += entry.aiRate;
            rateCount++;
        } else {
            aiNotUsingCount++;
            aiNotUsingMinutes += duration;
        }
    }

    const avgRate = rateCount > 0 ? Math.round(totalRate / rateCount) : 0;

    return [
        date,                               // 日付
        userEmail,                          // ユーザー
        calendarName,                       // カレンダー名
        entries.length,                     // 予定数
        totalMinutes,                       // 合計時間(分)
        aiUsingCount,                       // AI活用件数
        aiNotUsingCount,                    // AI未活用件数
        aiPotentialCount,                   // AI活用可能件数
        avgRate,                            // 平均活用率(%)
        aiUsingMinutes,                     // AI活用時間(分)
        aiNotUsingMinutes,                  // AI未活用時間(分)
        aiPotentialMinutes,                 // AI活用可能時間(分)
        new Date().toISOString()            // 更新日時
    ];
}

/**
 * CalendarSummary行を更新または挿入（同日・同ユーザー・同カレンダー）
 */
async function upsertCalendarSummaryRow(token, spreadsheetId, sheetName, newRow) {
    const range = `'${sheetName}'!A:C`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`;

    const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${token}` }
    });

    if (!response.ok) {
        await appendCalendarSummaryRow(token, spreadsheetId, sheetName, newRow);
        return;
    }

    const data = await response.json();
    const values = data.values || [];

    const targetDate = newRow[0];
    const targetUser = newRow[1];
    const targetCal = newRow[2];
    let rowIndex = -1;

    for (let i = 1; i < values.length; i++) {
        if (values[i][0] === targetDate && values[i][1] === targetUser && values[i][2] === targetCal) {
            rowIndex = i + 1;
            break;
        }
    }

    if (rowIndex > 0) {
        const updateRange = `'${sheetName}'!A${rowIndex}:M${rowIndex}`;
        const encodedUpdateRange = encodeURIComponent(updateRange);
        const updateUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedUpdateRange}?valueInputOption=USER_ENTERED`;

        await fetch(updateUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [newRow] })
        });
    } else {
        await appendCalendarSummaryRow(token, spreadsheetId, sheetName, newRow);
    }
}

/**
 * CalendarSummary行を追加
 */
async function appendCalendarSummaryRow(token, spreadsheetId, sheetName, row) {
    const range = `'${sheetName}'!A:M`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;

    await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [row] })
    });
}

// ========================================
// WeeklyReport / MonthlyReport シート保存
// ========================================

/**
 * レポートをスプレッドシートに保存
 * @param {string} token アクセストークン
 * @param {string} type 'weekly' or 'monthly'
 * @param {Object} summary 集計データ
 * @param {string} analysis AI分析結果テキスト
 * @returns {Promise<Object>} 保存結果
 */
export async function saveReportToSheet(token, type, summary, analysis) {
    const { spreadsheetId } = await chrome.storage.local.get(['spreadsheetId']);

    if (!spreadsheetId) {
        throw new Error('スプレッドシートIDが設定されていません');
    }

    const sheetName = type === 'weekly' ? 'WeeklyReport' : 'MonthlyReport';

    // ヘッダー確認・作成
    await ensureReportHeaders(token, spreadsheetId, sheetName);

    // レポート行を作成
    const row = [
        new Date().toISOString(),                                   // 生成日時
        summary.period.start,                                       // 期間開始
        summary.period.end,                                         // 期間終了
        summary.period.days,                                        // 日数
        1,                                                          // ユーザー数
        summary.totals.events,                                      // 予定数
        Math.round(summary.totals.totalMinutes / 60),               // 稼働時間(h)
        summary.totals.aiUsingCount,                                // AI活用件数
        summary.totals.aiNotUsingCount,                             // AI未活用件数
        summary.totals.aiPotentialCount,                            // AI余地件数
        summary.averages.aiRate,                                    // AI活用率(%)
        analysis                                                    // AI分析レポート
    ];

    // シートに追加
    const range = `'${sheetName}'!A:L`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;

    await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [row] })
    });

    return { success: true };
}

/**
 * レポートシートのヘッダーを確認・作成
 */
async function ensureReportHeaders(token, spreadsheetId, sheetName) {
    const range = `'${sheetName}'!A1:L1`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`;

    try {
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!response.ok) {
            await createSheet(token, spreadsheetId, sheetName);
        }

        const data = await response.json();

        if (!data.values || data.values.length === 0) {
            const headers = [[
                '生成日時', '期間開始', '期間終了', '日数', 'ユーザー数',
                '予定数', '稼働時間(h)', 'AI活用件数', 'AI未活用件数',
                'AI余地件数', 'AI活用率(%)', 'AI分析レポート'
            ]];

            await fetch(`${url}?valueInputOption=RAW`, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: headers })
            });
        }
    } catch (error) {
        await createSheet(token, spreadsheetId, sheetName);

        const headers = [[
            '生成日時', '期間開始', '期間終了', '日数', 'ユーザー数',
            '予定数', '稼働時間(h)', 'AI活用件数', 'AI未活用件数',
            'AI余地件数', 'AI活用率(%)', 'AI分析レポート'
        ]];

        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${sheetName}'!A1:L1?valueInputOption=RAW`, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: headers })
        });
    }
}

// ========================================
// Wizard Reports シート保存 (ウィザード連携)
// ========================================

/**
 * ウィザードの日報テキストをスプレッドシートに保存
 * @param {string} token アクセストークン
 * @param {string} userEmail ユーザーメールアドレス
 * @param {string} reportText 日報テキスト
 * @returns {Promise<Object>} 保存結果
 */
export async function saveWizardReportToSheet(token, userEmail, reportText) {
    const { spreadsheetId } = await chrome.storage.local.get(['spreadsheetId']);

    if (!spreadsheetId) {
        throw new Error('スプレッドシートIDが設定されていません');
    }

    const sheetName = 'WizardReports';

    // ヘッダー確認・作成
    await ensureWizardReportHeaders(token, spreadsheetId, sheetName);

    // レポート行を作成
    const row = [
        new Date().toISOString(),                                   // 送信日時
        userEmail || 'unknown',                                     // ユーザーメール
        reportText                                                  // 日報テキスト
    ];

    // シートに追加
    const range = `'${sheetName}'!A:C`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;

    await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [row] })
    });

    return { success: true };
}

// ========================================
// AI Impact Log シート保存 (ウィザード連携)
// ========================================

/**
 * AI活用実績（インパクト等）を構造化データとして保存
 * @param {string} token アクセストークン
 * @param {string} userEmail ユーザーメールアドレス
 * @param {Array} aiUsageData AI活用データの配列
 * @param {string} valueCreation 価値創造アクション
 * @returns {Promise<Object>} 保存結果
 */
export async function saveAiImpactLog(token, userEmail, aiUsageData, valueCreation) {
    const { spreadsheetId } = await chrome.storage.local.get(['spreadsheetId']);

    if (!spreadsheetId) {
        throw new Error('スプレッドシートIDが設定されていません');
    }

    const sheetName = 'AiImpactLog';

    // ヘッダー確認・作成
    await ensureAiImpactHeaders(token, spreadsheetId, sheetName);

    // AI活用したタスク（used === true）のみ抽出して行データを作成
    const rows = aiUsageData.filter(a => a.used).map(a => {
        return [
            new Date().toISOString(),                                   // 日時
            userEmail || 'unknown',                                     // ユーザー
            a.task || '',                                               // タスク名
            a.tool || '',                                               // 使用ツール
            a.oneliner || '',                                           // 一言
            a.impactScore ? a.impactScore.toString() : '',              // インパクトスコア(1-5)
            a.timeSaved ? a.timeSaved.toString() : '',                  // 短縮時間(分)
            valueCreation || ''                                         // 価値創造アクション
        ];
    });

    if (rows.length === 0) {
        return { success: true, message: 'no_ai_usage' };
    }

    // シートに追加
    const range = `'${sheetName}'!A:H`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;

    await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: rows })
    });

    return { success: true, count: rows.length };
}

/**
 * WizardReportsシートのヘッダーを確認・作成
 */
async function ensureWizardReportHeaders(token, spreadsheetId, sheetName) {
    const range = `'${sheetName}'!A1:C1`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`;

    try {
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!response.ok) {
            await createSheet(token, spreadsheetId, sheetName);
        }

        const data = await response.json();

        if (!data.values || data.values.length === 0) {
            const headers = [['送信日時', 'ユーザーメール', '日報テキスト']];

            await fetch(`${url}?valueInputOption=RAW`, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: headers })
            });
        }
    } catch (error) {
        await createSheet(token, spreadsheetId, sheetName);

        const headers = [['送信日時', 'ユーザーメール', '日報テキスト']];

        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${sheetName}'!A1:C1?valueInputOption=RAW`, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: headers })
        });
    }
}

/**
 * AiImpactLogシートのヘッダーを確認・作成
 */
async function ensureAiImpactHeaders(token, spreadsheetId, sheetName) {
    const range = `'${sheetName}'!A1:H1`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`;

    const headers = [['日時', 'ユーザー', 'タスク名', '使用ツール', '一言', 'インパクトスコア(1-5)', '短縮時間(分)', '価値創造アクション']];

    try {
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!response.ok) {
            await createSheet(token, spreadsheetId, sheetName);
        }

        const data = await response.json();

        if (!data.values || data.values.length === 0) {
            await fetch(`${url}?valueInputOption=RAW`, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: headers })
            });
        }
    } catch (error) {
        await createSheet(token, spreadsheetId, sheetName);

        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${sheetName}'!A1:H1?valueInputOption=RAW`, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: headers })
        });
    }
}

// ========================================
// Decision Log シート保存 (判断の記録)
// ========================================

/**
 * 判断ログをスプレッドシートに保存
 * @param {string} token アクセストークン
 * @param {string} userEmail ユーザーメールアドレス
 * @param {Array} decisions 判断データの配列 [{task, tags[], memo}]
 * @returns {Promise<Object>} 保存結果
 */
export async function saveDecisionLog(token, userEmail, decisions) {
    const { spreadsheetId } = await chrome.storage.local.get(['spreadsheetId']);

    if (!spreadsheetId) {
        throw new Error('スプレッドシートIDが設定されていません');
    }

    const sheetName = 'DecisionLog';

    // ヘッダー確認・作成
    await ensureDecisionLogHeaders(token, spreadsheetId, sheetName);

    // 行データを作成
    const rows = decisions.map(d => {
        return [
            new Date().toISOString(),              // 日時
            userEmail || 'unknown',                 // ユーザー
            d.task || '',                           // 予定名
            (d.tags || []).join(', '),               // 判断タグ（カンマ区切り）
            d.memo || '',                           // 一言メモ
        ];
    });

    if (rows.length === 0) {
        return { success: true, message: 'no_decisions' };
    }

    // シートに追加
    const range = `'${sheetName}'!A:E`;
    const encodedRange = encodeURIComponent(range);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;

    await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: rows })
    });

    return { success: true, count: rows.length };
}

/**
 * DecisionLogシートのヘッダーを確認・作成
 */
async function ensureDecisionLogHeaders(token, spreadsheetId, sheetName) {
    const headers = [['日時', 'ユーザー', '予定名', '判断タグ', '一言メモ']];

    try {
        const res = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${sheetName}'!A1:E1`,
            { headers: { 'Authorization': `Bearer ${token}` } }
        );

        if (!res.ok) {
            throw new Error('Sheet not found');
        }

        const data = await res.json();
        if (!data.values || data.values.length === 0) {
            await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${sheetName}'!A1:E1?valueInputOption=RAW`, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ values: headers })
            });
        }
    } catch (error) {
        await createSheet(token, spreadsheetId, sheetName);

        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${sheetName}'!A1:E1?valueInputOption=RAW`, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: headers })
        });
    }
}
