// AI活用度測定 Chrome拡張機能 - バックグラウンドサービスワーカー
import { sanitizeForPrompt } from './src/utils/sanitize.js';
// サイドパネルを開く設定
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch((error) => console.error('サイドパネル設定エラー:', error));

// storage.sync → storage.local へのワンタイム移行
(async function migrateStorageSyncToLocal() {
  const { storageMigrated } = await chrome.storage.local.get('storageMigrated');
  if (storageMigrated) return;

  try {
    const syncData = await chrome.storage.sync.get(null);
    const keysToMigrate = [
      'geminiApiKey', 'notionToken', 'notionDatabaseId',
      'spreadsheetId', 'sheetName', 'enableSheets', 'enableNotion',
      'enableAiAnalysis', 'userSettings'
    ];

    const dataToMigrate = {};
    for (const key of keysToMigrate) {
      if (syncData[key] !== undefined) {
        dataToMigrate[key] = syncData[key];
      }
    }

    if (Object.keys(dataToMigrate).length > 0) {
      await chrome.storage.local.set(dataToMigrate);
      await chrome.storage.sync.remove(keysToMigrate); // セキュリティ修正: 移行済みのデータをsyncから完全に削除
      console.log('ストレージ移行完了: sync → local に移動し、sync のデータを削除しました', Object.keys(dataToMigrate));
    }

    await chrome.storage.local.set({ storageMigrated: true });
  } catch (error) {
    console.error('ストレージ移行エラー:', error.message);
  }

  // 不要になったストレージキーのクリーンアップ
  try {
    const staleKeys = ['userSettings', 'lastSubmittedDate'];
    const existing = await chrome.storage.local.get(staleKeys);
    const keysToRemove = staleKeys.filter(k => existing[k] !== undefined);
    if (keysToRemove.length > 0) {
      await chrome.storage.local.remove(keysToRemove);
      console.log('不要なストレージキーを削除しました:', keysToRemove);
    }
  } catch (error) {
    console.warn('不要キー削除エラー:', error.message);
  }
})();

// メッセージリスナー
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'getAuthToken') {
    const interactive = message.interactive !== undefined ? message.interactive : true;
    handleGetAuthToken(interactive, sendResponse);
    return true; // 非同期レスポンスのため
  }

  if (message.action === 'updateAlarms') {
    setupReportAlarms();
    setupDailyAlarm();
    sendResponse({ success: true });
    return false;
  }

  if (message.action === 'generateDailySlack') {
    // 手動で日報生成・Slack送信を実行
    runDailyReportCheck(true).then(result => {
      sendResponse(result);
    }).catch(e => {
      sendResponse({ success: false, error: e.message });
    });
    return true; // 非同期レスポンス
  }

  if (message.action === 'revokeToken') {
    handleRevokeToken(sendResponse);
    return true;
  }

  if (message.action === 'retryOfflineQueue') {
    handleRetryOfflineQueue(sendResponse);
    return true;
  }

  if (message.action === 'notionRequest') {
    handleNotionRequest(message, sendResponse);
    return true;
  }
});

// OAuth認証トークン取得
async function handleGetAuthToken(interactive, sendResponse) {
  try {
    const token = await new Promise((resolve, reject) => {
      chrome.identity.getAuthToken({ interactive: interactive }, (token) => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
        } else {
          resolve(token);
        }
      });
    });
    sendResponse({ success: true, token });
  } catch (error) {
    console.error('認証エラー詳細:', error);
    sendResponse({ success: false, error: '認証に失敗しました。再度ログインしてください。' });
  }
}

// トークン失効処理
async function handleRevokeToken(sendResponse) {
  try {
    const token = await new Promise((resolve) => {
      chrome.identity.getAuthToken({ interactive: false }, resolve);
    });

    if (token) {
      await new Promise((resolve) => {
        chrome.identity.removeCachedAuthToken({ token }, resolve);
      });

      // Googleのトークン失効エンドポイントを呼び出し（POST方式）
      await fetch('https://accounts.google.com/o/oauth2/revoke', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: `token=${encodeURIComponent(token)}`
      });
    }

    sendResponse({ success: true });
  } catch (error) {
    console.error('ログアウトエラー詳細:', error);
    sendResponse({ success: false, error: 'ログアウトに失敗しました。再試行してください。' });
  }
}

// オフラインキュー再送処理
async function handleRetryOfflineQueue(sendResponse) {
  try {
    const { offlineQueue = [] } = await chrome.storage.local.get('offlineQueue');

    if (offlineQueue.length === 0) {
      sendResponse({ success: true, message: 'キューは空です' });
      return;
    }

    const results = [];
    const failedItems = [];

    for (const item of offlineQueue) {
      try {
        const response = await sendToSpreadsheet(item.data);
        if (response.success) {
          results.push({ id: item.id, success: true });
        } else {
          failedItems.push(item);
          results.push({ id: item.id, success: false, error: response.error });
        }
      } catch (error) {
        failedItems.push(item);
        console.error('アイテム再送エラー詳細:', error);
        results.push({ id: item.id, success: false, error: '送信に失敗しました' });
      }
    }

    // 失敗したアイテムをキューに戻す
    await chrome.storage.local.set({ offlineQueue: failedItems });

    sendResponse({ success: true, results, remaining: failedItems.length });
  } catch (error) {
    console.error('キュー再送エラー詳細:', error);
    sendResponse({ success: false, error: 'データの再送信に失敗しました。' });
  }
}

// スプレッドシートへの送信（background.jsから呼び出し用）
async function sendToSpreadsheet(data) {
  const { spreadsheetId, sheetName } = await chrome.storage.local.get(['spreadsheetId', 'sheetName']);

  if (!spreadsheetId) {
    throw new Error('スプレッドシートIDが設定されていません');
  }

  const token = await new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: false }, (token) => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        resolve(token);
      }
    });
  });

  // スプレッドシートIDのフォーマット検証
  if (!/^[a-zA-Z0-9_-]+$/.test(spreadsheetId)) {
    throw new Error('スプレッドシートIDのフォーマットが不正です');
  }

  // シート名のバリデーション
  const safeName = (sheetName && typeof sheetName === 'string' && !/[\\/*?\[\]':!]/.test(sheetName) && sheetName.length <= 100) ? sheetName : 'Sheet1';

  const range = `'${safeName}'!A:J`;
  const encodedRange = encodeURIComponent(range);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      values: [data]
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'スプレッドシート送信エラー');
  }

  return { success: true };
}

// ネットワーク状態の監視とオフラインキューの自動再送
chrome.runtime.onStartup.addListener(async () => {
  // 拡張機能起動時にオフラインキューを確認
  const { offlineQueue = [] } = await chrome.storage.local.get('offlineQueue');
  if (offlineQueue.length > 0) {
    console.log(`オフラインキュー: ${offlineQueue.length}件の未送信データがあります`);
  }
});

// Notion APIリクエストのプロキシ
async function handleNotionRequest(message, sendResponse) {
  try {
    const { url, method, headers, body } = message;

    // URLホワイトリスト検証（SSRF対策の強化）
    let parsedUrl;
    try {
      parsedUrl = new URL(url);
    } catch (e) {
      sendResponse({ success: false, error: '不正なURL形式です' });
      return;
    }

    if (parsedUrl.hostname !== 'api.notion.com') {
      sendResponse({ success: false, error: '許可されていないURLです' });
      return;
    }

    // SSRF対策: 許可するエンドポイントのパスを限定
    const allowedPaths = [
      '/v1/pages',
      new RegExp('^/v1/databases/[^/]+$') // /v1/databases/{id}
    ];

    const isAllowedPath = allowedPaths.some(path => {
      if (typeof path === 'string') return parsedUrl.pathname === path;
      return path.test(parsedUrl.pathname);
    });

    if (!isAllowedPath) {
      sendResponse({ success: false, error: '許可されていないAPIエンドポイントです' });
      return;
    }

    // HTTPメソッド許可リスト（任意文字列注入防止）
    const ALLOWED_METHODS = ['GET', 'POST'];
    const safeMethod = ALLOWED_METHODS.includes((method || 'GET').toUpperCase())
      ? (method || 'GET').toUpperCase()
      : 'GET';

    const response = await fetch(url, {
      method: safeMethod,
      headers: headers || {},
      body: body ? JSON.stringify(body) : undefined
    });


    const data = await response.json();

    if (!response.ok) {
      // Notion APIのエラー詳細を取得
      const errorMessage = data.message || data.error?.message || `HTTP ${response.status}`;
      const errorCode = data.code || data.status || response.status;
      console.error('Notion API Error:', errorCode);

      sendResponse({
        success: false,
        error: 'Notion APIリクエストに失敗しました。',
        status: response.status
      });
    } else {
      sendResponse({ success: true, data });
    }
  } catch (error) {
    console.error('Notion APIリクエストエラー詳細:', error.message);
    sendResponse({ success: false, error: 'Notion APIとの通信に失敗しました。' });
  }
}

// ========================================
// 定期レポート自動生成（chrome.alarms）
// ========================================

// アラームの設定
function setupReportAlarms() {
  // 既存アラームを確認してなければ作成
  chrome.alarms.get('weeklyReport', (alarm) => {
    if (!alarm) {
      // 次の月曜9:00を計算
      const now = new Date();
      const nextMonday = new Date(now);
      const dayOfWeek = now.getDay();
      const daysUntilMonday = dayOfWeek === 0 ? 1 : dayOfWeek === 1 ? 7 : 8 - dayOfWeek;
      nextMonday.setDate(now.getDate() + daysUntilMonday);
      nextMonday.setHours(9, 0, 0, 0);

      chrome.alarms.create('weeklyReport', {
        when: nextMonday.getTime(),
        periodInMinutes: 7 * 24 * 60 // 毎週
      });
      console.log('週次レポートアラーム設定:', nextMonday.toLocaleString());
    }
  });

  chrome.alarms.get('monthlyReport', (alarm) => {
    if (!alarm) {
      // 次の月初1日9:00を計算
      const now = new Date();
      const nextFirst = new Date(now.getFullYear(), now.getMonth() + 1, 1, 9, 0, 0);

      chrome.alarms.create('monthlyReport', {
        when: nextFirst.getTime(),
        periodInMinutes: 30 * 24 * 60 // 約1ヶ月
      });
      console.log('月次レポートアラーム設定:', nextFirst.toLocaleString());
    }
  });
}

// 日次チェック用のアラーム設定
async function setupDailyAlarm() {
  const config = await chrome.storage.local.get(['enableDailyAlarm', 'alarmTime']);

  // アラームのクリア
  await chrome.alarms.clear('dailyReportCheck');

  if (config.enableDailyAlarm) {
    let hour = 18;
    let minute = 0;

    if (config.alarmTime !== undefined) {
      if (typeof config.alarmTime === 'string' && config.alarmTime.includes(':')) {
        const parts = config.alarmTime.split(':');
        hour = parseInt(parts[0], 10);
        minute = parseInt(parts[1], 10) || 0;
      } else {
        hour = parseInt(config.alarmTime, 10) || 18;
      }
    }

    const now = new Date();
    const nextAlarm = new Date(now);
    nextAlarm.setHours(hour, minute, 0, 0);

    // すでに時間を過ぎている場合は翌日に設定
    if (now.getTime() >= nextAlarm.getTime()) {
      nextAlarm.setDate(nextAlarm.getDate() + 1);
    }

    chrome.alarms.create('dailyReportCheck', {
      when: nextAlarm.getTime(),
      periodInMinutes: 24 * 60 // 毎日
    });

    const timeStr = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
    console.log(`日次チェックアラーム設定: 毎日 ${timeStr} (${nextAlarm.toLocaleString()})`);
  } else {
    console.log('日次チェックアラームは無効化されています');
  }
}

// アラーム発火時の処理
chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name === 'weeklyReport') {
    console.log('週次レポート自動生成開始');
    await generateScheduledReport('weekly');
  } else if (alarm.name === 'monthlyReport') {
    console.log('月次レポート自動生成開始');
    await generateScheduledReport('monthly');
  } else if (alarm.name === 'dailyReportCheck') {
    console.log('日報未提出チェック・自動生成開始');
    await runDailyReportCheck();
  }
});

// バックグラウンドでのレポート生成
async function generateScheduledReport(type) {
  try {
    // 設定を取得
    const { geminiApiKey, spreadsheetId } = await chrome.storage.local.get(['geminiApiKey', 'spreadsheetId']);

    if (!geminiApiKey || !spreadsheetId) {
      console.log('APIキーまたはスプレッドシートが未設定のためスキップ');
      return;
    }

    // トークン取得
    const token = await new Promise((resolve, reject) => {
      chrome.identity.getAuthToken({ interactive: false }, (token) => {
        if (chrome.runtime.lastError) reject(chrome.runtime.lastError);
        else resolve(token);
      });
    });

    if (!token) {
      console.log('認証トークンが取得できないためスキップ');
      return;
    }

    // DailySummaryからデータを取得
    const daysAgo = type === 'weekly' ? 7 : 30;
    const now = new Date();
    const startDate = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);

    const summaryData = await fetchSummaryFromSheet(token, spreadsheetId, startDate, now, type);

    if (!summaryData || summaryData.totals.events === 0) {
      console.log(`${type}レポート: データが不足のためスキップ`);
      return;
    }

    // Gemini APIでレポート生成
    const analysis = await callGeminiBackground(geminiApiKey, summaryData, type);

    // シートに保存
    await saveReportBackground(token, spreadsheetId, type, summaryData, analysis);

    // 通知を表示
    chrome.notifications.create(`${type}-report-${Date.now()}`, {
      type: 'basic',
      iconUrl: 'icons/icon128.svg',
      title: `${type === 'weekly' ? '週次' : '月次'}AI活用レポート`,
      message: `${type === 'weekly' ? '週次' : '月次'}レポートを自動生成しスプレッドシートに保存しました。`
    });

    console.log(`${type}レポート自動生成完了`);
  } catch (error) {
    console.error(`${type}レポート自動生成エラー:`, error);
  }
}

// シートからサマリーデータを取得（バックグラウンド用）
async function fetchSummaryFromSheet(token, spreadsheetId, startDate, endDate, type) {
  try {
    const range = 'DailySummary!A2:K';
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;

    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${token}` }
    });

    if (!response.ok) return null;

    const data = await response.json();
    if (!data.values) return null;

    const rows = data.values
      .map(row => ({
        date: row[0],
        user: row[1],
        eventCount: parseInt(row[2]) || 0,
        totalMinutes: parseInt(row[3]) || 0,
        aiUsingCount: parseInt(row[4]) || 0,
        aiNotUsingCount: parseInt(row[5]) || 0,
        aiPotentialCount: parseInt(row[6]) || 0,
        avgRate: parseInt(row[7]) || 0,
        aiUsingMinutes: parseInt(row[8]) || 0,
        aiNotUsingMinutes: parseInt(row[9]) || 0,
        aiPotentialMinutes: parseInt(row[10]) || 0
      }))
      .filter(row => {
        const d = new Date(row.date);
        return !isNaN(d.getTime()) && d >= startDate && d <= endDate;
      });

    if (rows.length === 0) return null;

    const totals = rows.reduce((acc, row) => {
      acc.events += row.eventCount;
      acc.totalMinutes += row.totalMinutes;
      acc.aiUsingCount += row.aiUsingCount;
      acc.aiNotUsingCount += row.aiNotUsingCount;
      acc.aiPotentialCount += row.aiPotentialCount;
      acc.aiUsingMinutes += row.aiUsingMinutes;
      acc.aiNotUsingMinutes += row.aiNotUsingMinutes;
      acc.aiPotentialMinutes += row.aiPotentialMinutes;
      return acc;
    }, {
      events: 0, totalMinutes: 0, aiUsingCount: 0, aiNotUsingCount: 0,
      aiPotentialCount: 0, aiUsingMinutes: 0, aiNotUsingMinutes: 0, aiPotentialMinutes: 0
    });

    const uniqueDays = new Set(rows.map(r => r.date)).size;
    const totalAiRelated = totals.aiUsingCount + totals.aiNotUsingCount + totals.aiPotentialCount;
    const aiRate = totalAiRelated > 0 ? Math.round(totals.aiUsingCount / totalAiRelated * 100) : 0;

    // CalendarSummaryも取得
    let calendarSummary = [];
    try {
      const calRange = 'CalendarSummary!A2:L';
      const calUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${calRange}`;
      const calResponse = await fetch(calUrl, { headers: { 'Authorization': `Bearer ${token}` } });
      if (calResponse.ok) {
        const calData = await calResponse.json();
        if (calData.values) {
          const calMap = {};
          for (const row of calData.values) {
            const d = new Date(row[0]);
            if (isNaN(d.getTime()) || d < startDate || d > endDate) continue;
            const name = row[2];
            if (!calMap[name]) calMap[name] = { events: 0, aiUsing: 0, totalMinutes: 0 };
            calMap[name].events += parseInt(row[3]) || 0;
            calMap[name].aiUsing += parseInt(row[5]) || 0;
            calMap[name].totalMinutes += parseInt(row[4]) || 0;
          }
          calendarSummary = Object.entries(calMap)
            .map(([name, d]) => ({
              name, events: d.events, aiUsing: d.aiUsing,
              aiRate: d.events > 0 ? Math.round(d.aiUsing / d.events * 100) : 0,
              hours: Math.round(d.totalMinutes / 60)
            }))
            .sort((a, b) => b.aiRate - a.aiRate);
        }
      }
    } catch (e) { /* CalendarSummaryが無くても続行 */ }

    const formatDate = (d) => `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;

    // --- AI Impact Log の取得 (オプション) ---
    let impactText = '';
    let aiTimeSaved = 0;
    try {
      const impactRange = 'AiImpactLog!A2:H';
      const impactUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${impactRange}`;
      const impactResponse = await fetch(impactUrl, { headers: { 'Authorization': `Bearer ${token}` } });
      if (impactResponse.ok) {
        const impactData = await impactResponse.json();
        if (impactData.values) {
          const mvps = [];
          for (const row of impactData.values) {
            const d = new Date(row[0]);
            if (isNaN(d.getTime()) || d < startDate || d > endDate) continue;

            const saved = parseInt(row[6]) || 0;
            aiTimeSaved += saved;

            const score = parseInt(row[5]) || 0;
            if (score >= 4) {
              mvps.push({ task: row[2], score, action: row[7], time: saved });
            }
          }
          if (mvps.length > 0) {
            impactText = '\n## 高いインパクトを生んだ事例 (Impact MVP)\n' + mvps.map(m =>
              `  - タスク: ${sanitizeForPrompt(m.task)} (スコア: Lv${m.score}, 削減時間: ${m.time}分)\n    価値創造アクション: ${sanitizeForPrompt(m.action)}`
            ).join('\n');
          }
        }
      }
    } catch (e) { /* AiImpactLogが無くても続行 */ }

    // 短縮時間の合計をtotalsに追加
    totals.aiTimeSaved = aiTimeSaved;

    return {
      period: { start: formatDate(startDate), end: formatDate(endDate), days: uniqueDays },
      totals,
      averages: {
        dailyEvents: Math.round(totals.events / uniqueDays),
        dailyMinutes: Math.round(totals.totalMinutes / uniqueDays),
        aiRate
      },
      calendarSummary,
      impactText
    };
  } catch (e) {
    console.error('バックグラウンドデータ取得エラー:', e);
    return null;
  }
}

// バックグラウンド用Gemini API呼び出し
async function callGeminiBackground(apiKey, summaryData, type) {
  const periodLabel = type === 'weekly' ? '週次' : '月次';

  let calendarText = '';
  if (summaryData.calendarSummary && summaryData.calendarSummary.length > 0) {
    calendarText = '\n## カレンダー別ランキング\n' + summaryData.calendarSummary
      .map((c, i) => `  ${i + 1}. ${c.name}: ${c.events}件, AI活用${c.aiUsing}件, 活用率${c.aiRate}%, ${c.hours}h`)
      .join('\n');
  }

  const prompt = `あなたは業務効率化の専門家だ。データから読み取れる事実に基づいて、${periodLabel}フィードバックを行え。

## 入力データ
- 期間: ${summaryData.period.start} ～ ${summaryData.period.end}（${summaryData.period.days}日間）
- 予定数: ${summaryData.totals.events}件 / 稼働: ${Math.round(summaryData.totals.totalMinutes / 60)}h
- AI活用中: ${summaryData.totals.aiUsingCount}件（${Math.round(summaryData.totals.aiUsingMinutes / 60)}h）
- AI未活用: ${summaryData.totals.aiNotUsingCount}件（${Math.round(summaryData.totals.aiNotUsingMinutes / 60)}h）
- AI余地あり: ${summaryData.totals.aiPotentialCount}件（${Math.round(summaryData.totals.aiPotentialMinutes / 60)}h）
- AI活用率: ${summaryData.averages.aiRate}%
- 今期AIによる全体の削減時間: ${Math.round((summaryData.totals.aiTimeSaved || 0) / 60)}h${summaryData.calendarSummary && summaryData.calendarSummary.length > 0 ? '\n' + calendarText : ''}${summaryData.impactText ? '\n' + summaryData.impactText : ''}

## 回答フォーマット（必ずこの構造で出力せよ）

### 💪 労い
（${type === 'weekly' ? '今週' : '今月'}の数値に触れつつ、取り組みへの労いを1文で）

### 🎯 AI活用の提案
- 【具体ツール名】を【どの業務に】【どう使うか】
- （同上、もう1点あれば）

### 🔄 続けるべきこと
- （データから読み取れる良い傾向を1点、数値根拠付きで）

### 💥 来${type === 'weekly' ? '週' : '月'}のベストアクション
- （1つだけ。最も時間削減効果が大きいもの。具体的な行動を書け）

### 📊 スコア: ○○点
（一言理由）

## 出力ルール
- 前置き・挨拶・まとめ文は書くな
- 「〜と思います」「〜でしょう」は使うな。言い切れ
- 同じ内容を別の表現で繰り返すな
- ツール名は具体的に書け（例: ChatGPT、GitHub Copilot、Gemini、NotebookLM）
- 数値は元データから引用し、根拠のない数値を出すな`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`;

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 8192 }
    })
  });

  if (!response.ok) {
    throw new Error('Gemini API呼び出し失敗');
  }

  const data = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text || '分析結果を取得できませんでした';
}

// バックグラウンド用レポートシート保存
async function saveReportBackground(token, spreadsheetId, type, summary, analysis) {
  const sheetName = type === 'weekly' ? 'WeeklyReport' : 'MonthlyReport';

  // シート存在確認（なければ作成）
  try {
    const checkUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${sheetName}!A1:L1`;
    const checkResponse = await fetch(checkUrl, { headers: { 'Authorization': `Bearer ${token}` } });

    if (!checkResponse.ok) {
      // シート作成
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          requests: [{ addSheet: { properties: { title: sheetName } } }]
        })
      });
    }

    const checkData = await checkResponse.json();
    if (!checkData.values || checkData.values.length === 0) {
      // ヘッダー追加
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${sheetName}!A1:L1?valueInputOption=RAW`, {
        method: 'PUT',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          values: [['生成日時', '期間開始', '期間終了', '日数', 'ユーザー数', '予定数', '稼働時間(h)', 'AI活用件数', 'AI未活用件数', 'AI余地件数', 'AI活用率(%)', 'AI分析レポート']]
        })
      });
    }
  } catch (e) { /* 初回は失敗しうるので続行 */ }

  // レポート行を追加
  const row = [
    new Date().toISOString(),
    summary.period.start, summary.period.end, summary.period.days,
    1, summary.totals.events, Math.round(summary.totals.totalMinutes / 60),
    summary.totals.aiUsingCount, summary.totals.aiNotUsingCount, summary.totals.aiPotentialCount,
    summary.averages.aiRate, analysis
  ];

  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${sheetName}!A:L:append?valueInputOption=USER_ENTERED`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [row] })
  });
}

// ========================================
// 日次自動日報・Slack通知
// ========================================

async function runDailyReportCheck(isManual = false) {
  let slackWebhook = '';
  let enableSlack = false;
  try {
    const config = await chrome.storage.local.get(['geminiApiKey', 'spreadsheetId', 'slackWebhookUrl', 'selectedCalendars', 'enableSlackNotification', 'sheetName']);
    slackWebhook = config.slackWebhookUrl || '';
    enableSlack = config.enableSlackNotification !== false && slackWebhook !== '';

    if (!config.geminiApiKey || !config.spreadsheetId) {
      console.log('必要な設定（APIキー、スプレッドシートID）が不足しています。');
      return { success: false, error: '設定(APIキー、シートID)が不足しています' };
    }

    // トークン取得
    const token = await new Promise((resolve, reject) => {
      chrome.identity.getAuthToken({ interactive: false }, (token) => {
        if (chrome.runtime.lastError) reject(chrome.runtime.lastError);
        else resolve(token);
      });
    });

    if (!token) {
      console.log('認証トークンが取得できないため日次チェックをスキップ');
      if (enableSlack) await sendSlackMessageBackground(slackWebhook, "⚠️ *自動日報エラー*\nGoogle認証トークンが取得できませんでした。\n拡張機能のサイドパネルを開き、再度Googleログインを行ってください。");
      return { success: false, error: 'Google認証トークンが取得できませんでした' };
    }

    // 1. 本日の記録があるかチェック
    const hasRecord = await checkTodayRecordBackground(token, config.spreadsheetId);
    if (hasRecord) {
      console.log('本日の日報はすでに提出されています。');
      if (isManual) {
        // 手動実行時は強制出力させる
        if (enableSlack) await sendSlackMessageBackground(slackWebhook, "ℹ️ *【インフォメーション】*\n本日の「AI時間計測・日報」はすでにDailySummaryに記録されていますが、手動リクエストにより日報を生成・送信します。");
      } else {
        return { success: true, message: 'すでに提出済みです' };
      }
    } else {
      console.log('本日の日報が未提出です。自動生成を開始します。');
      // 2. Slackに未提出アラート送信 (自動実行時のみ)
      if (!isManual && enableSlack) {
        await sendSlackMessageBackground(slackWebhook, "🚨 *【未提出アラート】*\n本日の「AI時間計測・日報」がまだ記録されていません。\nカレンダーの予定データから、AI(Gemini)を用いて日報を自動生成します...");
      }
    }

    // 3. 今日の予定を取得
    const eventsInfo = await getTodayCalendarEventsBackground(token, config.spreadsheetId, config.sheetName, config.selectedCalendars);
    if (eventsInfo.eventCount === 0) {
      if (enableSlack) await sendSlackMessageBackground(slackWebhook, "ℹ️ *【自動日報スキップ】*\n本日のカレンダー予定が見つからなかったため、日報の自動生成をスキップしました。");
      return { success: false, error: '本日のカレンダー予定がありません' };
    }

    // 4. Geminiでレポート生成
    let reportText = '';
    try {
      reportText = await generateDailyReportWithGemini(config.geminiApiKey, eventsInfo);
    } catch (apiErr) {
      if (enableSlack) await sendSlackMessageBackground(slackWebhook, "⚠️ *自動日報エラー*\nGemini APIによる日報テキストの生成に失敗しました。");
      throw apiErr;
    }

    // 5. Slackに送信
    const headerPrefix = isManual ? "⚡ *【手動リクエスト日報】*" : "📝 *【自動生成された本日の日報】*";
    if (enableSlack) {
      await sendSlackMessageBackground(slackWebhook, `${headerPrefix}\n\n${reportText}`);
    }

    // 6. 次回以降に重複して実行されないようDailySummaryに簡易記録を残す
    await appendAutoDailySummaryBackground(token, config.spreadsheetId, eventsInfo);

    console.log('日次チェックおよび自動日報の処理が完了しました。');
    return { success: true };
  } catch (error) {
    console.error('日次日報自動生成エラー:', error);
    if (enableSlack) {
      await sendSlackMessageBackground(slackWebhook, `⚠️ *自動日報エラー*\n予期せぬエラーが発生しました: ${error.message}`);
    }
    return { success: false, error: error.message };
  }
}

async function checkTodayRecordBackground(token, spreadsheetId) {
  try {
    const range = 'DailySummary!A:A'; // 日付列のみチェック
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;
    const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
    if (!response.ok) return false;

    const data = await response.json();
    if (!data.values || data.values.length <= 1) return false;

    const today = new Date();
    // app.js の formatDate に合わせた YYYY-MM-DD 形式で直接照合
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const todayStr = `${year}-${month}-${day}`;

    for (let i = 1; i < data.values.length; i++) {
      const rowDateStr = data.values[i][0];
      if (todayStr === rowDateStr) {
        return true; // 本日の記録あり
      }
    }
    return false;
  } catch (e) {
    console.error('本日記録の確認に失敗:', e);
    return false; // エラー時はとりあえず無いと判定して進む
  }
}

async function getTodayCalendarEventsBackground(token, spreadsheetId, sheetName, selectedCalendars = []) {
  try {
    const today = new Date();
    const startObj = new Date(today);
    startObj.setHours(0, 0, 0, 0);
    const endObj = new Date(today);
    endObj.setHours(23, 59, 59, 999);

    const timeMin = startObj.toISOString();
    const timeMax = endObj.toISOString();

    const fetchEventsForCalendar = async (calendarId, calendarName) => {
      const url = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events?timeMin=${encodeURIComponent(timeMin)}&timeMax=${encodeURIComponent(timeMax)}&singleEvents=true&orderBy=startTime`;
      const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
      if (!response.ok) return [];
      const data = await response.json();
      return (data.items || []).map(e => ({ ...e, calendarName: calendarName || 'primary' }));
    };

    let allItems = [];
    if (!selectedCalendars || selectedCalendars.length === 0) {
      allItems = await fetchEventsForCalendar('primary', 'primary');
    } else {
      // 選択されたカレンダーIDからカレンダー情報を取得するためにAPIをコール
      const listUrl = 'https://www.googleapis.com/calendar/v3/users/me/calendarList';
      const listRes = await fetch(listUrl, { headers: { 'Authorization': `Bearer ${token}` } });
      let calendars = [];
      if (listRes.ok) {
        const listData = await listRes.json();
        calendars = listData.items || [];
      }
      
      const targetCalendars = calendars.filter(c => selectedCalendars.includes(c.id));
      if (targetCalendars.length === 0) {
        allItems = await fetchEventsForCalendar('primary', 'primary');
      } else {
        const promises = targetCalendars.map(cal => fetchEventsForCalendar(cal.id, cal.summary));
        const results = await Promise.all(promises);
        results.forEach(items => allItems.push(...items));
      }
    }

    // 取得した全アイテムを時間順にソート
    allItems.sort((a, b) => {
      const startA = a.start?.dateTime || a.start?.date || '';
      const startB = b.start?.dateTime || b.start?.date || '';
      return startA.localeCompare(startB);
    });

    // スプレッドシートから本日の既存AIスコアを取得
    let aiScoreMap = {};
    if (spreadsheetId) {
      try {
        const safeSheetName = sheetName ? sheetName.replace(/[\\/*?\\[\\]':!]/g, '_') : 'Sheet1';
        const range = `'${safeSheetName}'!B:H`; // B:日付, C:タイトル ... H:活用率
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`;
        const res = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
        if (res.ok) {
          const data = await res.json();
          if (data.values && data.values.length > 0) {
            const year = today.getFullYear();
            const month = String(today.getMonth() + 1).padStart(2, '0');
            const day = String(today.getDate()).padStart(2, '0');
            const todayFormatted = `${year}/${month}/${day}`; // formatReportForSheet matches app.js? Wait, app.js saves 'YYYY-MM-DD' usually. Let's try both
            const todayDash = `${year}-${month}-${day}`;
            
            for (const row of data.values) {
              const rowDate = row[0]; // 列B
              if (rowDate === todayFormatted || rowDate === todayDash) {
                 const title = row[1]; // 列C
                 const rate = row[6]; // 列H (index 6 from B)
                 if (title && rate) {
                   aiScoreMap[title] = rate;
                 }
              }
            }
          }
        }
      } catch (e) {
        console.error('既存スコアの取得に失敗:', e);
      }
    }

    let totalMinutes = 0;
    let eventListText = "";
    let allItemsWithDetails = [];
    const padZero = (num) => String(num).padStart(2, '0');

    for (const event of allItems) {
      if (!event.start.dateTime) continue; // 終日イベントスキップ

      const start = new Date(event.start.dateTime);
      const end = new Date(event.end.dateTime);
      const durationMin = Math.round((end.getTime() - start.getTime()) / (1000 * 60));
      totalMinutes += durationMin;

      const startStr = `${padZero(start.getHours())}:${padZero(start.getMinutes())}`;
      const endStr = `${padZero(end.getHours())}:${padZero(end.getMinutes())}`;
      let aiScore = aiScoreMap[event.summary] || '50'; // 既存記録があればそれ、なければ50%
      
      allItemsWithDetails.push({
        summary: event.summary || '(無題)',
        startStr,
        endStr,
        durationMin,
        aiScore
      });
      eventListText += `${startStr} - ${endStr} | スコア: ${aiScore}%\n${event.summary || '(無題)'}\n`;
    }

    // 平均AIスコアの算出
    let totalScore = 0;
    for (const item of allItemsWithDetails) {
        totalScore += parseInt(item.aiScore, 10) || 50;
    }
    const avgScore = allItemsWithDetails.length > 0 ? Math.round(totalScore / allItemsWithDetails.length) : 50;

    return {
      eventCount: allItems.filter(e => e.start.dateTime).length,
      totalMinutes,
      eventListText,
      avgScore,
      items: allItemsWithDetails
    };
  } catch (e) {
    console.error('日次予定取得エラー:', e);
    return { eventCount: 0, totalMinutes: 0, eventListText: '' };
  }
}

async function appendAutoDailySummaryBackground(token, spreadsheetId, eventsInfo) {
  try {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const todayStr = `${year}-${month}-${day}`;

    const { userInfo } = await chrome.storage.local.get('userInfo');
    const userEmail = userInfo ? userInfo.email : 'auto-generated';

    const row = [
      todayStr, userEmail, eventsInfo.eventCount, eventsInfo.totalMinutes,
      0, 0, 0, 0, 0, 0, 0, new Date().toISOString()
    ];

    const range = `'DailySummary'!A:L`;
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
  } catch (e) {
    console.error('バックグラウンドでのDailySummary記録に失敗:', e);
  }
}

async function generateDailyReportWithGemini(apiKey, eventsInfo) {
  const dateStr = new Date().toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });
  const totalHours = Math.round(eventsInfo.totalMinutes / 60 * 10) / 10;

  const prompt = `あなたは優秀なアシスタントです。以下の本日のカレンダー予定データをもとに、日報を自動作成してください。
ただし、内容は以下のフォーマットに**完全に従って**出力してください。余計な挨拶や説明は不要です。

【フォーマット】
${dateStr}
予定数
${eventsInfo.eventCount}件
合計時間
${totalHours}時間
平均AIスコア
${eventsInfo.avgScore || 50}%

【予定データ一覧】
${eventsInfo.eventListText}

【指示】
・上記【フォーマット】の構成を守って出力してください。
・各予定のタイトル（業務内容）を出力し、その次の行に「[開始時間] - [終了時間] | スコア: [スコア]%」を出力してください（スコアは予定データ一覧のものをそのまま使用すること）。
・「平均AIスコア」も予定データ一覧から算出したものをそのまま使用してください。

出力例：
${dateStr}
予定数
3件
合計時間
8時間
平均AIスコア
50%
アプリ修正等
11:00 - 13:00 | スコア: 50%
要件定義等の修正、確認、アプリ開発
13:00 - 18:00 | スコア: 100%
アプリの修正
19:15 - 20:15 | スコア: 0%`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`;
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 2048 }
    })
  });

  if (!response.ok) throw new Error('Gemini API呼び出し失敗');
  const data = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text || '日報のテキスト生成に失敗しました（Geminiから空のレスポンス）。';
}

async function sendSlackMessageBackground(webhookUrl, text) {
  try {
    const payload = { text };
    await fetch(webhookUrl, {
      method: 'POST',
      body: JSON.stringify(payload)
    });
  } catch (e) {
    console.error('Slack通知エラー:', e);
  }
}


// 拡張機能インストール時にアラーム設定
chrome.runtime.onInstalled.addListener(() => {
  setupReportAlarms();
  setupDailyAlarm();
  console.log('レポート自動生成アラームを設定しました');
});

// サービスワーカー起動時にもアラーム設定（復帰対策）
setupReportAlarms();
setupDailyAlarm();

// === スリープ復帰・ブラウザ起動時のリカバー処理 ===
// スリープ中にアラームが発火せずにスキップされてしまった場合を救済する
async function checkAndRecoverDailyReport() {
  try {
    const config = await chrome.storage.local.get(['enableDailyAlarm', 'alarmTime']);
    if (!config.enableDailyAlarm) return;

    let hour = 18;
    let minute = 0;
    if (config.alarmTime !== undefined) {
      if (typeof config.alarmTime === 'string' && config.alarmTime.includes(':')) {
        const parts = config.alarmTime.split(':');
        hour = parseInt(parts[0], 10);
        minute = parseInt(parts[1], 10) || 0;
      } else {
        hour = parseInt(config.alarmTime, 10) || 18;
      }
    }

    const now = new Date();
    // 設定時刻を【超えている】かチェック
    const isPastAlarmTime = (now.getHours() > hour) || (now.getHours() === hour && now.getMinutes() >= minute);

    if (isPastAlarmTime) {
      // トークン取得を試行
      chrome.identity.getAuthToken({ interactive: false }, async (token) => {
        if (chrome.runtime.lastError || !token) return; // サイレントに失敗させる（設定不備や未ログイン）

        const { spreadsheetId } = await chrome.storage.local.get('spreadsheetId');
        if (!spreadsheetId) return;

        // すでに本日の記録があるか確認
        const hasRecord = await checkTodayRecordBackground(token, spreadsheetId);
        if (!hasRecord) {
          console.log('【リカバー処理】設定時刻を過ぎていますが、本日の日報が未提出のため自動チェックを開始します。');
          await runDailyReportCheck(false);
        }
      });
    }
  } catch (e) {
    console.warn('Recover check failed:', e);
  }
}

// OSのアイドル状態（スリープ）からの復帰時
chrome.idle.onStateChanged.addListener((newState) => {
  if (newState === 'active') {
    checkAndRecoverDailyReport();
    setupDailyAlarm(); // アラームの再登録
  }
});

// ブラウザ起動時等でSWがアクティブになった時も一応チェック
checkAndRecoverDailyReport();
