// AI活用度測定 Chrome拡張機能 - バックグラウンドサービスワーカー

// サイドパネルを開く設定
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch((error) => console.error('サイドパネル設定エラー:', error));

// メッセージリスナー
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'getAuthToken') {
    handleGetAuthToken(sendResponse);
    return true; // 非同期レスポンスのため
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
async function handleGetAuthToken(sendResponse) {
  try {
    const token = await new Promise((resolve, reject) => {
      chrome.identity.getAuthToken({ interactive: true }, (token) => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
        } else {
          resolve(token);
        }
      });
    });
    sendResponse({ success: true, token });
  } catch (error) {
    console.error('認証エラー:', error);
    sendResponse({ success: false, error: error.message });
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

      // Googleのトークン失効エンドポイントを呼び出し
      await fetch(`https://accounts.google.com/o/oauth2/revoke?token=${token}`);
    }

    sendResponse({ success: true });
  } catch (error) {
    console.error('ログアウトエラー:', error);
    sendResponse({ success: false, error: error.message });
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
        results.push({ id: item.id, success: false, error: error.message });
      }
    }

    // 失敗したアイテムをキューに戻す
    await chrome.storage.local.set({ offlineQueue: failedItems });

    sendResponse({ success: true, results, remaining: failedItems.length });
  } catch (error) {
    console.error('キュー再送エラー:', error);
    sendResponse({ success: false, error: error.message });
  }
}

// スプレッドシートへの送信（background.jsから呼び出し用）
async function sendToSpreadsheet(data) {
  const { spreadsheetId, sheetName } = await chrome.storage.sync.get(['spreadsheetId', 'sheetName']);

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

  const range = `${sheetName || 'Sheet1'}!A:J`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}:append?valueInputOption=USER_ENTERED`;

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

    console.log('Notion API Request:', { url, method, body });

    const response = await fetch(url, {
      method: method || 'GET',
      headers: headers || {},
      body: body ? JSON.stringify(body) : undefined
    });

    const data = await response.json();

    console.log('Notion API Response:', { status: response.status, data });

    if (!response.ok) {
      // Notion APIのエラー詳細を取得
      const errorMessage = data.message || data.error?.message || `HTTP ${response.status}`;
      const errorCode = data.code || data.status || response.status;
      console.error('Notion API Error:', { errorCode, errorMessage, fullResponse: data });

      sendResponse({
        success: false,
        error: `${errorCode}: ${errorMessage}`,
        status: response.status,
        details: data
      });
    } else {
      sendResponse({ success: true, data });
    }
  } catch (error) {
    console.error('Notion APIリクエストエラー:', error);
    sendResponse({ success: false, error: error.message });
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

// アラーム発火時の処理
chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name === 'weeklyReport') {
    console.log('週次レポート自動生成開始');
    await generateScheduledReport('weekly');
  } else if (alarm.name === 'monthlyReport') {
    console.log('月次レポート自動生成開始');
    await generateScheduledReport('monthly');
  }
});

// バックグラウンドでのレポート生成
async function generateScheduledReport(type) {
  try {
    // 設定を取得
    const { geminiApiKey, spreadsheetId } = await chrome.storage.sync.get(['geminiApiKey', 'spreadsheetId']);

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

    return {
      period: { start: formatDate(startDate), end: formatDate(endDate), days: uniqueDays },
      totals,
      averages: {
        dailyEvents: Math.round(totals.events / uniqueDays),
        dailyMinutes: Math.round(totals.totalMinutes / uniqueDays),
        aiRate
      },
      calendarSummary
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
${calendarText}

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

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
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

// 拡張機能インストール時にアラーム設定
chrome.runtime.onInstalled.addListener(() => {
  setupReportAlarms();
  console.log('レポート自動生成アラームを設定しました');
});

// サービスワーカー起動時にもアラーム設定（復帰対策）
setupReportAlarms();
