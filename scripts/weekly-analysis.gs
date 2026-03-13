/**
 * AI活用度測定 - 週次・月次分析スクリプト
 * 
 * このスクリプトをスプレッドシートのApps Scriptエディタに貼り付けて使用します。
 * 設定手順:
 * 1. スプレッドシートを開く
 * 2. 拡張機能 > Apps Script
 * 3. このコードを貼り付けて保存
 * 4. トリガーを設定（週次/月次実行）
 */

// ========================================
// 設定
// ========================================
const CONFIG = {
  // Gemini API Key（スクリプトプロパティから取得）
  get GEMINI_API_KEY() { return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || ''; },
  
  // シート名
  DAILY_SUMMARY_SHEET: 'DailySummary',
  WEEKLY_REPORT_SHEET: 'WeeklyReport',
  MONTHLY_REPORT_SHEET: 'MonthlyReport',
  
  // メール送信先（スクリプトプロパティから取得）
  get REPORT_EMAIL() { return PropertiesService.getScriptProperties().getProperty('REPORT_EMAIL') || ''; },
};

// ========================================
// メイン関数
// ========================================

/**
 * 週次レポートを生成（手動実行またはトリガー実行用）
 */
function generateWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName(CONFIG.DAILY_SUMMARY_SHEET);
  
  if (!dailySheet) {
    Logger.log('DailySummaryシートが見つかりません');
    return;
  }
  
  // 過去7日間のデータを取得
  const data = getLastWeekData(dailySheet);
  
  if (data.length === 0) {
    Logger.log('過去7日間のデータがありません');
    return;
  }
  
  // 集計
  const summary = aggregateWeeklyData(data);
  
  // カレンダー別データを取得・集計
  const calendarSheet = ss.getSheetByName('CalendarSummary');
  if (calendarSheet) {
    const calData = getCalendarDataForPeriod(calendarSheet, 7);
    summary.calendarSummary = aggregateCalendarData(calData);
  }
  
  // AI分析実行
  const analysis = analyzeWithGemini(summary);
  
  // レポートシートに出力
  outputWeeklyReport(ss, summary, analysis);
  
  // メール送信
  if (CONFIG.REPORT_EMAIL) {
    sendReportEmail(summary, analysis);
  }
  
  Logger.log('週次レポート生成完了');
}

/**
 * 過去7日間のデータを取得
 */
function getLastWeekData(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const today = new Date();
  const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  return rows.filter(row => {
    const date = new Date(row[0]); // 日付列
    return date >= weekAgo && date <= today;
  }).map(row => ({
    date: row[0],
    user: row[1],
    eventCount: row[2],
    totalMinutes: row[3],
    aiUsingCount: row[4],
    aiNotUsingCount: row[5],
    aiPotentialCount: row[6],
    avgRate: row[7],
    aiUsingMinutes: row[8],
    aiNotUsingMinutes: row[9],
    aiPotentialMinutes: row[10]
  }));
}

/**
 * 指定期間のCalendarSummaryデータを取得
 */
function getCalendarDataForPeriod(sheet, days) {
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // ヘッダースキップ
  
  const today = new Date();
  const startDate = new Date(today.getTime() - days * 24 * 60 * 60 * 1000);
  
  return rows.filter(row => {
    const date = new Date(row[0]);
    return date >= startDate && date <= today;
  }).map(row => ({
    date: row[0],
    user: row[1],
    calendarName: row[2],
    eventCount: row[3],
    totalMinutes: row[4],
    aiUsingCount: row[5],
    aiNotUsingCount: row[6],
    aiPotentialCount: row[7],
    avgRate: row[8],
    aiUsingMinutes: row[9],
    aiNotUsingMinutes: row[10],
    aiPotentialMinutes: row[11]
  }));
}

/**
 * カレンダー別データを集計
 */
function aggregateCalendarData(data) {
  const calMap = {};
  for (const row of data) {
    const key = row.calendarName;
    if (!calMap[key]) {
      calMap[key] = { events: 0, aiUsing: 0, aiNotUsing: 0, totalMinutes: 0, aiUsingMinutes: 0 };
    }
    calMap[key].events += row.eventCount;
    calMap[key].aiUsing += row.aiUsingCount;
    calMap[key].aiNotUsing += row.aiNotUsingCount;
    calMap[key].totalMinutes += row.totalMinutes;
    calMap[key].aiUsingMinutes += row.aiUsingMinutes;
  }
  
  // AI活用率でソート
  return Object.entries(calMap)
    .map(([name, d]) => {
      const aiRate = d.events > 0 ? Math.round(d.aiUsing / d.events * 100) : 0;
      return { name, events: d.events, aiUsing: d.aiUsing, aiNotUsing: d.aiNotUsing, aiRate, hours: Math.round(d.totalMinutes / 60) };
    })
    .sort((a, b) => b.aiRate - a.aiRate);
}

/**
 * 週次データを集計
 */
function aggregateWeeklyData(data) {
  const summary = {
    period: {
      start: data[0].date,
      end: data[data.length - 1].date
    },
    totals: {
      days: [...new Set(data.map(d => d.date))].length,
      users: [...new Set(data.map(d => d.user))].length,
      events: 0,
      totalMinutes: 0,
      aiUsingCount: 0,
      aiNotUsingCount: 0,
      aiPotentialCount: 0,
      aiUsingMinutes: 0,
      aiNotUsingMinutes: 0,
      aiPotentialMinutes: 0
    },
    averages: {
      dailyEvents: 0,
      dailyMinutes: 0,
      aiRate: 0
    },
    dailyData: [],
    userSummary: {}
  };
  
  // 集計
  for (const row of data) {
    summary.totals.events += row.eventCount;
    summary.totals.totalMinutes += row.totalMinutes;
    summary.totals.aiUsingCount += row.aiUsingCount;
    summary.totals.aiNotUsingCount += row.aiNotUsingCount;
    summary.totals.aiPotentialCount += row.aiPotentialCount;
    summary.totals.aiUsingMinutes += row.aiUsingMinutes;
    summary.totals.aiNotUsingMinutes += row.aiNotUsingMinutes;
    summary.totals.aiPotentialMinutes += row.aiPotentialMinutes;
    
    // ユーザー別集計
    if (!summary.userSummary[row.user]) {
      summary.userSummary[row.user] = {
        events: 0,
        minutes: 0,
        aiUsingCount: 0,
        aiNotUsingCount: 0,
        avgRate: 0,
        rateSum: 0,
        count: 0
      };
    }
    const user = summary.userSummary[row.user];
    user.events += row.eventCount;
    user.minutes += row.totalMinutes;
    user.aiUsingCount += row.aiUsingCount;
    user.aiNotUsingCount += row.aiNotUsingCount;
    user.rateSum += row.avgRate;
    user.count++;
  }
  
  // 平均計算
  summary.averages.dailyEvents = Math.round(summary.totals.events / summary.totals.days);
  summary.averages.dailyMinutes = Math.round(summary.totals.totalMinutes / summary.totals.days);
  
  const totalAiRelated = summary.totals.aiUsingCount + summary.totals.aiPotentialCount;
  const totalEvents = summary.totals.aiUsingCount + summary.totals.aiNotUsingCount + summary.totals.aiPotentialCount;
  summary.averages.aiRate = totalEvents > 0 ? Math.round(summary.totals.aiUsingCount / totalEvents * 100) : 0;
  
  // ユーザー別平均レート
  for (const user in summary.userSummary) {
    const u = summary.userSummary[user];
    u.avgRate = u.count > 0 ? Math.round(u.rateSum / u.count) : 0;
  }
  
  return summary;
}

/**
 * Gemini APIで分析
 */
function analyzeWithGemini(summary) {
  if (!CONFIG.GEMINI_API_KEY) {
    return '（Gemini API Keyが設定されていません。スクリプトプロパティに GEMINI_API_KEY を設定してください。）';
  }
  
  // カレンダー別ランキングテキスト生成
  let calendarText = '';
  if (summary.calendarSummary && summary.calendarSummary.length > 0) {
    calendarText = '\n## カレンダー別ランキング\n' + summary.calendarSummary
      .map((c, i) => `  ${i + 1}. ${c.name}: ${c.events}件, AI活用${c.aiUsing}件, 活用率${c.aiRate}%, ${c.hours}h`)
      .join('\n');
  }

  const prompt = `あなたは業務効率化の専門家だ。データから読み取れる事実に基づいて、週次フィードバックを行え。

## 入力データ
- 期間: ${summary.period.start} ～ ${summary.period.end}（${summary.totals.days}日間）
- ユーザー数: ${summary.totals.users}人 / 予定数: ${summary.totals.events}件 / 稼働: ${Math.round(summary.totals.totalMinutes / 60)}h
- AI活用中: ${summary.totals.aiUsingCount}件（${Math.round(summary.totals.aiUsingMinutes / 60)}h）
- AI未活用: ${summary.totals.aiNotUsingCount}件（${Math.round(summary.totals.aiNotUsingMinutes / 60)}h）
- AI余地あり: ${summary.totals.aiPotentialCount}件（${Math.round(summary.totals.aiPotentialMinutes / 60)}h）
- AI活用率: ${summary.averages.aiRate}%
${calendarText}

## 回答フォーマット（必ずこの構造で出力せよ）

### 💪 労い
（今週の数値に触れつつ、取り組みへの労いを1文で）

### 🎯 AI活用の提案
- 【具体ツール名】を【どの業務に】【どう使うか】
- （同上、もう1点あれば）

### 🔄 続けるべきこと
- （データから読み取れる良い傾向を1点、数値根拠付きで）

### 💥 来週のベストアクション
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
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': CONFIG.GEMINI_API_KEY },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 2048 }
      })
    });
    
    const json = JSON.parse(response.getContentText());
    return json.candidates?.[0]?.content?.parts?.[0]?.text || '分析結果を取得できませんでした';
  } catch (e) {
    Logger.log('Gemini APIエラー: ' + e.message);
    return 'AI分析エラー: ' + e.message;
  }
}

/**
 * 週次レポートを出力
 */
function outputWeeklyReport(ss, summary, analysis) {
  let sheet = ss.getSheetByName(CONFIG.WEEKLY_REPORT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.WEEKLY_REPORT_SHEET);
  }
  
  const now = new Date();
  const row = [
    now,
    summary.period.start,
    summary.period.end,
    summary.totals.days,
    summary.totals.users,
    summary.totals.events,
    Math.round(summary.totals.totalMinutes / 60),
    summary.totals.aiUsingCount,
    summary.totals.aiNotUsingCount,
    summary.totals.aiPotentialCount,
    summary.averages.aiRate,
    analysis
  ];
  
  // ヘッダーがなければ追加
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '生成日時', '期間開始', '期間終了', '日数', 'ユーザー数', 
      '予定数', '稼働時間(h)', 'AI活用件数', 'AI未活用件数', 
      'AI余地件数', 'AI活用率(%)', 'AI分析レポート'
    ]);
  }
  
  sheet.appendRow(row);
}

/**
 * レポートをメール送信
 */
function sendReportEmail(summary, analysis) {
  const subject = `【週次AI活用レポート】${summary.period.start} ～ ${summary.period.end}`;
  const body = `
AI活用度測定 週次レポート

========================================
期間: ${summary.period.start} ～ ${summary.period.end}
対象: ${summary.totals.users}名 / ${summary.totals.days}日
========================================

【サマリー】
・予定数: ${summary.totals.events}件
・稼働時間: ${Math.round(summary.totals.totalMinutes / 60)}時間
・AI活用率: ${summary.averages.aiRate}%

【AI活用状況】
・活用中: ${summary.totals.aiUsingCount}件
・未活用: ${summary.totals.aiNotUsingCount}件
・余地あり: ${summary.totals.aiPotentialCount}件

========================================
【AI分析レポート】
========================================
${analysis}

----------------------------------------
このメールはGoogle Apps Scriptにより自動送信されました。
`;

  GmailApp.sendEmail(CONFIG.REPORT_EMAIL, subject, body);
}

// ========================================
// トリガー設定用関数
// ========================================

/**
 * 週次トリガーを設定
 * 毎週月曜日の朝9時に実行
 */
function setupWeeklyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'generateWeeklyReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成
  ScriptApp.newTrigger('generateWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
    
  Logger.log('週次トリガーを設定しました（毎週月曜9時）');
}

// ========================================
// 月次レポート機能
// ========================================

/**
 * 月次レポートを生成（手動実行またはトリガー実行用）
 */
function generateMonthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName(CONFIG.DAILY_SUMMARY_SHEET);
  
  if (!dailySheet) {
    Logger.log('DailySummaryシートが見つかりません');
    return;
  }
  
  // 過去30日間のデータを取得
  const data = getLastMonthData(dailySheet);
  
  if (data.length === 0) {
    Logger.log('過去30日間のデータがありません');
    return;
  }
  
  // 集計（週次と同じ関数を使用）
  const summary = aggregateWeeklyData(data);
  
  // カレンダー別データを取得・集計
  const calendarSheet = ss.getSheetByName('CalendarSummary');
  if (calendarSheet) {
    const calData = getCalendarDataForPeriod(calendarSheet, 30);
    summary.calendarSummary = aggregateCalendarData(calData);
  }
  
  // AI分析実行（月次用プロンプト）
  const analysis = analyzeMonthlyWithGemini(summary);
  
  // レポートシートに出力
  outputMonthlyReport(ss, summary, analysis);
  
  // メール送信
  if (CONFIG.REPORT_EMAIL) {
    sendMonthlyReportEmail(summary, analysis);
  }
  
  Logger.log('月次レポート生成完了');
}

/**
 * 過去30日間のデータを取得
 */
function getLastMonthData(sheet) {
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  
  const today = new Date();
  const monthAgo = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  
  return rows.filter(row => {
    const date = new Date(row[0]);
    return date >= monthAgo && date <= today;
  }).map(row => ({
    date: row[0],
    user: row[1],
    eventCount: row[2],
    totalMinutes: row[3],
    aiUsingCount: row[4],
    aiNotUsingCount: row[5],
    aiPotentialCount: row[6],
    avgRate: row[7],
    aiUsingMinutes: row[8],
    aiNotUsingMinutes: row[9],
    aiPotentialMinutes: row[10]
  }));
}

/**
 * 月次分析用のGemini API呼び出し
 */
function analyzeMonthlyWithGemini(summary) {
  if (!CONFIG.GEMINI_API_KEY) {
    return '（Gemini API Keyが設定されていません。スクリプトプロパティに GEMINI_API_KEY を設定してください。）';
  }
  
  // ユーザー別サマリーを整形
  const userSummaryText = Object.entries(summary.userSummary)
    .map(([user, data]) => `  - ${user}: ${data.events}件, AI活用${data.aiUsingCount}件, 平均${data.avgRate}%`)
    .join('\n');
  
  // カレンダー別ランキングテキスト生成
  let calendarText = '';
  if (summary.calendarSummary && summary.calendarSummary.length > 0) {
    calendarText = '\n## カレンダー別ランキング\n' + summary.calendarSummary
      .map((c, i) => `  ${i + 1}. ${c.name}: ${c.events}件, AI活用${c.aiUsing}件, 活用率${c.aiRate}%, ${c.hours}h`)
      .join('\n');
  }

  const prompt = `あなたは業務効率化の専門家だ。データから読み取れる事実に基づいて、月次フィードバックを行え。

## 入力データ
- 期間: ${summary.period.start} ～ ${summary.period.end}（${summary.totals.days}日間）
- ユーザー数: ${summary.totals.users}人 / 予定数: ${summary.totals.events}件 / 稼働: ${Math.round(summary.totals.totalMinutes / 60)}h
- AI活用中: ${summary.totals.aiUsingCount}件（${Math.round(summary.totals.aiUsingMinutes / 60)}h）
- AI未活用: ${summary.totals.aiNotUsingCount}件（${Math.round(summary.totals.aiNotUsingMinutes / 60)}h）
- AI余地あり: ${summary.totals.aiPotentialCount}件（${Math.round(summary.totals.aiPotentialMinutes / 60)}h）
- AI活用率: ${summary.averages.aiRate}%
${calendarText}

## ユーザー別データ
${userSummaryText}

## 回答フォーマット（必ずこの構造で出力せよ）

### 💪 労い
（今月の数値に触れつつ、取り組みへの労いを1文で）

### 📊 月間スコア: ○○点（Sランク）
（数値根拠を含む一言理由）

### 🏆 ユーザーランキング
| 順位 | ユーザー | AI活用率 | 特記事項 |
|------|----------|----------|----------|
（上位者から順に。全員分記載）

### 🎯 AI活用の提案
- 【具体ツール名】を【どの業務領域に】【どう導入するか】
- （もう1点）

### 🔄 続けるべきこと
- （今月の良い傾向を数値根拠付きで）

### 💥 来月のベストアクション
- （1つだけ。最もインパクトが大きい改善）

### 📋 経営層サマリー
（数値を含む3行以内の要約。修飾語を使わず事実のみ）

## 出力ルール
- 前置き・挨拶・まとめ文は書くな
- 「〜と思います」「〜でしょう」は使うな。言い切れ
- 同じ内容を別の表現で繰り返すな
- ツール名は具体的に書け（例: ChatGPT、GitHub Copilot、Gemini、NotebookLM）
- 数値は元データから引用し、根拠のない数値を出すな`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': CONFIG.GEMINI_API_KEY },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 8192 }
      })
    });
    
    const json = JSON.parse(response.getContentText());
    let text = json.candidates?.[0]?.content?.parts?.[0]?.text || '分析結果を取得できませんでした';
    
    // トークン制限に到達した場合の警告
    const finishReason = json.candidates?.[0]?.finishReason;
    if (finishReason === 'MAX_TOKENS') {
      Logger.log('警告: 月次レポートがトークン上限に達しました');
      text += '\n\n⚠️ レポートがトークン上限に達したため、一部省略されている可能性があります。';
    }
    
    return text;
  } catch (e) {
    Logger.log('Gemini APIエラー: ' + e.message);
    return 'AI分析エラー: ' + e.message;
  }
}

/**
 * 月次レポートを出力
 */
function outputMonthlyReport(ss, summary, analysis) {
  let sheet = ss.getSheetByName(CONFIG.MONTHLY_REPORT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.MONTHLY_REPORT_SHEET);
  }
  
  const now = new Date();
  const row = [
    now,
    summary.period.start,
    summary.period.end,
    summary.totals.days,
    summary.totals.users,
    summary.totals.events,
    Math.round(summary.totals.totalMinutes / 60),
    summary.totals.aiUsingCount,
    summary.totals.aiNotUsingCount,
    summary.totals.aiPotentialCount,
    summary.averages.aiRate,
    analysis
  ];
  
  // ヘッダーがなければ追加
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '生成日時', '期間開始', '期間終了', '日数', 'ユーザー数', 
      '予定数', '稼働時間(h)', 'AI活用件数', 'AI未活用件数', 
      'AI余地件数', 'AI活用率(%)', 'AI分析レポート'
    ]);
  }
  
  sheet.appendRow(row);
}

/**
 * 月次レポートをメール送信
 */
function sendMonthlyReportEmail(summary, analysis) {
  const subject = `【月次AI活用レポート】${summary.period.start} ～ ${summary.period.end}`;
  const body = `
AI活用度測定 月次レポート

========================================
期間: ${summary.period.start} ～ ${summary.period.end}（${summary.totals.days}日間）
対象: ${summary.totals.users}名
========================================

【サマリー】
・予定数: ${summary.totals.events}件
・稼働時間: ${Math.round(summary.totals.totalMinutes / 60)}時間
・AI活用率: ${summary.averages.aiRate}%

【AI活用状況】
・活用中: ${summary.totals.aiUsingCount}件
・未活用: ${summary.totals.aiNotUsingCount}件
・余地あり: ${summary.totals.aiPotentialCount}件

========================================
【AI分析レポート】
========================================
${analysis}

----------------------------------------
このメールはGoogle Apps Scriptにより自動送信されました。
`;

  GmailApp.sendEmail(CONFIG.REPORT_EMAIL, subject, body);
}

/**
 * 月次トリガーを設定
 * 毎月1日の朝9時に実行
 */
function setupMonthlyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'generateMonthlyReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成（毎月1日）
  ScriptApp.newTrigger('generateMonthlyReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
    
  Logger.log('月次トリガーを設定しました（毎月1日9時）');
}

/**
 * カスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AI活用度分析')
    .addItem('週次レポート生成', 'generateWeeklyReport')
    .addItem('月次レポート生成', 'generateMonthlyReport')
    .addSeparator()
    .addItem('週次トリガー設定', 'setupWeeklyTrigger')
    .addItem('月次トリガー設定', 'setupMonthlyTrigger')
    .addToUi();
}

