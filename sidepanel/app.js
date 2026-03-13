// サイドパネル メインアプリケーション

import { getAuthToken, revokeToken, getUserInfo } from '../src/auth/oauth.js';
import { getEventsForDate, calculateTotalWorkTime, formatDuration, formatTime, getEventDuration, getCalendarList, saveSelectedCalendars, getSelectedCalendars } from '../src/api/calendar.js';
import { saveReport, ensureHeaders, saveSpreadsheetConfig, getSpreadsheetConfig, testSpreadsheetConnection, updateDailySummary, updateCalendarSummary, saveReportToSheet } from '../src/api/sheets.js';
import { saveToNotion, testNotionConnection, saveNotionConfig, getNotionConfig, saveDestinationConfig, getDestinationConfig } from '../src/api/notion.js';
import { analyzeWithGemini, saveGeminiConfig, getGeminiConfig, testGeminiConnection } from '../src/api/gemini.js';
import { estimateAIUsage, AI_FLAG_OPTIONS } from '../src/utils/ai-estimator.js';
import { addToQueue, getQueue, retryQueue, isOffline, watchOnlineStatus } from '../src/utils/offline-queue.js';
import { saveDraft, getDraft, deleteDraft, saveUserInfo, formatDate, formatDateJapanese } from '../src/utils/storage.js';
import { markdownToSafeHtml, sanitizeForPrompt, escapeHtml } from '../src/utils/sanitize.js';
import { initWizard, startWizard, updateWizardToken } from './wizard.js';

// グローバル状態
let currentDate = new Date();
let currentEvents = [];
let scheduleData = new Map();
let currentUser = null;
let authToken = null;

// DOM要素
const elements = {
    // スクリーン
    loginScreen: document.getElementById('login-screen'),
    mainScreen: document.getElementById('main-screen'),

    // ログイン
    loginBtn: document.getElementById('login-btn'),

    // ヘッダー
    userAvatar: document.getElementById('user-avatar'),
    userName: document.getElementById('user-name'),
    logoutBtn: document.getElementById('logout-btn'),
    settingsBtn: document.getElementById('settings-btn'),
    prevDay: document.getElementById('prev-day'),
    nextDay: document.getElementById('next-day'),
    todayBtn: document.getElementById('today-btn'),
    currentDate: document.getElementById('current-date'),
    eventCount: document.getElementById('event-count'),
    totalTime: document.getElementById('total-time'),
    avgRate: document.getElementById('avg-rate'),

    // メインコンテンツ
    loading: document.getElementById('loading'),
    emptyState: document.getElementById('empty-state'),
    scheduleList: document.getElementById('schedule-list'),

    // フッター
    autoEstimateBtn: document.getElementById('auto-estimate-btn'),
    selectAllBtn: document.getElementById('select-all-btn'),
    previewBtn: document.getElementById('preview-btn'),
    wizardBtn: document.getElementById('wizard-btn'),

    // プレビューモーダル
    previewModal: document.getElementById('preview-modal'),
    closeModal: document.getElementById('close-modal'),
    previewContent: document.getElementById('preview-content'),
    cancelBtn: document.getElementById('cancel-btn'),
    submitBtn: document.getElementById('submit-btn'),

    // 設定モーダル
    settingsModal: document.getElementById('settings-modal'),
    closeSettings: document.getElementById('close-settings'),
    toggleSheets: document.getElementById('toggle-sheets'),
    toggleNotion: document.getElementById('toggle-notion'),
    sheetsSettings: document.getElementById('sheets-settings'),
    notionSettings: document.getElementById('notion-settings'),
    spreadsheetId: document.getElementById('spreadsheet-id'),
    sheetName: document.getElementById('sheet-name'),
    sheetsStatus: document.getElementById('sheets-status'),
    notionToken: document.getElementById('notion-token'),
    notionDatabaseId: document.getElementById('notion-database-id'),
    notionStatus: document.getElementById('notion-status'),
    testNotionBtn: document.getElementById('test-notion-btn'),
    testSheetsBtn: document.getElementById('test-sheets-btn'),
    enableSheets: document.getElementById('enable-sheets'),
    enableNotion: document.getElementById('enable-notion'),
    geminiApiKey: document.getElementById('gemini-api-key'),
    geminiStatus: document.getElementById('gemini-status'),
    testGeminiBtn: document.getElementById('test-gemini-btn'),
    enableAiAnalysis: document.getElementById('enable-ai-analysis'),
    aiAnalysisModal: document.getElementById('ai-analysis-modal'),
    aiAnalysisContent: document.getElementById('ai-analysis-content'),
    closeAiAnalysis: document.getElementById('close-ai-analysis'),
    closeAiAnalysisBtn: document.getElementById('close-ai-analysis-btn'),
    // 自動日報・Slack通知設定
    enableDailyAlarm: document.getElementById('enable-daily-alarm'),
    alarmTime: document.getElementById('alarm-time'),
    enableSlackNotification: document.getElementById('enable-slack-notification'),
    slackWebhookUrl: document.getElementById('slack-webhook-url'),
    slackStatus: document.getElementById('slack-status'),
    testSlackBtn: document.getElementById('test-slack-btn'),
    //
    wizardProfileList: document.getElementById('wizard-profile-list'),
    wizardProfileCategoryInput: document.getElementById('wizard-profile-category-input'),
    wizardProfileOnelinersInput: document.getElementById('wizard-profile-oneliners-input'),
    addWizardProfileBtn: document.getElementById('add-wizard-profile-btn'),
    calendarList: document.getElementById('calendar-list'),
    refreshCalendarsBtn: document.getElementById('refresh-calendars-btn'),
    saveSettingsBtn: document.getElementById('save-settings-btn'),

    // ダッシュボード
    dashboardBtn: document.getElementById('dashboard-btn'),
    dashboardModal: document.getElementById('dashboard-modal'),
    closeDashboard: document.getElementById('close-dashboard'),
    closeDashboardBtn: document.getElementById('close-dashboard-btn'),
    statEvents: document.getElementById('stat-events'),
    statHours: document.getElementById('stat-hours'),
    statAiUsing: document.getElementById('stat-ai-using'),
    statAiRate: document.getElementById('stat-ai-rate'),
    generateWeeklyReport: document.getElementById('generate-weekly-report'),
    generateMonthlyReport: document.getElementById('generate-monthly-report'),
    generateDailySlackBtn: document.getElementById('generate-daily-slack-btn'),
    dashboardReportSectionWeekly: document.getElementById('dashboard-report-section-weekly'),
    dashboardReportContentWeekly: document.getElementById('dashboard-report-content-weekly'),
    dashboardReportSectionMonthly: document.getElementById('dashboard-report-section-monthly'),
    dashboardReportContentMonthly: document.getElementById('dashboard-report-content-monthly'),

    // AI分析レポート（メインページ）
    quickReportHeader: document.getElementById('quick-report-header'),
    quickReportBody: document.getElementById('quick-report-body'),
    quickReportDaily: document.getElementById('quick-report-daily'),
    quickReportWeekly: document.getElementById('quick-report-weekly'),
    quickReportMonthly: document.getElementById('quick-report-monthly'),
    quickReportResult: document.getElementById('quick-report-result'),
    quickReportText: document.getElementById('quick-report-text'),
    copyReportBtn: document.getElementById('copy-report-btn'),

    // トースト
    toast: document.getElementById('toast'),
    toastMessage: document.getElementById('toast-message')
};

// 初期化
async function init() {
    setupEventListeners();
    await checkAuth();
    watchOnlineStatus(handleOnlineStatusChange);
}

// イベントリスナー設定
function setupEventListeners() {
    // ログイン
    elements.loginBtn.addEventListener('click', handleLogin);
    elements.logoutBtn.addEventListener('click', handleLogout);

    // 設定
    elements.settingsBtn.addEventListener('click', showSettings);

    // 日付ナビゲーション
    elements.prevDay.addEventListener('click', () => changeDate(-1));
    elements.nextDay.addEventListener('click', () => changeDate(1));
    elements.todayBtn.addEventListener('click', goToToday);

    // フッターアクション
    elements.autoEstimateBtn.addEventListener('click', handleAutoEstimate);
    elements.selectAllBtn.addEventListener('click', handleSelectAll);
    elements.previewBtn.addEventListener('click', showPreview);
    elements.wizardBtn.addEventListener('click', () => {
        startWizard(currentEvents, scheduleData);
    });

    // ウィザード設定
    elements.addWizardProfileBtn.addEventListener('click', handleAddWizardProfile);

    // プレビューモーダル
    elements.closeModal.addEventListener('click', hidePreview);
    elements.cancelBtn.addEventListener('click', hidePreview);
    elements.submitBtn.addEventListener('click', handleSubmit);

    // 設定モーダル
    elements.closeSettings.addEventListener('click', hideSettings);
    elements.toggleSheets.addEventListener('click', () => switchSettingsTab('sheets'));
    elements.toggleNotion.addEventListener('click', () => switchSettingsTab('notion'));
    elements.testNotionBtn.addEventListener('click', handleTestNotion);
    elements.testSheetsBtn.addEventListener('click', handleTestSheets);
    elements.testGeminiBtn.addEventListener('click', handleTestGemini);
    elements.refreshCalendarsBtn.addEventListener('click', loadCalendarList);
    elements.saveSettingsBtn.addEventListener('click', handleSaveSettings);

    // AI分析モーダル
    elements.closeAiAnalysis.addEventListener('click', hideAiAnalysis);
    elements.closeAiAnalysisBtn.addEventListener('click', hideAiAnalysis);

    // AI分析レポート（メインページ）
    elements.quickReportHeader.addEventListener('click', () => {
        elements.quickReportBody.classList.toggle('hidden');
        // 矢印アイコンの回転
        const arrow = elements.quickReportHeader.querySelector('svg');
        if (arrow) arrow.style.transform = elements.quickReportBody.classList.contains('hidden') ? '' : 'rotate(180deg)';
    });
    elements.quickReportDaily.addEventListener('click', () => generateQuickReport('daily'));
    elements.quickReportWeekly.addEventListener('click', () => generateQuickReport('weekly'));
    elements.quickReportMonthly.addEventListener('click', () => generateQuickReport('monthly'));
    elements.copyReportBtn.addEventListener('click', () => {
        const text = elements.quickReportText.innerText;
        navigator.clipboard.writeText(text).then(() => showToast('コピーしました', 'success'));
    });

    // ダッシュボード
    elements.dashboardBtn.addEventListener('click', showDashboard);
    elements.closeDashboard.addEventListener('click', hideDashboard);
    elements.closeDashboardBtn.addEventListener('click', hideDashboard);
    elements.generateWeeklyReport.addEventListener('click', () => handleGenerateReport('weekly'));
    elements.generateMonthlyReport.addEventListener('click', () => handleGenerateReport('monthly'));
    elements.generateDailySlackBtn.addEventListener('click', handleGenerateDailySlack);
    elements.dashboardModal.querySelectorAll('.report-delete-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            if (e.currentTarget.dataset.type === 'weekly') {
                elements.dashboardReportSectionWeekly.style.display = 'none';
            } else {
                elements.dashboardReportSectionMonthly.style.display = 'none';
            }
        });
    });

    // チャート期間切替タブ
    document.querySelectorAll('.chart-tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            document.querySelectorAll('.chart-tab').forEach(t => t.classList.remove('active'));
            e.target.classList.add('active');
            renderAiTimeChart(e.target.dataset.period);
        });
    });
}

// 認証チェック
async function checkAuth() {
    try {
        authToken = await getAuthToken();
        currentUser = await getUserInfo(authToken);
        await saveUserInfo(currentUser);
        initWizard(authToken);
        showMainScreen();
        await loadEvents();
    } catch (error) {
        console.log('認証が必要です:', error);
        showLoginScreen();
    }
}

// ログイン処理
async function handleLogin() {
    try {
        elements.loginBtn.disabled = true;
        elements.loginBtn.textContent = 'ログイン中...';

        authToken = await getAuthToken();
        currentUser = await getUserInfo(authToken);
        await saveUserInfo(currentUser);
        initWizard(authToken);

        showMainScreen();
        await loadEvents();
    } catch (error) {
        console.error('ログインエラー:', error);
        showToast('ログインに失敗しました', 'error');
    } finally {
        elements.loginBtn.disabled = false;
        elements.loginBtn.innerHTML = `
      <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
        <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>
        <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>
        <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/>
        <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/>
      </svg>
      Googleでログイン
    `;
    }
}

// ログアウト処理
async function handleLogout() {
    try {
        await revokeToken();
        currentUser = null;
        authToken = null;
        showLoginScreen();
    } catch (error) {
        console.error('ログアウトエラー:', error);
        showToast('ログアウトに失敗しました', 'error');
    }
}

// 画面切り替え
function showLoginScreen() {
    elements.loginScreen.classList.remove('hidden');
    elements.mainScreen.classList.add('hidden');
}

function showMainScreen() {
    elements.loginScreen.classList.add('hidden');
    elements.mainScreen.classList.remove('hidden');

    if (currentUser) {
        elements.userAvatar.src = currentUser.picture || '';
        elements.userName.textContent = currentUser.name || currentUser.email;
    }
}

// 日付変更
function changeDate(delta) {
    currentDate.setDate(currentDate.getDate() + delta);
    loadEvents();
}

function goToToday() {
    currentDate = new Date();
    loadEvents();
}

// イベント読み込み
async function loadEvents() {
    elements.loading.classList.remove('hidden');
    elements.emptyState.classList.add('hidden');
    elements.scheduleList.innerHTML = '';

    elements.currentDate.textContent = formatDateJapanese(currentDate);

    try {
        const draft = await getDraft(formatDate(currentDate));
        if (draft) {
            scheduleData = new Map(Object.entries(draft));
        } else {
            scheduleData = new Map();
        }

        currentEvents = await getEventsForDate(authToken, currentDate);

        if (currentEvents.length === 0) {
            elements.emptyState.classList.remove('hidden');
        } else {
            renderScheduleList();
        }

        updateStats();
    } catch (error) {
        console.error('イベント取得エラー:', error);
        if (error.message && (error.message.includes('authentication credentials') || error.message.includes('OAuth 2'))) {
            showToast('認証の有効期限が切れました。再度ログインが必要です。', 'error');
            setTimeout(() => {
                handleLogout();
            }, 1000);
        } else {
            showToast('予定の取得に失敗しました', 'error');
        }
    } finally {
        elements.loading.classList.add('hidden');
    }
}

// 予定リストをレンダリング
function renderScheduleList() {
    elements.scheduleList.innerHTML = '';

    currentEvents.forEach((event, index) => {
        const eventId = event.id || `event-${index}`;
        const data = scheduleData.get(eventId) || {
            aiFlag: 'no',
            aiRate: 0,
            note: '',
            included: true
        };

        const item = createScheduleItem(event, eventId, data);
        elements.scheduleList.appendChild(item);
    });
}

// 予定アイテムを作成
function createScheduleItem(event, eventId, data) {
    const div = document.createElement('div');
    div.className = `schedule-item ${!data.included ? 'excluded' : ''}`;
    div.dataset.eventId = eventId;

    const startTime = formatTime(event.start.dateTime || event.start.date);
    const endTime = formatTime(event.end.dateTime || event.end.date);
    const duration = getEventDuration(event);
    const calendarColor = event.calendarColor || '#4285f4';

    const estimate = estimateAIUsage(event.summary || '');

    div.innerHTML = `
    <div class="schedule-item-header">
      <div class="schedule-color" data-color="${calendarColor}"></div>
      <div class="schedule-info">
        <div class="schedule-title">${escapeHtml(event.summary || '（タイトルなし）')}</div>
        <div class="schedule-time">
          ${startTime} - ${endTime}
          <span class="schedule-duration">(${formatDuration(duration)})</span>
        </div>
        ${event.calendarName ? `<div class="schedule-calendar">${escapeHtml(event.calendarName)}</div>` : ''}
        ${estimate.matchedKeyword ? `
          <span class="ai-estimate-badge ${estimate.confidence}">
            自動推定: ${escapeHtml(estimate.matchedKeyword)}
          </span>
        ` : ''}
      </div>
      <label class="schedule-include" title="日報に含める">
        <input type="checkbox" data-field="included" ${data.included ? 'checked' : ''}>
      </label>
    </div>
    <div class="schedule-controls">
      <div class="ai-flag-group">
        <label class="ai-flag-label">AI活用余地</label>
        <select class="ai-flag-select" data-field="aiFlag">
          ${AI_FLAG_OPTIONS.map(opt => `
            <option value="${opt.value}" ${data.aiFlag === opt.value ? 'selected' : ''}>${opt.label}</option>
          `).join('')}
        </select>
      </div>
      <div class="rate-group">
        <div class="rate-label">
          <span>活用率</span>
          <span class="rate-value">${data.aiRate}%</span>
        </div>
        <input type="range" class="rate-slider" data-field="aiRate" 
               min="0" max="100" step="5" value="${data.aiRate}">
        <div class="rate-presets">
          ${[0, 25, 50, 75, 100].map(val => `
            <button class="rate-preset ${data.aiRate === val ? 'active' : ''}" data-value="${val}">${val}%</button>
          `).join('')}
        </div>
      </div>
      <div class="memo-group">
        <label class="memo-toggle">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
          </svg>
          メモを追加
        </label>
        <textarea class="memo-input ${data.note ? '' : 'hidden'}" data-field="note" 
                  placeholder="使用したツール、プロンプト、生成物のリンクなど">${data.note}</textarea>
      </div>
    </div>
  `;

    setupItemEventListeners(div, eventId);

    return div;
}

// アイテムのイベントリスナー設定
function setupItemEventListeners(item, eventId) {
    const colorEl = item.querySelector('.schedule-color');
    if (colorEl && colorEl.dataset.color) {
        colorEl.style.backgroundColor = colorEl.dataset.color;
    }

    const checkbox = item.querySelector('[data-field="included"]');
    checkbox.addEventListener('change', (e) => {
        updateScheduleData(eventId, 'included', e.target.checked);
        item.classList.toggle('excluded', !e.target.checked);
    });

    const select = item.querySelector('[data-field="aiFlag"]');
    select.addEventListener('change', (e) => {
        updateScheduleData(eventId, 'aiFlag', e.target.value);
    });

    const slider = item.querySelector('[data-field="aiRate"]');
    const rateValue = item.querySelector('.rate-value');
    const presets = item.querySelectorAll('.rate-preset');

    slider.addEventListener('input', (e) => {
        const value = parseInt(e.target.value);
        rateValue.textContent = `${value}%`;
        updateScheduleData(eventId, 'aiRate', value);
        updatePresetButtons(presets, value);
    });

    presets.forEach(btn => {
        btn.addEventListener('click', () => {
            const value = parseInt(btn.dataset.value);
            slider.value = value;
            rateValue.textContent = `${value}%`;
            updateScheduleData(eventId, 'aiRate', value);
            updatePresetButtons(presets, value);
        });
    });

    const memoToggle = item.querySelector('.memo-toggle');
    const memoInput = item.querySelector('[data-field="note"]');

    memoToggle.addEventListener('click', () => {
        memoInput.classList.toggle('hidden');
        if (!memoInput.classList.contains('hidden')) {
            memoInput.focus();
        }
    });

    memoInput.addEventListener('input', (e) => {
        updateScheduleData(eventId, 'note', e.target.value);
    });
}

function updatePresetButtons(presets, value) {
    presets.forEach(btn => {
        btn.classList.toggle('active', parseInt(btn.dataset.value) === value);
    });
}

function updateScheduleData(eventId, field, value) {
    const data = scheduleData.get(eventId) || {
        aiFlag: 'no',
        aiRate: 0,
        note: '',
        included: true
    };
    data[field] = value;
    scheduleData.set(eventId, data);

    saveDraft(formatDate(currentDate), Object.fromEntries(scheduleData));
    updateStats();
}

function updateStats() {
    const includedEvents = currentEvents.filter((event, index) => {
        const eventId = event.id || `event-${index}`;
        const data = scheduleData.get(eventId);
        return !data || data.included !== false;
    });

    elements.eventCount.textContent = includedEvents.length;
    elements.totalTime.textContent = formatDuration(calculateTotalWorkTime(includedEvents));

    let totalRate = 0;
    let rateCount = 0;

    scheduleData.forEach((data) => {
        if (data.included !== false && data.aiFlag !== 'no') {
            totalRate += data.aiRate;
            rateCount++;
        }
    });

    const avgRate = rateCount > 0 ? Math.round(totalRate / rateCount) : 0;
    elements.avgRate.textContent = `${avgRate}%`;
}

function handleAutoEstimate() {
    currentEvents.forEach((event, index) => {
        const eventId = event.id || `event-${index}`;
        const estimate = estimateAIUsage(event.summary || '');

        if (estimate.flag !== 'unknown') {
            updateScheduleData(eventId, 'aiFlag', estimate.flag);
            if (estimate.suggestedRate > 0) {
                updateScheduleData(eventId, 'aiRate', estimate.suggestedRate);
            }
        }
    });

    renderScheduleList();
    showToast('自動推定を適用しました');
}

function handleSelectAll() {
    const allIncluded = Array.from(scheduleData.values()).every(d => d.included !== false);

    currentEvents.forEach((event, index) => {
        const eventId = event.id || `event-${index}`;
        updateScheduleData(eventId, 'included', !allIncluded);
    });

    renderScheduleList();
}

function showPreview() {
    const previewData = generatePreviewData();
    elements.previewContent.innerHTML = generatePreviewHTML(previewData);
    elements.previewModal.classList.remove('hidden');
}

function hidePreview() {
    elements.previewModal.classList.add('hidden');
}

function generatePreviewData() {
    const entries = [];
    let totalRate = 0;
    let rateCount = 0;

    currentEvents.forEach((event, index) => {
        const eventId = event.id || `event-${index}`;
        const data = scheduleData.get(eventId) || {
            aiFlag: 'no',
            aiRate: 0,
            note: '',
            included: true
        };

        if (data.included !== false) {
            entries.push({
                title: event.summary || '（タイトルなし）',
                start: formatTime(event.start.dateTime || event.start.date),
                end: formatTime(event.end.dateTime || event.end.date),
                calendar: event.calendarName || 'primary',
                aiFlag: data.aiFlag,
                aiRate: data.aiRate,
                note: data.note || ''
            });

            if (data.aiFlag !== 'no') {
                totalRate += data.aiRate;
                rateCount++;
            }
        }
    });

    return {
        date: formatDateJapanese(currentDate),
        dateFormatted: formatDate(currentDate),
        totalEvents: entries.length,
        totalTime: formatDuration(calculateTotalWorkTime(
            currentEvents.filter((e, i) => {
                const id = e.id || `event-${i}`;
                const d = scheduleData.get(id);
                return !d || d.included !== false;
            })
        )),
        avgRate: rateCount > 0 ? Math.round(totalRate / rateCount) : 0,
        entries
    };
}

function generatePreviewHTML(data) {
    return `
    <div class="preview-date">${data.date}</div>
    <div class="preview-summary">
      <div class="preview-summary-item">
        <span>予定数</span>
        <strong>${data.totalEvents}件</strong>
      </div>
      <div class="preview-summary-item">
        <span>合計時間</span>
        <strong>${data.totalTime}</strong>
      </div>
      <div class="preview-summary-item">
        <span>平均活用率</span>
        <strong>${data.avgRate}%</strong>
      </div>
    </div>
    <div class="preview-entries">
      ${data.entries.map(entry => `
        <div class="preview-entry">
          <div class="preview-entry-title">${escapeHtml(entry.title)}</div>
          <div class="preview-entry-details">
            ${entry.start} - ${entry.end} | 
            ${AI_FLAG_OPTIONS.find(o => o.value === entry.aiFlag)?.label || 'No'} | 
            活用率: ${entry.aiRate}%
            ${entry.note ? `<br>メモ: ${escapeHtml(entry.note)}` : ''}
          </div>
        </div>
      `).join('')}
    </div>
  `;
}

// 送信処理
async function handleSubmit() {
    elements.submitBtn.disabled = true;
    elements.submitBtn.textContent = '送信中...';

    try {
        const previewData = generatePreviewData();
        const destConfig = await getDestinationConfig();

        const reportData = {
            userEmail: currentUser.email,
            date: previewData.dateFormatted,
            totalWorkHours: previewData.totalTime,
            scheduleEntries: previewData.entries,
            generatedText: '',
            timestamp: new Date().toISOString()
        };

        if (isOffline()) {
            await addToQueue(reportData);
            showToast('オフラインのため、後で自動送信します', 'warning');
        } else {
            const results = [];

            // Googleスプレッドシートに保存
            if (destConfig.enableSheets) {
                try {
                    await ensureHeaders(authToken);
                    await saveReport(authToken, reportData);
                    // 日次集計シートも自動更新
                    await updateDailySummary(authToken, reportData);
                    // カレンダー別集計シートも自動更新
                    await updateCalendarSummary(authToken, reportData);
                    results.push('スプレッドシート');
                } catch (error) {
                    console.error('スプレッドシート保存エラー:', error);
                    showToast(`スプレッドシート保存失敗: ${error.message}`, 'error');
                }
            }

            // Notionに保存
            if (destConfig.enableNotion) {
                try {
                    await saveToNotion(reportData);
                    results.push('Notion');
                } catch (error) {
                    console.error('Notion保存エラー:', error);
                    // エラー詳細を表示
                    const errorMsg = error.message || 'Unknown error';
                    showToast(`Notion保存失敗: ${errorMsg}`, 'error');
                    console.log('Notion Error Details:', JSON.stringify(error, null, 2));
                }
            }

            if (results.length > 0) {
                showToast(`日報を保存しました (${results.join(', ')})`, 'success');
                await deleteDraft(previewData.dateFormatted);

                // AI分析を実行
                const { enableAiAnalysis } = await chrome.storage.local.get(['enableAiAnalysis']);
                if (enableAiAnalysis) {
                    try {
                        showToast('AI分析中...', 'default');
                        const analysisResult = await runAiAnalysis(reportData);
                        if (analysisResult.success) {
                            showAiAnalysis(analysisResult.analysis);
                        }
                    } catch (error) {
                        console.error('AI分析エラー:', error);
                        showToast(`AI分析エラー: ${error.message}`, 'error');
                    }
                }
            } else if (!destConfig.enableSheets && !destConfig.enableNotion) {
                showToast('保存先が設定されていません。設定画面から設定してください。', 'warning');
            }
        }

        hidePreview();
    } catch (error) {
        console.error('送信エラー:', error);

        const previewData = generatePreviewData();
        await addToQueue({
            userEmail: currentUser.email,
            date: previewData.dateFormatted,
            totalWorkHours: previewData.totalTime,
            scheduleEntries: previewData.entries,
            timestamp: new Date().toISOString()
        });

        showToast('送信に失敗しました。後で再送信します', 'error');
    } finally {
        elements.submitBtn.disabled = false;
        elements.submitBtn.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M22 2L11 13"/>
        <path d="M22 2l-7 20-4-9-9-4 20-7z"/>
      </svg>
      送信
    `;
    }
}

// 設定モーダル
async function showSettings() {
    // スプレッドシート設定を読み込み
    const sheetsConfig = await getSpreadsheetConfig();
    elements.spreadsheetId.value = sheetsConfig.spreadsheetId || '';
    elements.sheetName.value = sheetsConfig.sheetName || 'Sheet1';

    // Notion設定を読み込み
    const notionConfig = await getNotionConfig();
    elements.notionToken.value = notionConfig.notionToken || '';
    elements.notionDatabaseId.value = notionConfig.notionDatabaseId || '';

    // Gemini設定を読み込み
    const geminiConfig = await getGeminiConfig();
    elements.geminiApiKey.value = geminiConfig.geminiApiKey || '';

    // 送信先設定を読み込み
    const destConfig = await getDestinationConfig();
    elements.enableSheets.checked = destConfig.enableSheets;
    elements.enableNotion.checked = destConfig.enableNotion;

    // AI分析設定を読み込み
    const { enableAiAnalysis } = await chrome.storage.local.get(['enableAiAnalysis']);
    elements.enableAiAnalysis.checked = enableAiAnalysis === true;

    // 自動日報・Slack通知設定を読み込み
    const alarmConfig = await chrome.storage.local.get(['enableDailyAlarm', 'alarmTime', 'enableSlackNotification', 'slackWebhookUrl']);
    elements.enableDailyAlarm.checked = alarmConfig.enableDailyAlarm === true;
    elements.enableSlackNotification.checked = alarmConfig.enableSlackNotification !== false; // デフォルトは互換性重視でON気味にするか、設定がなければundefined

    // alarmTime は "HH:MM" 形式が望ましいが、旧データの number(例:18) にも対応させる
    let timeVal = "18:00";
    if (alarmConfig.alarmTime !== undefined) {
        if (typeof alarmConfig.alarmTime === 'number') {
            timeVal = alarmConfig.alarmTime.toString().padStart(2, '0') + ':00';
        } else {
            timeVal = alarmConfig.alarmTime;
        }
    }
    elements.alarmTime.value = timeVal;
    elements.slackWebhookUrl.value = alarmConfig.slackWebhookUrl || '';

    // ウィザード設定（カスタムプロファイル）を読み込み
    await loadAndRenderWizardProfiles();

    // 接続状態を更新
    updateConnectionStatus('sheets', sheetsConfig.spreadsheetId ? 'connected' : 'disconnected');
    updateConnectionStatus('notion', notionConfig.notionToken ? 'connected' : 'disconnected');
    updateConnectionStatus('gemini', geminiConfig.geminiApiKey ? 'connected' : 'disconnected');
    updateConnectionStatus('slack', elements.slackWebhookUrl.value ? 'connected' : 'disconnected');

    elements.settingsModal.classList.remove('hidden');

    // カレンダー一覧を読み込み
    await loadCalendarList();
}

// ===== ウィザード設定管理 =====
let wizardProfiles = []; // [{ category: '...', oneliners: ['...'] }]

async function loadAndRenderWizardProfiles() {
    const { wizardWorkProfile } = await chrome.storage.local.get(['wizardWorkProfile']);
    // 'カテゴリ: 一言1, 一言2' の形式からパース
    const lines = (wizardWorkProfile || '').split('\n').map(l => l.trim()).filter(Boolean);
    wizardProfiles = lines.map(line => {
        const parts = line.split(/[:：]/);
        if (parts.length >= 2) {
            const category = parts[0].trim();
            const oneliners = parts.slice(1).join(':').split(/[,、]/).map(s => s.trim()).filter(Boolean);
            return { category, oneliners };
        }
        return { category: line, oneliners: [] };
    });
    renderWizardProfiles();
}

function renderWizardProfiles() {
    if (wizardProfiles.length === 0) {
        elements.wizardProfileList.innerHTML = '<div class="empty-state-small">設定はありません</div>';
        return;
    }

    elements.wizardProfileList.innerHTML = wizardProfiles.map((p, index) => `
        <div class="group-item" style="padding: 8px; margin-bottom: 8px; background: var(--bg-primary); border: 1px solid var(--border-color); border-radius: var(--border-radius); display: flex; justify-content: space-between; align-items: flex-start;">
            <div style="flex: 1;">
                <div style="font-weight: 500; font-size: 13px; margin-bottom: 4px;">${escapeHtml(p.category)}</div>
                <div style="font-size: 11px; color: var(--text-secondary);">${escapeHtml(p.oneliners.join(', '))}</div>
            </div>
            <button class="btn btn-icon btn-small delete-wizard-profile-btn" data-index="${index}" style="color: var(--color-error); padding: 4px;">✕</button>
        </div>
    `).join('');

    // 削除ボタンのイベントリスナー
    elements.wizardProfileList.querySelectorAll('.delete-wizard-profile-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.currentTarget.dataset.index, 10);
            deleteWizardProfile(idx);
        });
    });
}

function handleAddWizardProfile() {
    const category = elements.wizardProfileCategoryInput.value.trim();
    const onelinerStr = elements.wizardProfileOnelinersInput.value.trim();

    if (!category || !onelinerStr) {
        showToast('業務名と一言を入力してください', 'error');
        return;
    }

    const oneliners = onelinerStr.split(/[,、]/).map(s => s.trim()).filter(Boolean);
    if (oneliners.length === 0) {
        showToast('一言を入力してください', 'error');
        return;
    }

    wizardProfiles.push({ category, oneliners });
    elements.wizardProfileCategoryInput.value = '';
    elements.wizardProfileOnelinersInput.value = '';
    renderWizardProfiles();
    showToast('リストに追加しました（保存ボタンで確定します）', 'success');
}

function deleteWizardProfile(index) {
    wizardProfiles.splice(index, 1);
    renderWizardProfiles();
}
// =============================

// カレンダー一覧を読み込んで表示
async function loadCalendarList() {
    elements.calendarList.innerHTML = '<div class="loading-text">カレンダーを読み込み中...</div>';

    try {
        if (!authToken) {
            authToken = await getAuthToken();
        }

        const calendars = await getCalendarList(authToken);
        const selectedCalendars = await getSelectedCalendars();

        if (calendars.length === 0) {
            elements.calendarList.innerHTML = '<div class="loading-text">カレンダーが見つかりません</div>';
            return;
        }

        elements.calendarList.innerHTML = calendars
            .filter(cal => cal.accessRole !== 'freeBusyReader')
            .map(cal => {
                const isChecked = selectedCalendars.length === 0 || selectedCalendars.includes(cal.id);
                const color = cal.backgroundColor || '#4285f4';
                return `
                    <label class="calendar-item">
                        <input type="checkbox" value="${cal.id}" ${isChecked ? 'checked' : ''}>
                        <span class="calendar-color" data-color="${color}"></span>
                        <span class="calendar-name">${escapeHtml(cal.summary)}</span>
                    </label>
                `;
            }).join('');

        elements.calendarList.querySelectorAll('.calendar-color[data-color]').forEach(el => {
            el.style.backgroundColor = el.dataset.color;
        });

    } catch (error) {
        console.error('カレンダー一覧の取得に失敗:', error);
        elements.calendarList.innerHTML = '<div class="loading-text">カレンダーの読み込みに失敗しました</div>';
    }
}

function hideSettings() {
    elements.settingsModal.classList.add('hidden');
}

function switchSettingsTab(tab) {
    if (tab === 'sheets') {
        elements.toggleSheets.classList.add('active');
        elements.toggleNotion.classList.remove('active');
        elements.sheetsSettings.classList.remove('hidden');
        elements.notionSettings.classList.add('hidden');
    } else {
        elements.toggleSheets.classList.remove('active');
        elements.toggleNotion.classList.add('active');
        elements.sheetsSettings.classList.add('hidden');
        elements.notionSettings.classList.remove('hidden');
    }
}

async function handleTestNotion() {
    elements.testNotionBtn.disabled = true;
    elements.testNotionBtn.textContent = 'テスト中...';

    // 一時的に保存してテスト
    const token = elements.notionToken.value.trim();
    const databaseId = elements.notionDatabaseId.value.trim();

    if (!token || !databaseId) {
        showToast('トークンとデータベースIDを入力してください', 'error');
        elements.testNotionBtn.disabled = false;
        elements.testNotionBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/>
      </svg>
      接続テスト
    `;
        return;
    }

    await saveNotionConfig(token, databaseId);
    const result = await testNotionConnection();

    if (result.success) {
        updateConnectionStatus('notion', 'connected', `接続成功: ${result.databaseTitle}`);
        showToast(`Notionに接続しました: ${result.databaseTitle}`, 'success');
    } else {
        updateConnectionStatus('notion', 'error', `接続失敗: ${result.error}`);
        showToast(`Notion接続エラー: ${result.error}`, 'error');
    }

    elements.testNotionBtn.disabled = false;
    elements.testNotionBtn.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
      <polyline points="22 4 12 14.01 9 11.01"/>
    </svg>
    接続テスト
  `;
}

async function handleTestSheets() {
    elements.testSheetsBtn.disabled = true;
    elements.testSheetsBtn.textContent = 'テスト中...';

    // 一時的に保存してテスト
    const spreadsheetId = elements.spreadsheetId.value.trim();
    const sheetName = elements.sheetName.value.trim() || 'Sheet1';

    if (!spreadsheetId) {
        showToast('スプレッドシートIDを入力してください', 'error');
        elements.testSheetsBtn.disabled = false;
        elements.testSheetsBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/>
      </svg>
      接続テスト
    `;
        return;
    }

    await saveSpreadsheetConfig(spreadsheetId, sheetName);
    const result = await testSpreadsheetConnection(authToken);

    if (result.success) {
        updateConnectionStatus('sheets', 'connected', result.message);
        showToast(result.message, 'success');
    } else {
        updateConnectionStatus('sheets', 'error', `接続失敗: ${result.error}`);
        showToast(`スプレッドシート接続エラー: ${result.error}`, 'error');
    }

    elements.testSheetsBtn.disabled = false;
    elements.testSheetsBtn.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
      <polyline points="22 4 12 14.01 9 11.01"/>
    </svg>
    接続テスト
  `;
}

async function handleTestSlack() {
    elements.testSlackBtn.disabled = true;
    elements.testSlackBtn.textContent = 'テスト中...';

    const webhookUrl = elements.slackWebhookUrl.value.trim();

    if (!webhookUrl) {
        showToast('Slack Webhook URLを入力してください', 'error');
        elements.testSlackBtn.disabled = false;
        elements.testSlackBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/>
      </svg>
      接続テスト
    `;
        return;
    }

    try {
        const response = await fetch(webhookUrl, {
            method: 'POST',
            body: JSON.stringify({ text: 'Daily Report AI: Slack接続テスト成功！' }),
        });

        if (response.ok) {
            updateConnectionStatus('slack', 'connected', '接続成功');
            showToast('Slackに接続しました', 'success');
        } else {
            const errorText = await response.text();
            updateConnectionStatus('slack', 'error', `接続失敗: ${response.status} ${errorText}`);
            showToast(`Slack接続エラー: ${response.status} ${errorText}`, 'error');
        }
    } catch (error) {
        updateConnectionStatus('slack', 'error', `接続失敗: ${error.message}`);
        showToast(`Slack接続エラー: ${error.message}`, 'error');
    } finally {
        elements.testSlackBtn.disabled = false;
        elements.testSlackBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/>
      </svg>
      接続テスト
    `;
    }
}

async function handleGenerateDailySlack() {
    elements.generateDailySlackBtn.disabled = true;
    const originalText = elements.generateDailySlackBtn.innerHTML;
    elements.generateDailySlackBtn.textContent = '生成＆送信中...';
    showToast('Slackへの日報送信リクエストを開始しました...', 'info');

    try {
        const response = await chrome.runtime.sendMessage({ action: 'generateDailySlack' });
        if (response && response.success) {
            showToast('本日の日報をSlackへ送信しました', 'success');
        } else {
            showToast(`送信に失敗しました: ${response ? response.error : '不明なエラー'}`, 'error');
        }
    } catch (e) {
        showToast(`送信エラー: ${e.message}`, 'error');
    } finally {
        elements.generateDailySlackBtn.disabled = false;
        elements.generateDailySlackBtn.innerHTML = originalText;
    }
}

function updateConnectionStatus(type, status, message) {
    let statusEl;
    if (type === 'sheets') statusEl = elements.sheetsStatus;
    else if (type === 'notion') statusEl = elements.notionStatus;
    else if (type === 'gemini') statusEl = elements.geminiStatus;
    else if (type === 'slack') statusEl = elements.slackStatus;
    else return;

    const iconEl = statusEl.querySelector('.status-icon');
    const textEl = statusEl.querySelector('.status-text');

    statusEl.className = `connection-status ${status}`;

    if (status === 'connected') {
        iconEl.textContent = '●';
        textEl.textContent = message || '接続済み';
    } else if (status === 'error') {
        iconEl.textContent = '✕';
        textEl.textContent = message || 'エラー';
    } else {
        iconEl.textContent = '○';
        textEl.textContent = message || '未接続';
    }
}

async function handleSaveSettings() {
    const spreadsheetId = elements.spreadsheetId.value.trim();
    const sheetName = elements.sheetName.value.trim() || 'Sheet1';
    const notionToken = elements.notionToken.value.trim();
    const notionDatabaseId = elements.notionDatabaseId.value.trim();
    const geminiApiKey = elements.geminiApiKey.value.trim();
    const enableSheets = elements.enableSheets.checked;
    const enableNotion = elements.enableNotion.checked;
    const enableAiAnalysis = elements.enableAiAnalysis.checked;

    // 自動日報・通知設定
    const enableDailyAlarm = elements.enableDailyAlarm.checked;
    const alarmTime = elements.alarmTime.value || '18:00';
    const enableSlackNotification = elements.enableSlackNotification.checked;
    const slackWebhookUrl = elements.slackWebhookUrl.value.trim();

    // バリデーション
    if (enableSheets && !spreadsheetId) {
        showToast('スプレッドシートIDを入力してください', 'error');
        return;
    }

    if (enableNotion && (!notionToken || !notionDatabaseId)) {
        showToast('NotionのトークンとデータベースIDを入力してください', 'error');
        return;
    }

    if (enableAiAnalysis && !geminiApiKey) {
        showToast('AI分析を有効にするにはGemini API Keyを入力してください', 'error');
        return;
    }

    // 保存
    await saveSpreadsheetConfig(spreadsheetId, sheetName);
    await saveNotionConfig(notionToken, notionDatabaseId);
    await saveGeminiConfig(geminiApiKey);
    await saveDestinationConfig(enableSheets, enableNotion);
    await chrome.storage.local.set({
        enableAiAnalysis,
        enableDailyAlarm,
        alarmTime,
        enableSlackNotification,
        slackWebhookUrl
    });

    // バックグラウンドに設定変更を通知（アラームの再設定を促す）
    chrome.runtime.sendMessage({ action: 'updateAlarms' });

    // wizardProfiles を文字列形式に戻して保存
    const wizardWorkProfile = wizardProfiles
        .map(p => `${p.category}: ${p.oneliners.join(', ')}`)
        .join('\n');
    await chrome.storage.local.set({ wizardWorkProfile });

    // カレンダー選択を保存
    const calendarCheckboxes = elements.calendarList.querySelectorAll('input[type="checkbox"]:checked');
    const selectedCalendarIds = Array.from(calendarCheckboxes).map(cb => cb.value);
    await saveSelectedCalendars(selectedCalendarIds);

    showToast('設定を保存しました', 'success');
    hideSettings();

    // カレンダー設定が変更された場合は予定を再読み込み
    await loadEvents();
}

async function handleTestGemini() {
    elements.testGeminiBtn.disabled = true;
    elements.testGeminiBtn.textContent = 'テスト中...';

    const apiKey = elements.geminiApiKey.value.trim();

    if (!apiKey) {
        showToast('API Keyを入力してください', 'error');
        elements.testGeminiBtn.disabled = false;
        elements.testGeminiBtn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/>
      </svg>
      接続テスト
    `;
        return;
    }

    await saveGeminiConfig(apiKey);
    const result = await testGeminiConnection();

    if (result.success) {
        updateConnectionStatus('gemini', 'connected', result.message);
        showToast(result.message, 'success');
    } else {
        updateConnectionStatus('gemini', 'error', `接続失敗: ${result.error}`);
        showToast(`Gemini接続エラー: ${result.error}`, 'error');
    }

    elements.testGeminiBtn.disabled = false;
    elements.testGeminiBtn.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
      <polyline points="22 4 12 14.01 9 11.01"/>
    </svg>
    接続テスト
  `;
}

// AI分析の実行
async function runAiAnalysis(reportData) {
    const summaryData = {
        period: 'daily',
        date: reportData.date,
        userEmail: reportData.userEmail,
        totalEvents: reportData.scheduleEntries.length,
        totalMinutes: 0,
        aiUsingCount: 0,
        aiNotUsingCount: 0,
        aiPotentialCount: 0,
        aiUsingMinutes: 0,
        aiNotUsingMinutes: 0,
        aiPotentialMinutes: 0,
        avgRate: 0,
        totalTime: reportData.totalWorkHours,
        entries: reportData.scheduleEntries
    };

    // 集計
    let totalRate = 0;
    let rateCount = 0;
    for (const entry of reportData.scheduleEntries) {
        const duration = calculateEntryMinutes(entry);
        summaryData.totalMinutes += duration;

        if (entry.aiFlag === 'yes-using') {
            summaryData.aiUsingCount++;
            summaryData.aiUsingMinutes += duration;
            totalRate += entry.aiRate;
            rateCount++;
        } else if (entry.aiFlag === 'yes-potential') {
            summaryData.aiPotentialCount++;
            summaryData.aiPotentialMinutes += duration;
            totalRate += entry.aiRate;
            rateCount++;
        } else {
            summaryData.aiNotUsingCount++;
            summaryData.aiNotUsingMinutes += duration;
        }
    }
    summaryData.avgRate = rateCount > 0 ? Math.round(totalRate / rateCount) : 0;

    // AI分析実行
    const result = await analyzeWithGemini(summaryData);
    return result;
}

function calculateEntryMinutes(entry) {
    try {
        const [startH, startM] = entry.start.split(':').map(Number);
        const [endH, endM] = entry.end.split(':').map(Number);
        const startTotal = startH * 60 + startM;
        const endTotal = endH * 60 + endM;
        return (endTotal >= startTotal) ? (endTotal - startTotal) : ((endTotal + 24 * 60) - startTotal);
    } catch {
        return 0;
    }
}

function showAiAnalysis(analysisText) {
    // マークダウンを簡易的にHTMLに変換 (XSS対策済み)
    const html = markdownToSafeHtml(analysisText);

    elements.aiAnalysisContent.innerHTML = html;
    elements.aiAnalysisModal.classList.remove('hidden');
}

function hideAiAnalysis() {
    elements.aiAnalysisModal.classList.add('hidden');
}

async function handleOnlineStatusChange(isOnline) {
    if (isOnline) {
        const queue = await getQueue();
        if (queue.length > 0) {
            showToast(`オンライン復帰: ${queue.length}件のデータを送信中...`);
            await retryQueue();
        }
    }
}

function showToast(message, type = 'default') {
    elements.toast.className = `toast ${type}`;
    elements.toastMessage.textContent = message;
    elements.toast.classList.remove('hidden');

    setTimeout(() => {
        elements.toast.classList.add('hidden');
    }, 3000);
}

// ダッシュボード表示
async function showDashboard() {
    elements.dashboardModal.classList.remove('hidden');
    elements.dashboardReportSectionWeekly.style.display = 'none';
    elements.dashboardReportSectionMonthly.style.display = 'none';

    // サマリーをデフォルト（日別）で表示
    await updateSummaryStats('daily');

    // サマリータブのイベントリスナー
    document.querySelectorAll('#summary-period-tabs .chart-tab').forEach(tab => {
        tab.onclick = async () => {
            document.querySelectorAll('#summary-period-tabs .chart-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            await updateSummaryStats(tab.dataset.summaryPeriod);
        };
    });

    // グラフを描画（デフォルト: 日別）
    await renderAiTimeChart('daily');
    // タブをリセット
    document.querySelectorAll('.chart-period-tabs:not(#summary-period-tabs) .chart-tab').forEach(t => t.classList.remove('active'));
    document.querySelector('.chart-period-tabs:not(#summary-period-tabs) .chart-tab[data-period="daily"]')?.classList.add('active');

    // カレンダー別ランキングを描画
    await renderCalendarRanking();
}

// サマリー統計を期間別に更新
async function updateSummaryStats(period) {
    const periodLabel = document.getElementById('summary-period-label');
    const elImpactBar = document.getElementById('impact-stats-bar');
    const elImpactTimeSaved = document.getElementById('impact-time-saved');
    const elImpactHighScore = document.getElementById('impact-high-score');

    if (period === 'daily') {
        // 今日のデータ（currentEventsから）
        if (periodLabel) periodLabel.textContent = '本日';

        if (currentEvents.length > 0) {
            let totalMinutes = 0;
            let aiUsingCount = 0;
            let eventCount = 0;

            for (const event of currentEvents) {
                const eventId = event.id || `event-${currentEvents.indexOf(event)}`;
                const data = scheduleData.get(eventId);
                // チェックが入っていない予定はスキップ
                if (data && data.included === false) continue;

                eventCount++;
                const duration = getEventDuration(event);
                totalMinutes += duration;
                const estimation = estimateAIUsage(event.summary || '');
                if (estimation.flag === 'yes-using') {
                    aiUsingCount++;
                }
            }

            const hours = Math.floor(totalMinutes / 60);
            const mins = totalMinutes % 60;
            const aiRate = eventCount > 0 ? Math.round((aiUsingCount / eventCount) * 100) : 0;

            elements.statEvents.textContent = eventCount;
            elements.statHours.textContent = hours > 0 ? `${hours}h${mins}m` : `${mins}m`;
            elements.statAiUsing.textContent = aiUsingCount;
            elements.statAiRate.textContent = `${aiRate}%`;
        } else {
            elements.statEvents.textContent = '-';
            elements.statHours.textContent = '-';
            elements.statAiUsing.textContent = '-';
            elements.statAiRate.textContent = '-';
        }

        // 本日分のAIインパクトログ（スコープ機能がない前提）
        if (elImpactBar) {
            const impactData = await fetchAiImpactData('personal');
            const now = new Date();
            const todayStr = `${now.getFullYear()}/${(now.getMonth() + 1).toString().padStart(2, '0')}/${now.getDate().toString().padStart(2, '0')}`;

            let totalTimeSaved = 0;
            let highScoreCount = 0;

            impactData.forEach(row => {
                const rDate = row.date.split('T')[0].replace(/-/g, '/');
                if (rDate === todayStr || row.date.includes(todayStr)) {
                    totalTimeSaved += (row.timeSaved || 0);
                    if (row.impactScore && row.impactScore >= 4) {
                        highScoreCount++;
                    }
                }
            });

            if (totalTimeSaved > 0 || highScoreCount > 0) {
                elImpactBar.style.display = 'flex';
                const sHours = Math.floor(totalTimeSaved / 60);
                const sMins = totalTimeSaved % 60;
                elImpactTimeSaved.textContent = sHours > 0 ? `${sHours}時間${sMins}分` : `${sMins}分`;
                elImpactHighScore.textContent = `${highScoreCount}件`;
            } else {
                elImpactBar.style.display = 'none';
            }
        }

    } else {
        // 週別・月別はDailySummaryシートから集計
        const daysAgo = period === 'weekly' ? 7 : 30;
        const now = new Date();
        const startDate = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);

        const formatDateShort = (d) => `${d.getMonth() + 1}/${d.getDate()}`;
        if (periodLabel) {
            periodLabel.textContent = `${formatDateShort(startDate)} ～ ${formatDateShort(now)}（${period === 'weekly' ? '過去7日' : '過去30日'}）`;
        }

        const rawData = await fetchDailySummaryData();
        const periodData = rawData.filter(row => {
            const d = new Date(row.date);
            return !isNaN(d.getTime()) && d >= startDate && d <= now;
        });

        if (periodData.length > 0) {
            const totals = periodData.reduce((acc, row) => {
                acc.events += row.eventCount;
                acc.totalMinutes += row.totalMinutes;
                acc.aiUsingCount += row.aiUsingCount;
                return acc;
            }, { events: 0, totalMinutes: 0, aiUsingCount: 0 });

            const totalAiRelated = periodData.reduce((acc, row) => {
                return acc + row.aiUsingCount + row.aiNotUsingCount + row.aiPotentialCount;
            }, 0);

            const hours = Math.floor(totals.totalMinutes / 60);
            const mins = totals.totalMinutes % 60;
            const aiRate = totalAiRelated > 0 ? Math.round((totals.aiUsingCount / totalAiRelated) * 100) : 0;

            elements.statEvents.textContent = `${totals.events}件`;
            elements.statHours.textContent = hours > 0 ? `${hours}h${mins}m` : `${mins}m`;
            elements.statAiUsing.textContent = `${totals.aiUsingCount}件`;
            elements.statAiRate.textContent = `${aiRate}%`;
        } else {
            elements.statEvents.textContent = '-';
            elements.statHours.textContent = '-';
            elements.statAiUsing.textContent = '-';
            elements.statAiRate.textContent = '-';
        }

        // 期間分のAIインパクトログ（スコープ対応必要なら追加）
        if (elImpactBar) {
            const impactData = await fetchAiImpactData('personal'); // FIXME: if dashboard has scopes later
            let totalTimeSaved = 0;
            let highScoreCount = 0;

            impactData.forEach(row => {
                if (!row.date) return;
                const rDate = new Date(row.date);
                if (rDate >= startDate && rDate <= now) {
                    totalTimeSaved += (row.timeSaved || 0);
                    if (row.impactScore && row.impactScore >= 4) {
                        highScoreCount++;
                    }
                }
            });

            if (totalTimeSaved > 0 || highScoreCount > 0) {
                elImpactBar.style.display = 'flex';
                const sHours = Math.floor(totalTimeSaved / 60);
                const sMins = totalTimeSaved % 60;
                elImpactTimeSaved.textContent = sHours > 0 ? `${sHours}時間${sMins}分` : `${sMins}分`;
                elImpactHighScore.textContent = `${highScoreCount}件`;
            } else {
                elImpactBar.style.display = 'none';
            }
        }
    }
}

// グローバルなチャートインスタンス
let aiTimeChartInstance = null;

// DailySummaryシートからデータを取得
async function fetchDailySummaryData() {
    try {
        const { spreadsheetId } = await chrome.storage.local.get(['spreadsheetId']);
        if (!spreadsheetId || !authToken) return [];

        const sheetName = 'DailySummary';
        const range = `${sheetName}!A:L`;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${authToken}` }
        });

        if (!response.ok) return [];

        const data = await response.json();
        const values = data.values || [];
        if (values.length <= 1) return []; // ヘッダーのみ

        return values.slice(1).map(row => ({
            date: row[0] || '',
            user: row[1] || '',
            eventCount: parseInt(row[2]) || 0,
            totalMinutes: parseInt(row[3]) || 0,
            aiUsingCount: parseInt(row[4]) || 0,
            aiNotUsingCount: parseInt(row[5]) || 0,
            aiPotentialCount: parseInt(row[6]) || 0,
            avgRate: parseInt(row[7]) || 0,
            aiUsingMinutes: parseInt(row[8]) || 0,
            aiNotUsingMinutes: parseInt(row[9]) || 0,
            aiPotentialMinutes: parseInt(row[10]) || 0
        }));
    } catch (e) {
        console.error('DailySummaryデータ取得エラー:', e);
        return [];
    }
}

// AI活用時間チャートを描画
async function renderAiTimeChart(period) {
    const canvas = document.getElementById('ai-time-chart');
    const noDataEl = document.getElementById('chart-no-data');
    if (!canvas) return;

    // 既存チャートを破棄
    if (aiTimeChartInstance) {
        aiTimeChartInstance.destroy();
        aiTimeChartInstance = null;
    }

    const rawData = await fetchDailySummaryData();

    if (rawData.length === 0) {
        canvas.style.display = 'none';
        noDataEl.classList.remove('hidden');
        return;
    }

    canvas.style.display = 'block';
    noDataEl.classList.add('hidden');

    const { labels, aiUsing, aiNotUsing, aiPotential } = aggregateChartData(rawData, period);

    const ctx = canvas.getContext('2d');
    aiTimeChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: 'AI活用中',
                    data: aiUsing,
                    backgroundColor: 'rgba(99, 102, 241, 0.85)',
                    borderRadius: 3
                },
                {
                    label: 'AI余地あり',
                    data: aiPotential,
                    backgroundColor: 'rgba(245, 158, 11, 0.7)',
                    borderRadius: 3
                },
                {
                    label: 'AI未活用',
                    data: aiNotUsing,
                    backgroundColor: 'rgba(209, 213, 219, 0.6)',
                    borderRadius: 3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        boxWidth: 12,
                        padding: 12,
                        font: { size: 11, family: 'Inter' }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y}h`
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: { display: false },
                    ticks: { font: { size: 10, family: 'Inter' } }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: '時間 (h)',
                        font: { size: 11, family: 'Inter' }
                    },
                    ticks: { font: { size: 10, family: 'Inter' } }
                }
            }
        }
    });
}

// データを期間別に集約
function aggregateChartData(data, period) {
    const now = new Date();
    let grouped = {};

    if (period === 'daily') {
        // 直近14日
        for (let i = 13; i >= 0; i--) {
            const d = new Date(now);
            d.setDate(d.getDate() - i);
            const key = `${d.getMonth() + 1}/${d.getDate()}`;
            grouped[key] = { aiUsing: 0, aiNotUsing: 0, aiPotential: 0 };
        }
        for (const row of data) {
            const d = new Date(row.date);
            if (isNaN(d.getTime())) continue;
            const key = `${d.getMonth() + 1}/${d.getDate()}`;
            if (grouped[key]) {
                grouped[key].aiUsing += row.aiUsingMinutes / 60;
                grouped[key].aiNotUsing += row.aiNotUsingMinutes / 60;
                grouped[key].aiPotential += row.aiPotentialMinutes / 60;
            }
        }
    } else if (period === 'weekly') {
        // 直近8週
        for (let i = 7; i >= 0; i--) {
            const weekStart = new Date(now);
            weekStart.setDate(weekStart.getDate() - weekStart.getDay() - i * 7 + 1);
            const key = `${weekStart.getMonth() + 1}/${weekStart.getDate()}週`;
            grouped[key] = { aiUsing: 0, aiNotUsing: 0, aiPotential: 0, start: new Date(weekStart), end: new Date(weekStart) };
            grouped[key].end.setDate(grouped[key].end.getDate() + 6);
        }
        for (const row of data) {
            const d = new Date(row.date);
            if (isNaN(d.getTime())) continue;
            for (const [key, g] of Object.entries(grouped)) {
                if (d >= g.start && d <= g.end) {
                    g.aiUsing += row.aiUsingMinutes / 60;
                    g.aiNotUsing += row.aiNotUsingMinutes / 60;
                    g.aiPotential += row.aiPotentialMinutes / 60;
                    break;
                }
            }
        }
    } else {
        // 直近6ヶ月
        for (let i = 5; i >= 0; i--) {
            const m = new Date(now.getFullYear(), now.getMonth() - i, 1);
            const key = `${m.getFullYear()}/${m.getMonth() + 1}`;
            grouped[key] = { aiUsing: 0, aiNotUsing: 0, aiPotential: 0, year: m.getFullYear(), month: m.getMonth() };
        }
        for (const row of data) {
            const d = new Date(row.date);
            if (isNaN(d.getTime())) continue;
            for (const [key, g] of Object.entries(grouped)) {
                if (d.getFullYear() === g.year && d.getMonth() === g.month) {
                    g.aiUsing += row.aiUsingMinutes / 60;
                    g.aiNotUsing += row.aiNotUsingMinutes / 60;
                    g.aiPotential += row.aiPotentialMinutes / 60;
                    break;
                }
            }
        }
    }

    const labels = Object.keys(grouped);
    const aiUsing = labels.map(k => Math.round(grouped[k].aiUsing * 10) / 10);
    const aiNotUsing = labels.map(k => Math.round(grouped[k].aiNotUsing * 10) / 10);
    const aiPotential = labels.map(k => Math.round(grouped[k].aiPotential * 10) / 10);

    return { labels, aiUsing, aiNotUsing, aiPotential };
}

// CalendarSummaryシートからデータを取得
async function fetchCalendarSummaryData() {
    try {
        const { spreadsheetId } = await chrome.storage.local.get(['spreadsheetId']);
        if (!spreadsheetId || !authToken) return [];

        const sheetName = 'CalendarSummary';
        const range = `${sheetName}!A2:M`;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}`;

        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${authToken}` }
        });

        if (!response.ok) return [];

        const data = await response.json();
        if (!data.values) return [];

        return data.values.map(row => ({
            date: row[0],
            user: row[1],
            calendarName: row[2],
            eventCount: parseInt(row[3]) || 0,
            totalMinutes: parseInt(row[4]) || 0,
            aiUsingCount: parseInt(row[5]) || 0,
            aiNotUsingCount: parseInt(row[6]) || 0,
            aiPotentialCount: parseInt(row[7]) || 0,
            avgRate: parseInt(row[8]) || 0,
            aiUsingMinutes: parseInt(row[9]) || 0,
            aiNotUsingMinutes: parseInt(row[10]) || 0,
            aiPotentialMinutes: parseInt(row[11]) || 0
        }));
    } catch (e) {
        console.error('CalendarSummaryデータ取得エラー:', e);
        return [];
    }
}

// カレンダー別ランキングを描画
async function renderCalendarRanking() {
    const container = document.getElementById('calendar-ranking');
    if (!container) return;

    const rawData = await fetchCalendarSummaryData();

    if (rawData.length === 0) {
        container.innerHTML = '<div class="chart-no-data">データがありません</div>';
        return;
    }

    // カレンダー別に集計
    const calMap = {};
    for (const row of rawData) {
        const key = row.calendarName;
        if (!calMap[key]) {
            calMap[key] = { events: 0, aiUsing: 0, totalMinutes: 0, aiUsingMinutes: 0, days: new Set() };
        }
        calMap[key].events += row.eventCount;
        calMap[key].aiUsing += row.aiUsingCount;
        calMap[key].totalMinutes += row.totalMinutes;
        calMap[key].aiUsingMinutes += row.aiUsingMinutes;
        calMap[key].days.add(row.date);
    }

    // AI活用率でソート
    const sorted = Object.entries(calMap)
        .map(([name, d]) => {
            const totalAiRelated = d.events;
            const aiRate = totalAiRelated > 0 ? Math.round(d.aiUsing / totalAiRelated * 100) : 0;
            return { name, events: d.events, aiRate, aiUsingH: Math.round(d.aiUsingMinutes / 60 * 10) / 10, days: d.days.size };
        })
        .sort((a, b) => b.aiRate - a.aiRate);

    let html = `<table class="ranking-table">
        <thead><tr><th></th><th>カレンダー</th><th>予定数</th><th>AI活用率</th></tr></thead>
        <tbody>`;

    sorted.forEach((cal, i) => {
        const rank = i + 1;
        const badgeClass = rank <= 3 ? `rank-${rank}` : 'rank-other';
        html += `<tr>
            <td><span class="rank-badge ${badgeClass}">${rank}</span></td>
            <td>${escapeHtml(cal.name)}</td>
            <td>${cal.events}件<span style="color:var(--text-tertiary);font-size:11px"> / ${cal.days}日</span></td>
            <td><div class="rate-bar">
                <div class="rate-bar-track"><div class="rate-bar-fill" style="width:${cal.aiRate}%"></div></div>
                <span class="rate-bar-value">${cal.aiRate}%</span>
            </div></td>
        </tr>`;
    });

    html += '</tbody></table>';
    container.innerHTML = html;
}

// 初期化実行
init();

function hideDashboard() {
    elements.dashboardModal.classList.add('hidden');
}

// 週次・月次レポート生成
async function generateReport(type) {
    const btn = type === 'weekly' ? elements.generateWeeklyReport : elements.generateMonthlyReport;
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `<span class="loading-spinner"></span> ${type === 'weekly' ? '週次' : '月次'}レポート生成中...`;

    try {
        const { geminiApiKey } = await chrome.storage.local.get(['geminiApiKey']);
        if (!geminiApiKey) {
            showToast('Gemini API Keyが設定されていません。設定画面から設定してください。', 'error');
            return;
        }

        // DailySummaryからデータを取得するか、ダミーデータを使用
        const summaryData = await collectSummaryData(type);

        const analysis = await callGeminiForReport(geminiApiKey, summaryData, type);

        // レポートをスプレッドシートに保存
        if (authToken) {
            try {
                await saveReportToSheet(authToken, type, summaryData, analysis);
                console.log(`${type}レポートをスプレッドシートに保存しました`);
            } catch (e) {
                console.warn('レポートのシート保存に失敗:', e);
            }
        }

        // 結果を表示 (XSS対策済み)
        const html = markdownToSafeHtml(analysis);

        const reportContent = type === 'weekly' ? elements.dashboardReportContentWeekly : elements.dashboardReportContentMonthly;
        const reportSection = type === 'weekly' ? elements.dashboardReportSectionWeekly : elements.dashboardReportSectionMonthly;

        reportContent.innerHTML = html;
        reportSection.style.display = 'block';

        showToast(`${type === 'weekly' ? '週次' : '月次'}レポートを生成しました！`, 'success');

    } catch (error) {
        console.error('レポート生成エラー:', error);
        showToast(`レポート生成エラー: ${error.message}`, 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
    }
}

// メインページ「AI分析レポート」セクション用
async function generateQuickReport(type) {
    const btnMap = { daily: elements.quickReportDaily, weekly: elements.quickReportWeekly, monthly: elements.quickReportMonthly };
    const btn = btnMap[type];
    const originalText = btn.textContent;

    // ボタン無効化＋ローディング表示
    Object.values(btnMap).forEach(b => { b.disabled = true; });
    btn.textContent = '分析中...';
    elements.quickReportResult.classList.remove('hidden');
    elements.quickReportText.innerHTML = '<div class="report-loading"><div class="spinner-small"></div> 分析中...</div>';
    elements.copyReportBtn.classList.add('hidden');

    try {
        const { geminiApiKey } = await chrome.storage.local.get(['geminiApiKey']);
        if (!geminiApiKey) {
            showToast('Gemini API Keyが設定されていません。設定画面から設定してください。', 'error');
            elements.quickReportText.textContent = 'Gemini API Keyが未設定です。';
            return;
        }

        let analysisHtml = '';

        if (type === 'daily') {
            // 日次: 現在表示中のスケジュールデータから分析
            const previewData = generatePreviewData();
            const reportData = {
                userEmail: currentUser?.email || '',
                date: previewData.dateFormatted,
                totalWorkHours: previewData.totalTime,
                scheduleEntries: previewData.entries
            };
            const result = await runAiAnalysis(reportData);
            if (result.success) {
                analysisHtml = markdownToSafeHtml(result.analysis);
            } else {
                analysisHtml = '<p>分析結果を取得できませんでした。</p>';
            }
        } else {
            // 週次/月次: collectSummaryData + callGeminiForReport を再利用
            const summaryData = await collectSummaryData(type);
            const analysis = await callGeminiForReport(geminiApiKey, summaryData, type);
            analysisHtml = markdownToSafeHtml(analysis);
        }

        elements.quickReportText.innerHTML = analysisHtml;
        elements.copyReportBtn.classList.remove('hidden');
        showToast(`${type === 'daily' ? '日次' : type === 'weekly' ? '週次' : '月次'}分析を生成しました`, 'success');

    } catch (error) {
        console.error('クイックレポートエラー:', error);
        elements.quickReportText.textContent = `エラー: ${error.message}`;
        showToast(`分析エラー: ${error.message}`, 'error');
    } finally {
        Object.values(btnMap).forEach(b => { b.disabled = false; });
        btn.textContent = originalText;
    }
}

// サマリーデータ収集（DailySummaryシートから実データを取得）
async function collectSummaryData(type) {
    const now = new Date();
    const daysAgo = type === 'weekly' ? 7 : 30;
    const startDate = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);

    // DailySummaryシートからデータを取得
    const rawData = await fetchDailySummaryData();

    // 期間内のデータをフィルタ
    const periodData = rawData.filter(row => {
        const d = new Date(row.date);
        return !isNaN(d.getTime()) && d >= startDate && d <= now;
    });

    // CalendarSummaryデータも取得
    const calData = await fetchCalendarSummaryData();
    const periodCalData = calData.filter(row => {
        const d = new Date(row.date);
        return !isNaN(d.getTime()) && d >= startDate && d <= now;
    });

    // カレンダー別集計
    let calendarSummary = [];
    if (periodCalData.length > 0) {
        const calMap = {};
        for (const row of periodCalData) {
            const key = row.calendarName;
            if (!calMap[key]) {
                calMap[key] = { events: 0, aiUsing: 0, totalMinutes: 0, aiUsingMinutes: 0 };
            }
            calMap[key].events += row.eventCount;
            calMap[key].aiUsing += row.aiUsingCount;
            calMap[key].totalMinutes += row.totalMinutes;
            calMap[key].aiUsingMinutes += row.aiUsingMinutes;
        }
        calendarSummary = Object.entries(calMap)
            .map(([name, d]) => ({
                name,
                events: d.events,
                aiUsing: d.aiUsing,
                aiRate: d.events > 0 ? Math.round(d.aiUsing / d.events * 100) : 0,
                hours: Math.round(d.totalMinutes / 60)
            }))
            .sort((a, b) => b.aiRate - a.aiRate);
    }

    // DailySummaryにデータがある場合は実データを使用
    if (periodData.length > 0) {
        const totals = periodData.reduce((acc, row) => {
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

        const uniqueDays = new Set(periodData.map(r => r.date)).size;
        const totalEvents = totals.aiUsingCount + totals.aiNotUsingCount + totals.aiPotentialCount;
        const aiRate = totalEvents > 0 ? Math.round(totals.aiUsingCount / totalEvents * 100) : 0;

        // --- AI Impact Log の取得 (オプション) ---
        let impactText = '';
        let aiTimeSaved = 0;
        try {
            const impactData = await fetchAiImpactData('personal'); // TODO: scope
            if (impactData.length > 0) {
                const mvps = [];
                for (const row of impactData) {
                    if (!row.date) continue;
                    const d = new Date(row.date);
                    if (isNaN(d.getTime()) || d < startDate || d > now) continue;

                    const saved = row.timeSaved || 0;
                    aiTimeSaved += saved;

                    const score = row.impactScore || 0;
                    if (score >= 4) {
                        mvps.push({ task: row.task, score, action: row.valueCreation, time: saved });
                    }
                }
                if (mvps.length > 0) {
                    impactText = '\n## 高いインパクトを生んだ事例 (Impact MVP)\n' + mvps.map(m =>
                        `  - タスク: ${m.task} (スコア: Lv${m.score}, 削減時間: ${m.time}分)\n    価値創造アクション: ${m.action}`
                    ).join('\n');
                }
            }
        } catch (e) {
            console.warn('AI Impact Log取得エラー(レポート用 - 実データ枠):', e);
        }

        totals.aiTimeSaved = aiTimeSaved;

        return {
            period: { start: formatDate(startDate), end: formatDate(now), days: uniqueDays },
            totals,
            averages: {
                dailyEvents: Math.round(totals.events / uniqueDays),
                dailyMinutes: Math.round(totals.totalMinutes / uniqueDays),
                aiRate
            },
            calendarSummary,
            impactText
        };
    }

    // フォールバック: DailySummaryが空の場合は当日データから推定
    let totalEvents = currentEvents.length;
    let totalMinutes = 0;
    let aiUsingCount = 0;
    let aiNotUsingCount = 0;
    let aiPotentialCount = 0;

    for (const event of currentEvents) {
        totalMinutes += getEventDuration(event);
        const estimation = estimateAIUsage(event.summary || '');
        if (estimation.flag === 'yes-using') aiUsingCount++;
        else if (estimation.flag === 'yes-potential') aiPotentialCount++;
        else aiNotUsingCount++;
    }

    // --- AI Impact Log の取得 (オプション) ---
    let impactText = '';
    let aiTimeSaved = 0;
    try {
        const impactData = await fetchAiImpactData('personal');
        if (impactData.length > 0) {
            const mvps = [];
            for (const row of impactData) {
                if (!row.date) continue;
                const d = new Date(row.date);
                if (isNaN(d.getTime()) || d < startDate || d > now) continue;

                const saved = row.timeSaved || 0;
                aiTimeSaved += saved;

                const score = row.impactScore || 0;
                if (score >= 4) {
                    mvps.push({ task: row.task, score, action: row.valueCreation, time: saved });
                }
            }
            if (mvps.length > 0) {
                impactText = '\n## 高いインパクトを生んだ事例 (Impact MVP)\n' + mvps.map(m =>
                    `  - タスク: ${sanitizeForPrompt(m.task)} (スコア: Lv${m.score}, 削減時間: ${m.time}分)\n    価値創造アクション: ${sanitizeForPrompt(m.action)}`
                ).join('\n');
            }
        }
    } catch (e) {
        console.warn('AI Impact Log取得エラー(レポート用):', e);
    }

    return {
        period: { start: formatDate(startDate), end: formatDate(now), days: daysAgo },
        totals: {
            events: totalEvents * daysAgo,
            totalMinutes: totalMinutes * daysAgo,
            aiUsingCount: aiUsingCount * daysAgo,
            aiNotUsingCount: aiNotUsingCount * daysAgo,
            aiPotentialCount: aiPotentialCount * daysAgo,
            aiUsingMinutes: Math.round(totalMinutes * 0.3 * daysAgo),
            aiNotUsingMinutes: Math.round(totalMinutes * 0.5 * daysAgo),
            aiPotentialMinutes: Math.round(totalMinutes * 0.2 * daysAgo),
            aiTimeSaved: aiTimeSaved // 推定は0にして実データがある場合のみ上書き
        },
        averages: {
            dailyEvents: totalEvents,
            dailyMinutes: totalMinutes,
            aiRate: totalEvents > 0 ? Math.round((aiUsingCount / totalEvents) * 100) : 0
        },
        calendarSummary,
        impactText
    };
}

// Gemini API呼び出し（レポート用）
async function callGeminiForReport(apiKey, summaryData, type) {
    const periodLabel = type === 'weekly' ? '週次' : '月次';

    // カレンダー別ランキングテキスト生成
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
        const error = await response.json();
        throw new Error(error.error?.message || 'Gemini API呼び出し失敗');
    }

    const data = await response.json();
    let text = data.candidates?.[0]?.content?.parts?.[0]?.text || '分析結果を取得できませんでした';

    // トークン制限に到達した場合の警告
    const finishReason = data.candidates?.[0]?.finishReason;
    if (finishReason === 'MAX_TOKENS') {
        text += '\n\n⚠️ レポートがトークン上限に達したため、一部省略されている可能性があります。';
    }

    return text;
}

// AiImpactLogシートからデータを取得
async function fetchAiImpactData(scope = 'personal') {
    try {
        const { spreadsheetId, directorySpreadsheetId } = await chrome.storage.local.get(['spreadsheetId', 'directorySpreadsheetId']);
        if (!spreadsheetId || !authToken) return [];

        const targetConfig = {};

        if (scope === 'personal') {
            targetConfig[spreadsheetId] = [currentUser.email.toLowerCase()];
        } else if (scope.startsWith('group-')) {
            const groupId = scope.replace('group-', '');
            const group = dashboardGroups.find(g => g.id === groupId);

            if (group) {
                const members = group.members.map(m => m.trim().toLowerCase());
                let directoryMap = {};
                if (directorySpreadsheetId) {
                    directoryMap = await fetchDirectoryData(directorySpreadsheetId);
                }

                members.forEach(email => {
                    let targetSheetId;
                    if (email === currentUser.email.toLowerCase()) {
                        targetSheetId = spreadsheetId;
                    } else {
                        targetSheetId = directoryMap[email] || (directorySpreadsheetId ? null : spreadsheetId);
                    }

                    if (targetSheetId) {
                        if (!targetConfig[targetSheetId]) {
                            targetConfig[targetSheetId] = [];
                        }
                        targetConfig[targetSheetId].push(email);
                    }
                });
            }
        }

        const fetchPromises = Object.entries(targetConfig).map(async ([sId, targetEmails]) => {
            if (!sId) return [];

            const sheetName = 'AiImpactLog';
            const range = `${sheetName}!A:H`;
            const encodedRange = encodeURIComponent(range);
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${sId}/values/${encodedRange}`;

            try {
                const response = await fetch(url, {
                    headers: { 'Authorization': `Bearer ${authToken}` }
                });

                if (!response.ok) return [];

                const data = await response.json();
                const values = data.values || [];
                if (values.length <= 1) return [];

                return values.slice(1).filter(row => {
                    const rowUser = (row[1] || '').trim().toLowerCase();
                    return targetEmails.includes(rowUser);
                }).map(row => ({
                    date: row[0] || '',
                    user: row[1] || '',
                    task: row[2] || '',
                    tool: row[3] || '',
                    oneliner: row[4] || '',
                    impactScore: parseInt(row[5]) || null,
                    timeSaved: parseInt(row[6]) || 0,
                    valueCreation: row[7] || ''
                }));

            } catch (e) {
                console.warn(`シート(${sId})のAiImpactLog取得エラー:`, e);
                return [];
            }
        });

        const results = await Promise.all(fetchPromises);
        return results.flat();

    } catch (e) {
        console.error('AiImpactLogデータ取得エラー:', e);
        return [];
    }
}
