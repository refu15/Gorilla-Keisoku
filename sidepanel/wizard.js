/**
 * 日報作成ウィザード (wizard.js)
 * 5ステップの対話式日報作成フロー
 * 機能: 気づき学習・業務別AI提案・未完了タスク繰り越し・カレンダー登録
 */

import { getEventsForDate } from '../src/api/calendar.js';
import { saveWizardReportToSheet, saveAiImpactLog, saveDecisionLog } from '../src/api/sheets.js';
import { escapeHtml, sanitizeForPrompt } from '../src/utils/sanitize.js';
import { saveWizardToNotion, getDestinationConfig } from '../src/api/notion.js';

/** Gemini proposals[] の安全検証 */
function validateProposals(parsed) {
    if (!Array.isArray(parsed)) return null;
    const valid = parsed
        .filter(p => p && typeof p.label === 'string' && typeof p.text === 'string')
        .map(p => ({ label: p.label.slice(0, 50), text: p.text.slice(0, 1000) }));
    return valid.length >= 1 ? valid : null;
}

/** Gemini API 直接呼び出し */
async function callGeminiRaw(prompt) {
    const { geminiApiKey } = await chrome.storage.local.get(['geminiApiKey']);
    if (!geminiApiKey) throw new Error('Gemini API Keyが設定されていません');
    const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiApiKey}`,
        {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { maxOutputTokens: 2048, temperature: 0.7 } })
        }
    );
    if (!response.ok) {
        let errDesc = 'Gemini APIエラー';
        try {
            const errJson = await response.json();
            errDesc += ': ' + (errJson.error?.message || response.statusText);
        } catch (e) { }
        throw new Error(errDesc);
    }
    const data = await response.json();
    return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

// ===== 定数 =====
const STEP_LABELS = ['行動内容', 'AI活用', '判断の記録', '気づき', '明日挑戦・改善', 'その他', '翌日予定', '日報確認', 'カレンダー登録'];

// ===== 判断タグ =====
const DECISION_TAGS = [
    { emoji: '⏱️', label: '時間優先' },
    { emoji: '💰', label: 'コスト優先' },
    { emoji: '🎯', label: '品質優先' },
    { emoji: '🛡️', label: 'リスク回避' },
    { emoji: '🔄', label: 'ツール変更' },
    { emoji: '✋', label: '手動選択' },
    { emoji: '🧪', label: '新手法テスト' },
    { emoji: '💡', label: '閃き' },
];
const EXCLUDE_KEYWORDS = ['昼休憩', '昼食', 'ランチ', '移動', '準備', '休憩', 'break', 'lunch'];
const AI_TOOLS_DEFAULT = ['Antigravity', 'Claude code', 'Claude code + MCP', 'ChatGPT', 'Gemini'];
let AI_TOOLS = [...AI_TOOLS_DEFAULT];

const AI_ONELINER_SUGGESTIONS = [
    'タグ補完、クラス自動挿入',
    'Jsの作成、タグの補完、クラス自動挿入、Sassの自動生成、PHP',
    'コード補完、自動生成',
    'エージェント構築・調整',
    'draw.io連携で図式化',
    '反映確認補助',
    'ドキュメント作成・整理',
    'レビュー・フィードバック補助',
];

// ===== 業務別 AI 活用提案マップ =====
const BUSINESS_AI_MAP = [
    {
        keywords: ['SNS', 'インスタ', 'Instagram', 'X ', 'Twitter', 'TikTok', 'Facebook'], category: 'マーケ / SNS代行', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['SNS投稿文の作成・改善', 'ハッシュタグ提案・最適化', 'キャプション生成（商品・サービス紹介）', 'エンゲージメント向上のコメント返信案作成', '投稿スケジュール・カレンダー作成補助', 'トレンドキーワードのリサーチと提案', 'ターゲット別トンマナ調整', 'リール・ショート動画の構成案作成', 'ストーリーズ用コピー・CTA文作成', '競合アカウントの分析サポート', '月次レポート用コメント・考察文の下書き', 'キャンペーン告知文の作成', 'インフルエンサー向けDM文面の作成', 'ユーザーの声（UGC）まとめ文作成', 'ペルソナ設定・ターゲット整理補助', 'A/Bテスト用バリエーション文の生成', 'プロフィール文の改善提案']
    },
    {
        keywords: ['WEB', 'HP', 'LP', 'サイト', 'ホームページ', 'ランディング', 'コーディング', 'HTML', 'CSS', 'WordPress'], category: 'マーケ / WEB制作', tools: ['Antigravity', 'Claude code'],
        oneliners: ['HTML/CSS/JSの自動生成・補完', 'クラス名・変数名の命名提案', 'レスポンシブ対応のコード補助', 'LP構成・セクション構造の提案', 'コピーライティング（キャッチコピー・ボディ文）', 'WordPressカスタマイズコードの生成', 'アニメーション・インタラクションの実装補助', 'SEO対策：メタタグ・alt属性・構造化データ作成', 'ページ速度改善のリファクタリング提案', 'フォームバリデーションのコード生成', 'エラー修正・デバッグの補助', 'Sassの自動生成・クラス設計', 'PHPテンプレートの実装補助', 'アクセシビリティ改善の提案', 'コンテンツ構成（見出し・導線設計）の提案', 'OGP設定・SNSシェア最適化補助', '更新用テキスト・バナーコピーの作成']
    },
    {
        keywords: ['広告', 'Meta', 'Google広告', 'リスティング', 'バナー', '媒体'], category: 'マーケ / 広告作成', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['広告見出し・説明文の生成（複数バリエーション）', 'A/Bテスト用コピーの作成', 'ターゲット訴求軸の整理・提案', 'CVR改善のランディングページコピー提案', 'Meta広告用プライマリテキスト・ヘッドラインの作成', 'Google広告用レスポンシブ広告の文言生成', 'キーワードリストの拡張・整理補助', '除外キーワードの提案', '広告レポートの考察文・改善提案文作成', 'ペルソナごとのメッセージ訴求切り口の提案', 'バナー用コピー・CTAボタン文の作成', '競合広告分析のサマリー作成', 'クリエイティブブリーフの草案作成', 'リターゲティング広告向けメッセージの作成', 'コンバージョン最適化の施策案提案']
    },
    {
        keywords: ['LINE', 'LINEリッチ', 'LINEメッセージ', 'LINE配信'], category: 'マーケ / LINE運用代行', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['LINE配信文の作成（告知・キャンペーン）', 'ステップ配信シナリオの設計・文面作成', 'チャットボット応答文の作成・改善', 'リッチメニューのコピー・案内文作成', '友だち追加促進メッセージの作成', 'クーポン・特典案内文の生成', '定期配信コンテンツの企画・文案作成', 'セグメント配信用メッセージの分岐設計補助', 'クロスセル・アップセル訴求文の作成', '開封率向上のための件名・冒頭文改善提案', '購買後フォローアップメッセージの作成', '問い合わせ対応テンプレートの作成', 'キャラクター・トーン設定の整理', 'イベント・セミナー告知文の作成', 'アンケート・リサーチメッセージの作成']
    },
    {
        keywords: ['動画', '編集', 'YouTube', 'ショート', 'リール', '撮影', 'サムネ'], category: 'マーケ / 動画制作', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['動画台本・ナレーション原稿の作成', '構成案（オープニング〜CTA）の設計', 'YouTube動画タイトル・説明文の生成', 'サムネイルのキャッチコピー提案', 'ショート動画の切り口・ネタ提案', '字幕・テロップ用テキストの生成', 'BGM・効果音の使用場面提案', '編集指示書・カット割りの整理補助', '視聴者コメントへの返信文案作成', 'シリーズ企画・テーマ設定の提案', '動画SEO（タグ・章立て）の最適化補助', '企業紹介・サービス紹介動画の構成提案', 'インタビュー質問リストの作成', '撮影チェックリストの作成', 'チャンネル概要・プロフィール文の改善', 'エンドカード・概要欄CTAの文言作成']
    },
    {
        keywords: ['AI導入', 'AI支援', 'AI活用', 'AIリスキリング', 'リスキリング', '研修', '講座'], category: 'AX / AI導入・リスキリング', tools: ['Antigravity', 'Claude code + MCP'],
        oneliners: ['AI活用フロー・業務プロセスの設計補助', '研修資料・スライドのドラフト作成', 'ユースケース整理・優先順位付け', 'エージェント構築・プロンプト調整', '業務別AIツール比較表の作成', '導入効果測定の指標（KPI）設計補助', 'ワークショップ用演習問題の作成', '担当者向けAI活用マニュアル作成', 'ROI試算・コスト削減試算のサポート', '社内FAQ・Q&Aコンテンツの作成', '活用事例レポートの草案作成', 'プロンプト集・テンプレート整備の補助', 'ベンダー比較・RFP要件整理の補助', 'チェンジマネジメント用コミュニケーション文書作成', 'モニタリング・改善サイクルの設計補助', 'ヒアリング質問リストの自動生成', '各部署向けカスタマイズ提案資料の作成']
    },
    {
        keywords: ['プロンプト', 'prompt'], category: 'AX / プロンプト作成', tools: ['Antigravity', 'Claude code'],
        oneliners: ['プロンプト設計・最適化', 'Few-shot・Chain-of-Thought例の作成', 'チェーンプロンプト（多段階処理）の設計', 'System / User / Assistant役割設定の整理', 'タスク別プロンプトテンプレートの作成', '出力フォーマット指定の最適化', '日本語・英語プロンプトの相互変換', 'ハルシネーション防止のガードレール設計', 'プロンプト評価基準・採点ルーブリック作成', 'ユーザー向けプロンプト入力ガイドの作成', 'ロールプレイ・ペルソナ設定の設計', 'RAG連携を意識したプロンプト設計', '社内業務別プロンプト集の整備', '制約条件・禁止事項の明文化補助', 'プロンプトA/Bテストの結果分析補助']
    },
    {
        keywords: ['アプリ', '開発', 'AI開発', 'システム', 'API', 'バックエンド', 'フロントエンド'], category: 'AX / アプリ・AI開発', tools: ['Claude code', 'Claude code + MCP', 'Antigravity'],
        oneliners: ['コード補完・自動生成（関数・クラス）', 'デバッグ・エラー原因の特定と修正', 'API設計・エンドポイント定義補助', 'データベーススキーマ設計の提案', 'テストコード（ユニット・E2E）の自動生成', 'リファクタリング・コード品質改善', 'ドキュメント・コメント自動生成', 'セキュリティレビュー・脆弱性チェック補助', 'AIモデル選定・パラメータ調整の補助', 'プロンプトエンジニアリング組み込みの設計', 'CI/CDパイプライン設定の補助', 'パフォーマンス最適化の提案', 'ライブラリ・フレームワーク選定補助', '要件定義・仕様書のドラフト作成', 'コードレビュー・プルリクエスト文の作成', 'インフラ構成の設計補助', 'ログ分析・エラートレースの解釈']
    },
    // ▼ 今回追加した汎用・部門別業務用の定型文リスト（各20個）
    {
        keywords: ['MTG', 'ミーティング', '商談', 'クライアント', '顧客', '打合せ', 'キックオフ'], category: 'クライアントMTG', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['議事録の構造化', '要約とネクストアクションの抽出', '商談メモの整理', '顧客の課題の壁打ち', '提案の切り返しトーク案作成', '専門用語のかみ砕き', '競合他社との比較情報出し', '顧客インタビューの文字起こし要約', 'ヒアリング項目案の作成', '顧客からの質問への回答案作成', 'MTG前のアジェンダ作成補助', '業界動向のクイックリサーチ', 'リスク確認項目の洗い出し', '導入事例のピックアップ', '顧客の要望に合わせたプラン提案案', '失注理由の分析補助', '次回提案の構成案作成', 'クレーム対応の文面推敲', '契約書・NDAの要点確認補助', 'MTG後の御礼メール草案']
    },
    {
        keywords: ['社内MTG', '部会', '定例', 'ブレスト', '1on1', '朝会', '終礼', '会議'], category: '社内MTG', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['決定事項と宿題の整理', '議事要旨の作成', 'ブレストのアイデア出しと壁打ち', '曖昧なアイデアの言語化', '課題解決のフレームワーク当てはめ', 'チーム内共有事項の箇条書き化', '報告用サマリーの作成', 'KPT（振り返り）の整理', 'プロジェクト進捗の可視化案', 'MTGのタイムキープ案作成', '意見の対立点の整理', 'プロコン（賛否）表の作成', 'ファシリテーションのポイント出し', '新規企画のラフ案出し', 'メンバーへのフィードバック文面案', '目標設定（OKR等）のブラッシュアップ', '社内報・共有ドキュメントの推敲', 'MTG前の資料読み込みと要約', '議題の優先順位付け', 'アクションアイテムの担当者割り振り案']
    },
    {
        keywords: ['営業', '資料作成', '提案', 'プレゼン', '企画書', 'スライド'], category: '営業・提案資料作成', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['スライドの構成案出し', 'キャッチコピー・見出しの提案', '導入文・背景の文章作成', 'メリット・デメリットの整理', '構成の論理破綻チェック', '他社との差別化ポイントの言語化', 'ターゲットに合わせたトーン変更', 'ペルソナ設定の壁打ち', '統計データや根拠の構成補助', '料金プランの見せ方アイデア', 'FAQ（よくある質問）の想定作成', 'カスタマージャーニーの作成補助', '課題から解決策へのストーリー構築', '簡潔な箇条書きへのリライト', '図解やグラフの構成アイデア', '事例紹介のエピソード化', 'クロージングのメッセージ作成', '過去資料の再構成と統合', '相手の反対例（反論）の想定', 'プレゼン用のトークスクリプト作成']
    },
    {
        keywords: ['プログラミング', '実装', 'コーディング', 'コード', 'バグ'], category: '開発・コーディング', tools: ['Claude code', 'Antigravity'],
        oneliners: ['コードの解説とコメント追加', 'バグの特定と修正案の提案', 'リファクタリング案の提示', '単体テストコードの自動生成', '正規表現の作成・解説', 'SQLクエリの作成と最適化', 'エラーメッセージの原因調査', '新規機能のロジック実装補助', 'API仕様書のベース作成', 'クラス名・変数名のネーミング案', '処理の高速化・最適化の相談', '非同期処理・Promiseの書き方確認', '状態管理（State）の設計壁打ち', 'UIコンポーネントの構造設計', 'Gitコマンドの確認', '環境構築のトラブルシューティング', 'ライブラリの選定と比較', 'TypeScriptの型定義の支援', '複雑なアルゴリズムの擬似コード化', 'セキュリティの脆弱性チェック補助']
    },
    {
        keywords: ['設計', '要件定義', '仕様', 'アーキテクチャ', 'DB設計'], category: '設計・要件定義', tools: ['Claude code', 'ChatGPT', 'Gemini'],
        oneliners: ['要件定義書のドラフト作成', '機能一覧・画面一覧の洗い出し', 'ユーザーストーリーの作成', 'DBスキーマ・テーブル設計の壁打ち', 'APIエンドポイントの設計補助', 'アーキテクチャ図の構成案', '非機能要件のチェックリスト作成', 'エッジケースや例外処理の想定', 'フローチャートのステップ言語化', 'エラーハンドリングの設計', 'インフラ・サーバー構成の相談', '仕様の矛盾点・抜け漏れチェック', '既存システムからの移行手順案', '認証・認可周りの設計確認', 'データモデリングのリレーション確認', 'ワイヤーフレームの要素洗い出し', '外部サービス連携の技術検証', '開発フェーズ・マイルストーンの策定', 'テスト仕様書・シナリオの項目出し', 'ドメイン駆動設計のモデル抽出補助']
    },
    {
        keywords: ['リサーチ', '調査', 'リファレンス', '検索', '競合調査', '市場調査', '文献'], category: 'リサーチ・調査', tools: ['ChatGPT', 'Gemini', 'Perplexity'],
        oneliners: ['指定テーマの基礎知識まとめ', '特定のキーワードの網羅的洗い出し', '長文記事・論文の要約と翻訳', '競合他社のサービス比較表作成', 'トレンドや市場規模の推察', '複数の情報の比較と共通点抽出', '専門用語・概念のわかりやすい図解案', '法令やガイドラインの要点整理', 'アンケート結果の自由回答の分析・分類', 'PEIT/3Cなどのフレームワーク分析', '検索キーワードの拡張・類語探し', '調査結果のサマリーレポート化', '特定技術のトレンドと将来性調査', '国内外のベストプラクティス調査', '該当記事の信憑性チェックの観点出し', '歴史的背景や時系列の整理', '関連するニュースや事例のピックアップ', '特定業界の課題感の言語化', 'インタビュー対象者のリストアップ基準出し', '仮説構築のための壁打ち相手']
    },
    {
        keywords: ['マーケ', 'マーケティング', 'SEO', 'プロモーション'], category: 'SNS・マーケティング', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['投稿用のキャプション作成・推敲', 'ハッシュタグの選定・提案', 'ペルソナに刺さる切り口出し', 'ターゲットに応じたトンマナの変更', '投稿のタイトル・サムネ案出し', 'A/Bテスト用のバリエーション作成', '競合アカウントの傾向分析補助', 'ユーザーのコメント返信案作成', 'ストーリーズの企画・構成案', 'リール・ショート動画の台本作成', 'キャンペーン告知文のドラフト', 'プロフィール文の魅力的なリライト', 'UGC（ユーザー投稿）を促す企画案', '今月の投稿テーマのアイデア出し', '広告用キャッチコピー・見出しの100本ノック', 'メールマガジンの件名と本文作成', 'LINE配信のステップメッセージ作成', 'LPの構成とコピー作成', 'SEOを意識した記事の構成案', 'インフルエンサーへのアサインDM作成']
    },
    {
        keywords: ['メール', 'チャット', 'Slack', 'Chatwork', '連絡', '返信', '案内'], category: 'メール・チャット対応', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['丁寧なお詫びメールの推敲', '複雑な要件を伝えるメールの整理', 'ニュアンスが伝わりにくい文章の校正', '取引先への営業・提案メール文面', 'アポイントメント打診のメール', 'リスケ・日程調整のメール文面', 'お断り・辞退の角が立たない文章', '上司への報告・相談チャットの推敲', 'クレームに対する初期対応文面', '英語など外国語メールの作成・翻訳', '長いメールの要点3行まとめ', '相手の意図を汲み取るための壁打ち', '社外向けのお知らせ・プレスリリース案', '滞っている返信の催促文面', 'マニュアル化するためのテンプレート作成', '感謝や労いを伝える温かい文章', '専門外の人への噛み砕いた説明文', '季節の挨拶やフォーマルな言い回しの提案', 'トラブル発生時の緊急報告フォーマット', 'チャットツールの全体アナウンス文作成']
    },
    {
        keywords: ['経営', '戦略', '事業計画', '採用', '組織', 'マネジメント', '役員会'], category: '経営メンバー', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['事業計画の策定壁打ち', 'KPIツリーの作成', 'OKRの設定案', '競合優位性の言語化', 'ピッチデッキの構成案', '投資家向け想定問答', '新規事業のペルソナ設定', 'リスクシナリオの洗い出し', '組織課題の分析フレームワーク', 'ビジョン・ミッションの言語化', '経営合宿のアジェンダ作成', '採用要件の定義補助', '評価制度の見直し案', '全社総会用スピーチの推敲', '月次報告フォーマットの改善', '社内向けメッセージのトーン調整', '撤退基準の壁打ち', 'アライアンス候補の条件整理', 'コスト削減のアイデア出し', 'M&A基本方針の整理']
    },
    {
        keywords: ['経理', '会計', '決算', '精算', '請求', '財務'], category: '管理部（経理）', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['経費精算の仕訳確認', '勘定科目の分類相談', '決算業務のタスクチェックリスト', 'キャッシュフロー予測の補助', '節税対策の基礎調査', '財務諸表分析のサマリー作成', '予実管理フォーマット案', 'インボイス制度の要点確認', '監査法人の指摘事項の整理', '支払いスケジュールの最適化', '固定資産の減価償却確認', '経費ルール改定の案内文', '資金繰り表の作成補助', '立て替え精算の遅延督促文', '助成金・補助金情報の要約', '取引先の与信管理チェック', '稟議書の財務的レビュー', '月次決算報告資料の構成', '税理士への質問事項整理', '会計ソフト移行のタスク整理']
    },
    {
        keywords: ['法務', '契約', 'NDA', '規約', 'コンプライアンス', '法律'], category: '管理部（法務）', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['契約書の条項レビュー', 'NDAの一般的な記載漏れ確認', '業務委託契約のリスク洗い出し', 'プライバシーポリシーの改定文案', '利用規約の新旧対照表作成補助', '知財・商標の基礎調査', 'インシデント発生時の報告書ドラフト', 'コンプライアンス研修の資料作成', '下請法の適用可否チェックリスト', '個人情報保護法の要点整理', '反社チェックの基準設定案', '電子契約導入の社内案内文', '規程類の書式統一・校正', 'クレームの法務的見解の整理', '顧問弁護士への相談用メモ作成', '取締役会・株主総会アジェンダ案', '新規サービスの適法性リサーチ', 'ライセンス契約の要点確認', '労働基準法の概要調査', '解約・退会トラブルの対応フォーマット']
    },
    {
        keywords: ['雑務', '総務', '備品', '来客', 'イベント', '事務'], category: '管理部（雑務・総務）', tools: ['ChatGPT', 'Gemini'],
        oneliners: ['オフィス備品の発注リスト整理', '郵便物・宅急便の対応マニュアル', '来客対応・お茶出しのフロー作成', '社内イベントの企画・進行表作成', '歓送迎会のお礼メール・案内文', '年賀状・暑中見舞いの文面作成', 'オフィスのレイアウト変更案', '防災備蓄品のリストアップ', 'ファシリティ（設備）のトラブル報告文', '共有スペースの利用ルール作成', 'ゴミ出し・清掃ルールの案内', '会議室予約の調整メール', '名刺発注のタスク整理', '書類ファイリングのルール案', '社員旅行のしおり作成', '健康診断の案内と受診推奨文', '各種パスワード管理表のフォーマット', '社内アンケートの作成集計補助', 'インフルエンザ予防接種の通知', '来訪者向けアクセスマップのテキスト説明']
    }
];

function suggestForTask(title) {
    if (!title) return null;

    // カスタム設定(業種プロファイル)を優先してマッチング
    for (const profile of workProfileMapping) {
        if (title.includes(profile.category)) {
            // 見つかった場合はカテゴリと一言だけ返す（ツールはデフォルト）
            return { category: profile.category, tools: AI_TOOLS_DEFAULT, oneliners: profile.oneliners };
        }
    }

    // デフォルトのビジネスマップ
    for (const entry of BUSINESS_AI_MAP) {
        if (entry.keywords.some(kw => title.includes(kw))) return { category: entry.category, tools: entry.tools, oneliners: entry.oneliners };
    }
    return null;
}

// ===== 業務体系プロファイル =====
let workProfileMapping = []; // [{ category: '...', oneliners: ['...'] }]

async function loadWorkProfile() {
    const { wizardWorkProfile } = await chrome.storage.local.get(['wizardWorkProfile']);
    const lines = (wizardWorkProfile || '').split('\n').map(l => l.trim()).filter(Boolean);
    workProfileMapping = lines.map(line => {
        const parts = line.split(/[:：]/);
        if (parts.length >= 2) {
            const category = parts[0].trim();
            const oneliners = parts.slice(1).join(':').split(/[,、]/).map(s => s.trim()).filter(Boolean);
            return { category, oneliners };
        }
        return { category: line, oneliners: [] };
    });
}

// ===== ウィザード内からのプロファイル追加 =====
async function saveNewProfileFromWizard(categoryTitle, newOneliner) {
    if (!categoryTitle || !newOneliner) return false;

    // 現在の最新をロード
    await loadWorkProfile();

    const trimmedOneliner = newOneliner.trim().slice(0, 100);
    if (!trimmedOneliner) return false;

    let profile = workProfileMapping.find(p => p.category === categoryTitle);
    if (profile) {
        if (!profile.oneliners.includes(trimmedOneliner)) {
            profile.oneliners.push(trimmedOneliner);
        }
    } else {
        workProfileMapping.push({ category: categoryTitle, oneliners: [trimmedOneliner] });
    }

    // 文字列フォーマットに戻して保存
    const wizardWorkProfile = workProfileMapping
        .map(p => `${p.category}: ${p.oneliners.join(', ')}`)
        .join('\n');
    await chrome.storage.local.set({ wizardWorkProfile });
    return true;
}

// ===== 状態管理 =====
let wizardState = {};
let currentStep = 1;
let authTokenRef = null;
let el = {};

// ===== 初期化 =====
export function initWizard(token) {
    authTokenRef = token;
    el = {
        modal: document.getElementById('wizard-modal'),
        title: document.getElementById('wizard-title'),
        stepLabel: document.getElementById('wizard-step-label'),
        content: document.getElementById('wizard-step-content'),
        loading: document.getElementById('wizard-loading'),
        loadingMsg: document.getElementById('wizard-loading-msg'),
        dots: document.querySelectorAll('.wizard-dot'),
        backBtn: document.getElementById('wizard-back-btn'),
        okBtn: document.getElementById('wizard-ok-btn'),
        closeBtn: document.getElementById('wizard-close-btn'),
    };
    el.backBtn.addEventListener('click', handleBack);
    el.okBtn.addEventListener('click', handleOk);
    el.closeBtn.addEventListener('click', closeWizard);
    loadCustomTools();
    loadCustomOneliners();
    loadWorkProfile();
}

export function updateWizardToken(token) { authTokenRef = token; }

// ===== カスタムツール＆一言 永続化 =====
let CUSTOM_ONELINERS = [];
let DELETED_ONELINERS = [];
let DELETED_TOOLS = [];

async function loadCustomTools() {
    const { wizardCustomTools, wizardDeletedTools } = await chrome.storage.local.get(['wizardCustomTools', 'wizardDeletedTools']);
    if (Array.isArray(wizardDeletedTools)) {
        DELETED_TOOLS = wizardDeletedTools.filter(t => typeof t === 'string' && t.trim());
    }
    // 不要なツールを除外し再構成
    AI_TOOLS = AI_TOOLS_DEFAULT.filter(t => !DELETED_TOOLS.includes(t));
    if (Array.isArray(wizardCustomTools)) {
        wizardCustomTools.forEach(t => { if (typeof t === 'string' && t.trim() && !AI_TOOLS.includes(t) && !DELETED_TOOLS.includes(t)) AI_TOOLS.push(t.trim()); });
    }
}

async function saveCustomTool(name) {
    const trimmed = name.trim().slice(0, 50);
    if (!trimmed || AI_TOOLS.includes(trimmed)) return false;
    AI_TOOLS.push(trimmed);
    const { wizardCustomTools = [] } = await chrome.storage.local.get(['wizardCustomTools']);
    if (!wizardCustomTools.includes(trimmed)) { wizardCustomTools.push(trimmed); await chrome.storage.local.set({ wizardCustomTools }); }
    return true;
}

async function deleteCustomTool(name) {
    AI_TOOLS = AI_TOOLS.filter(t => t !== name);
    const { wizardCustomTools = [] } = await chrome.storage.local.get(['wizardCustomTools']);
    if (wizardCustomTools.includes(name)) {
        const updated = wizardCustomTools.filter(t => t !== name);
        await chrome.storage.local.set({ wizardCustomTools: updated });
    } else {
        if (!DELETED_TOOLS.includes(name)) {
            DELETED_TOOLS.push(name);
            await chrome.storage.local.set({ wizardDeletedTools: DELETED_TOOLS });
        }
    }
    return true;
}

async function loadCustomOneliners() {
    const { wizardCustomOneliners, wizardDeletedOneliners } = await chrome.storage.local.get(['wizardCustomOneliners', 'wizardDeletedOneliners']);
    if (Array.isArray(wizardCustomOneliners)) {
        CUSTOM_ONELINERS = wizardCustomOneliners.filter(t => typeof t === 'string' && t.trim());
    }
    if (Array.isArray(wizardDeletedOneliners)) {
        DELETED_ONELINERS = wizardDeletedOneliners.filter(t => typeof t === 'string' && t.trim());
    }
}

async function saveCustomOneliner(name) {
    const trimmed = name.trim().slice(0, 100);
    if (!trimmed || CUSTOM_ONELINERS.includes(trimmed)) return false;
    CUSTOM_ONELINERS.push(trimmed);
    await chrome.storage.local.set({ wizardCustomOneliners: CUSTOM_ONELINERS });
    return true;
}

async function deleteCustomOneliner(name) {
    if (CUSTOM_ONELINERS.includes(name)) {
        CUSTOM_ONELINERS = CUSTOM_ONELINERS.filter(t => t !== name);
        await chrome.storage.local.set({ wizardCustomOneliners: CUSTOM_ONELINERS });
    } else {
        if (!DELETED_ONELINERS.includes(name)) {
            DELETED_ONELINERS.push(name);
            await chrome.storage.local.set({ wizardDeletedOneliners: DELETED_ONELINERS });
        }
    }
    return true;
}

// ===== 気づき学習：過去の文章を保存・ロード =====
async function saveInsightHistory(purpose, result) {
    const { wizardInsightHistory = [] } = await chrome.storage.local.get(['wizardInsightHistory']);
    wizardInsightHistory.push({
        purpose: (purpose || '').slice(0, 300),
        result: (result || '').slice(0, 300),
        date: new Date().toISOString().slice(0, 10)
    });
    // 直近20件のみ保持
    if (wizardInsightHistory.length > 20) wizardInsightHistory.splice(0, wizardInsightHistory.length - 20);
    await chrome.storage.local.set({ wizardInsightHistory });
}

async function loadInsightHistory() {
    const { wizardInsightHistory = [] } = await chrome.storage.local.get(['wizardInsightHistory']);
    return wizardInsightHistory;
}

// ===== ウィザード開始 =====
export async function startWizard(passedEvents, passedScheduleData) {
    wizardState = {
        sessionGeneratedOneliners: [],
        rawEvents: passedEvents, scheduleData: passedScheduleData,
        events: [], additionalTasks: '', pendingTasks: [],
        actionItems: [], aiUsage: [], aiCurrentIdx: 0,
        decisions: [],
        insightPurpose: '', insightResult: '', valueCreation: '',
        tryAndImprove: '', otherNotes: '',
        tomorrowEvents: [], tomorrowAdditional: '',
        _tomorrowTaskItems: [],
        schedule: [], finalReport: '',
    };
    currentStep = 1;
    el.modal.classList.remove('hidden');
    await loadWorkProfile();
    await showStep(1);
}

function closeWizard() { el.modal.classList.add('hidden'); }

// ===== ステップ制御 =====
async function showStep(n) {
    currentStep = n;
    updateIndicator(n);
    updateStepLabel(n);
    el.backBtn.disabled = (n <= 1);
    el.okBtn.textContent = (n >= 8) ? '✓ 完了' : 'OK → 次へ';
    switch (n) {
        case 1: await renderStep1(); break;
        case 2: await renderStep2(); break;
        case 3: await renderStep3Decision(); break;
        case 4: await renderStep3(); break;
        case 5: renderStep4(); break;
        case 6: renderStep5(); break;
        case 7: await renderStep6(); break;
        case 8: renderStep7(); break;
    }
}

function handleBack() { if (currentStep > 1) showStep(currentStep - 1); }

async function handleOk() {
    switch (currentStep) {
        case 1: if (confirmStep1()) await showStep(2); break;
        case 2: if (confirmStep2()) await showStep(3); break;
        case 3: if (confirmStep3Decision()) await showStep(4); break;
        case 4: if (confirmStep3()) await showStep(5); break;
        case 5: if (confirmStep4()) await showStep(6); break;
        case 6: if (confirmStep5()) await showStep(7); break;
        case 7: if (confirmStep6()) renderFinal(); break;
        case 8:
            await saveWizardData();
            const hasNewItems = wizardState.schedule.some(s => s.isNew);
            if (!hasNewItems) {
                closeWizard();
            } else {
                confirmStep7();
            }
            break;
    }
}

async function saveWizardData() {
    el.okBtn.disabled = true;
    const oldText = el.okBtn.textContent;
    el.okBtn.textContent = '保存中...';
    try {
        const destConfig = await getDestinationConfig();
        const reportText = wizardState.finalReport;

        // ユーザーメールの取得
        let userEmail = 'unknown';
        if (chrome.identity) {
            try {
                const userInfo = await new Promise(resolve => chrome.identity.getProfileUserInfo({ accountStatus: 'ANY' }, resolve));
                if (userInfo && userInfo.email) userEmail = userInfo.email;
            } catch (e) { }
        }

        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        const dateStr = `${tomorrow.getFullYear()}/${String(tomorrow.getMonth() + 1).padStart(2, '0')}/${String(tomorrow.getDate()).padStart(2, '0')}`;

        if (destConfig.enableSheets) {
            await saveWizardReportToSheet(authTokenRef, userEmail, reportText);

            // AI Impact Logの構造化保存
            if (wizardState.aiUsage && wizardState.aiUsage.length > 0) {
                try {
                    await saveAiImpactLog(authTokenRef, userEmail, wizardState.aiUsage, wizardState.valueCreation);
                } catch (impactErr) {
                    console.warn('AI Impact Logの保存に失敗しました:', impactErr);
                }
            }

            // 判断ログの保存
            if (wizardState.decisions && wizardState.decisions.length > 0) {
                try {
                    await saveDecisionLog(authTokenRef, userEmail, wizardState.decisions);
                } catch (decisionErr) {
                    console.warn('判断ログの保存に失敗しました:', decisionErr);
                }
            }
        }
        if (destConfig.enableNotion) {
            await saveWizardToNotion(reportText, dateStr);
        }
    } catch (e) {
        console.error('ウィザード自動保存エラー:', e);
        showToastGlobal('保存処理に一部失敗しました', 'error');
    } finally {
        el.okBtn.disabled = false;
        el.okBtn.textContent = oldText;
    }
}

function updateIndicator(n) { el.dots.forEach((dot, i) => { dot.classList.toggle('active', i + 1 === n); dot.classList.toggle('done', i + 1 < n); }); }
function updateStepLabel(n) { if (el.stepLabel) el.stepLabel.textContent = STEP_LABELS[n - 1] || ''; }

// ===== STEP 1 =====
async function renderStep1() {
    setContent('<div class="wizard-loading"><div class="spinner"></div><p>カレンダーを取得中...</p></div>');
    el.okBtn.disabled = true;
    try {
        wizardState.events = [];
        wizardState.aiUsagePrecheck = {};

        const scheduleMap = wizardState.scheduleData;
        (wizardState.rawEvents || []).forEach((e, idx) => {
            if (!e.summary) return;
            const eventId = e.id || `event-${idx}`;
            const sData = scheduleMap ? scheduleMap.get(eventId) : null;
            const isIncluded = sData ? (sData.included !== false) : !shouldExclude(e.summary);

            if (isIncluded) {
                wizardState.events.push(e);
                if (sData) {
                    if (sData.aiFlag === 'yes-using' || sData.aiFlag === 'yes-potential') {
                        wizardState.aiUsagePrecheck[e.summary] = true;
                    } else if (sData.aiFlag === 'no') {
                        wizardState.aiUsagePrecheck[e.summary] = false;
                    }
                }
            }
        });

        renderStep1UI();
    } catch (e) {
        setContent(`<p class="wizard-hint">カレンダーの取得に失敗しました。手動で業務を入力してください。</p>${renderStep1Form()}`);
        el.okBtn.disabled = false;
    }
}

function renderStep1UI() {
    const listHtml = wizardState.events.length
        ? wizardState.events.map((e, i) => `
            <li class="wizard-event-item">
              <span class="wizard-event-time">${formatTime(e.start?.dateTime)} - ${formatTime(e.end?.dateTime)}</span>
              <span>${escapeHtml(e.summary || '')}</span>
              <label class="wizard-pending-label">
                <input type="checkbox" class="wizard-pending-cb" data-idx="${i}">未完了
              </label>
            </li>`).join('')
        : '<li class="wizard-hint">取得した予定はありません</li>';

    setContent(`
      <p class="wizard-step-title">📅 今日のカレンダー予定</p>
      <p class="wizard-step-sub">昼休憩・移動・準備系は自動除外。「未完了」をチェックすると翌日に自動繰り越し</p>
      <ul class="wizard-event-list">${listHtml}</ul>
      <p class="wizard-step-title wizard-step-title--mt">➕ カレンダー以外の業務</p>
      <textarea class="wizard-textarea" id="wizard-extra-tasks" placeholder="例：A社さんミーティング&#10;資料作成（1件1行で入力）" rows="4">${escapeHtml(wizardState.additionalTasks)}</textarea>
    `);
    el.okBtn.disabled = false;
}

function renderStep1Form() {
    return `<p class="wizard-step-title">📝 今日行った業務を入力</p>
      <textarea class="wizard-textarea" id="wizard-extra-tasks" placeholder="例：A社さんミーティング&#10;資料作成" rows="5">${escapeHtml(wizardState.additionalTasks)}</textarea>`;
}

function confirmStep1() {
    const ta = document.getElementById('wizard-extra-tasks');
    wizardState.additionalTasks = ta ? ta.value.trim() : '';
    const checked = [];
    document.querySelectorAll('.wizard-pending-cb:checked').forEach(cb => {
        const i = parseInt(cb.dataset.idx, 10);
        if (!isNaN(i) && wizardState.events[i]) checked.push(wizardState.events[i].summary || '');
    });
    wizardState.pendingTasks = checked;
    wizardState.actionItems = [
        ...wizardState.events.map(e => ({ title: e.summary || '', source: 'calendar' })),
        ...wizardState.additionalTasks.split('\n').filter(l => l.trim()).map(l => ({ title: l.trim(), source: 'manual' }))
    ];
    if (wizardState.actionItems.length === 0) { showInlineError('少なくとも1つの業務を入力してください'); return false; }
    return true;
}

// ===== STEP 2 =====
async function renderStep2() {
    wizardState.aiUsage = wizardState.actionItems.map(a => {
        let defaultUsed = null;
        if (wizardState.aiUsagePrecheck && wizardState.aiUsagePrecheck[a.title] !== undefined) {
            defaultUsed = wizardState.aiUsagePrecheck[a.title];
        }
        return { task: a.title, used: defaultUsed, tool: null, oneliner: null, impactScore: null, timeSaved: null };
    });
    wizardState.aiCurrentIdx = 0;
    renderAiQuestion(0);
}

async function renderAiQuestion(idx) {
    if (idx >= wizardState.aiUsage.length) { renderAiSummary(); return; }
    const item = wizardState.aiUsage[idx];
    const progress = `${idx + 1} / ${wizardState.aiUsage.length}件目`;
    const suggestion = suggestForTask(item.task);

    let suggTools = suggestion ? suggestion.tools.filter(t => !DELETED_TOOLS.includes(t)) : [];
    const suggHtml = suggTools.length > 0 ? `
      <div class="wizard-suggest-box">
        <p class="wizard-suggest-label">🎯 ${escapeHtml(suggestion.category)} の推奨</p>
        <div class="wizard-choices" id="wizard-tool-suggest">
          ${suggTools.map(t => `<button class="wizard-choice-btn wizard-suggest-btn" data-val="${escapeHtml(t)}">${escapeHtml(t)}</button>`).join('')}
        </div>
      </div>` : '';

    // UIを生成するためのヘルパー（削除ボタン付きかどうか）
    const makeToolBtn = (t) => `<div class="wizard-choice-btn-wrapper"><button class="wizard-choice-btn" data-val="${escapeHtml(t)}">${escapeHtml(t)}</button><button class="wizard-delete-btn" data-delete-tool="${escapeHtml(t)}">✕</button></div>`;
    const makeOnelinerBtn = (s, isHighlight) => `<div class="wizard-choice-btn-wrapper"><button class="wizard-choice-btn${isHighlight ? ' wizard-suggest-btn' : ''}" data-val="${escapeHtml(s)}">${escapeHtml(s)}</button><button class="wizard-delete-btn" data-delete-oneliner="${escapeHtml(s)}">✕</button></div>`;

    const populateOneliners = (selectedTool) => {
        const customList = CUSTOM_ONELINERS.filter(s => !DELETED_ONELINERS.includes(s));

        let profileOneliners = [];
        const taskLower = item.task.toLowerCase();
        for (const wp of workProfileMapping) {
            if (taskLower.includes(wp.category.toLowerCase())) profileOneliners.push(...wp.oneliners);
        }

        let mapOneliners = [];
        if (suggestion) mapOneliners.push(...suggestion.oneliners);
        const mergedMatched = Array.from(new Set([...profileOneliners, ...mapOneliners])).filter(s => !DELETED_ONELINERS.includes(s));
        const defOneliners = AI_ONELINER_SUGGESTIONS.filter(s => !DELETED_ONELINERS.includes(s));

        const getScore = (s) => {
            let score = 0;
            const sLower = s.toLowerCase();
            if (wizardState.sessionGeneratedOneliners && wizardState.sessionGeneratedOneliners.includes(s)) score += 100;
            if (mergedMatched.includes(s)) score += 50;

            let overlapFound = false;
            if (taskLower.length >= 2) {
                for (let i = 0; i < taskLower.length - 1; i++) {
                    const bg = taskLower.slice(i, i + 2);
                    if (/[ぁ-ん]{2}/.test(bg)) continue;
                    if (sLower.includes(bg)) { score += 20; overlapFound = true; }
                }
            }
            if (taskLower.length > 2 && sLower.includes(taskLower)) { score += 30; overlapFound = true; }
            if (customList.includes(s)) {
                if (overlapFound) score += 40;
                else score += 5;
            }
            return score;
        };

        const allCandidates = Array.from(new Set([...customList, ...mergedMatched, ...defOneliners]));
        const scored = allCandidates.map(s => ({ text: s, score: getScore(s) })).sort((a, b) => b.score - a.score);

        const relevantItems = scored.filter(s => s.score >= 40);
        const irrelevantItems = scored.filter(s => s.score < 40);

        const MAX_TOTAL = 10;
        const MAX_IRRELEVANT = 5;

        let finalItems = relevantItems.slice(0, MAX_TOTAL);
        if (finalItems.length < MAX_TOTAL) {
            const slots = Math.min(MAX_TOTAL - finalItems.length, MAX_IRRELEVANT);
            finalItems.push(...irrelevantItems.slice(0, slots));
        }

        const onelinerHtml = finalItems.map(item => {
            const isHighlight = mergedMatched.includes(item.text) || (wizardState.sessionGeneratedOneliners && wizardState.sessionGeneratedOneliners.includes(item.text));
            return makeOnelinerBtn(item.text, isHighlight);
        }).join('') + `<div class="wizard-choice-btn-wrapper"><button class="wizard-choice-btn" data-val="__other__">その他</button></div>`;

        const container = document.getElementById('wizard-oneliner-choices');
        if (container) container.innerHTML = onelinerHtml;
    };

    setContent(`
      <p class="wizard-step-title">🤖 AI活用確認 <small class="wizard-progress-label">${progress}</small></p>
      <div class="wizard-task-card" id="wizard-ai-card">
        <p class="wizard-task-card-title">「${escapeHtml(item.task)}」でAIを活用しましたか？</p>
        <div class="wizard-choices" id="wizard-used-choices">
          <button class="wizard-choice-btn ${item.used === 'no' || item.used === false ? 'selected' : ''}" data-val="no">活用していない</button>
          <button class="wizard-choice-btn ${item.used === 'yes' || item.used === true ? 'selected' : ''}" data-val="yes">活用した</button>
          <button class="wizard-choice-btn ${item.used === 'failed' ? 'selected' : ''}" data-val="failed">うまくいかなかった</button>
        </div>
        <div id="wizard-tool-area" class="${item.used === 'yes' || item.used === 'failed' || item.used === true ? '' : 'hidden'} wizard-sub-area">
          ${suggHtml}
          <p class="wizard-sublabel">${suggestion ? 'その他のツール:' : '使用ツール：'}</p>
          <div class="wizard-choices" id="wizard-tool-choices">
            ${AI_TOOLS.map(t => makeToolBtn(t)).join('')}
          </div>
          <div class="wizard-add-tool-row">
            <input type="text" class="wizard-add-tool-input hidden" id="wizard-add-tool-input" placeholder="ツール名を入力..." maxlength="50" />
            <button class="wizard-choice-btn" id="wizard-add-tool-btn">➕ ツール追加</button>
            <button class="wizard-choice-btn wizard-choice-btn--primary hidden" id="wizard-save-tool-btn">保存</button>
          </div>
        </div>
        <div id="wizard-oneliner-area" class="hidden wizard-sub-area">
          <p class="wizard-sublabel">一言（何に活用した？）：</p>
          <div class="wizard-choices" id="wizard-oneliner-choices">
            <!-- ツール選択時に動的生成される（最大20個） -->
          </div>
          <div class="wizard-add-tool-row wizard-mt-sm">
            <input type="text" class="wizard-add-tool-input hidden" id="wizard-add-oneliner-input" placeholder="カスタム一言を入力..." maxlength="100" />
            <button class="wizard-choice-btn" id="wizard-add-oneliner-btn">➕ 一言を保存</button>
            <button class="wizard-choice-btn wizard-choice-btn--primary hidden" id="wizard-save-oneliner-btn">保存</button>
            <button class="wizard-choice-btn wizard-suggest-btn" id="wizard-ai-generate-oneliners-btn" data-task="${escapeHtml(item.task)}">✨ AIで生成して追加</button>
          </div>
          <textarea class="wizard-textarea hidden wizard-mt-sm" id="wizard-oneliner-input" placeholder="一言を入力..." rows="2"></textarea>
          <div id="wizard-impact-area" class="hidden wizard-sub-area">
            <p class="wizard-sublabel">📈 事業インパクトスコア（どれくらい価値があったか）：</p>
            <div class="wizard-choices" id="wizard-impact-choices">
              <button class="wizard-choice-btn wizard-impact-btn tooltip" data-val="1" data-tooltip="Lv1: 作業の時短（文章校正、単純作業代替など）">★1</button>
              <button class="wizard-choice-btn wizard-impact-btn tooltip" data-val="2" data-tooltip="Lv2: 品質の底上げ（見落とし防止、構成改善など）">★2</button>
              <button class="wizard-choice-btn wizard-impact-btn tooltip" data-val="3" data-tooltip="Lv3: クライアント満足度向上（提案の質向上、速いレスポンス）">★3</button>
              <button class="wizard-choice-btn wizard-impact-btn tooltip" data-val="4" data-tooltip="Lv4: 新しい価値の提供（一人では出せなかったアイデア提供）">★4</button>
              <button class="wizard-choice-btn wizard-impact-btn tooltip" data-val="5" data-tooltip="Lv5: 業績直結（受注決定、契約単価アップなど）">★5</button>
            </div>
          </div>
          <div id="wizard-time-saved-area" class="hidden wizard-sub-area">
            <p class="wizard-sublabel">⏱️ AIで短縮できた推定時間（分）：</p>
            <input type="number" class="wizard-number-input" id="wizard-time-saved-input" placeholder="例: 30" min="0" step="5" />
          </div>
          <div id="wizard-failure-area" class="hidden wizard-sub-area">
            <p class="wizard-sublabel">⚠️ 何がうまくいかなかったか等のメモ：</p>
            <textarea class="wizard-textarea" id="wizard-failure-note-input" placeholder="例：プロンプトの指示が伝わらなかった、生成に時間がかかった..." rows="2">${escapeHtml(item.note || '')}</textarea>
          </div>
        </div>
      </div>
      <div class="wizard-choices wizard-mt-sm">
        <button class="wizard-choice-btn wizard-choice-btn--primary" id="wizard-next-task-btn">次の業務へ →</button>
      </div>
    `);
    el.okBtn.disabled = false;

    const suggestArea = document.getElementById('wizard-tool-suggest');
    if (suggestArea) {
        suggestArea.addEventListener('click', e => {
            if (!e.target.classList.contains('wizard-choice-btn')) return;
            document.querySelectorAll('#wizard-tool-suggest .wizard-choice-btn, #wizard-tool-choices .wizard-choice-btn').forEach(b => b.classList.remove('selected'));
            e.target.classList.add('selected');
            wizardState.aiUsage[idx].tool = e.target.dataset.val;
            populateOneliners(e.target.dataset.val);
            document.getElementById('wizard-oneliner-area').classList.remove('hidden');
        });
    }
    document.getElementById('wizard-add-tool-btn').addEventListener('click', () => {
        document.getElementById('wizard-add-tool-input').classList.remove('hidden');
        document.getElementById('wizard-save-tool-btn').classList.remove('hidden');
        document.getElementById('wizard-add-tool-input').focus();
    });
    document.getElementById('wizard-save-tool-btn').addEventListener('click', async () => {
        const inputEl = document.getElementById('wizard-add-tool-input');
        const val = inputEl.value.trim();
        if (!val) return;
        const added = await saveCustomTool(val);
        if (added) {
            const choices = document.getElementById('wizard-tool-choices');
            if (choices) {
                const wrapper = document.createElement('div');
                wrapper.className = 'wizard-choice-btn-wrapper';
                wrapper.innerHTML = `<button class="wizard-choice-btn" data-val="${escapeHtml(val)}">${escapeHtml(val)}</button><button class="wizard-delete-btn" data-delete-tool="${escapeHtml(val)}">✕</button>`;
                choices.appendChild(wrapper);
            }
            inputEl.value = '';
        }
        inputEl.classList.add('hidden');
        document.getElementById('wizard-save-tool-btn').classList.add('hidden');
    });
    document.getElementById('wizard-add-tool-input')?.addEventListener('keydown', e => { if (e.key === 'Enter') document.getElementById('wizard-save-tool-btn')?.click(); });

    // カスタム一言追加
    document.getElementById('wizard-add-oneliner-btn').addEventListener('click', () => {
        document.getElementById('wizard-add-oneliner-input').classList.remove('hidden');
        document.getElementById('wizard-save-oneliner-btn').classList.remove('hidden');
        document.getElementById('wizard-add-oneliner-input').focus();
    });
    document.getElementById('wizard-save-oneliner-btn').addEventListener('click', async () => {
        const inputEl = document.getElementById('wizard-add-oneliner-input');
        const val = inputEl.value.trim();
        if (!val) return;

        let added = await saveCustomOneliner(val);

        // 独自の機能：現在のカテゴリ（または予定タイトル）に紐づけて永続化設定に追加する
        const currentSuggestion = suggestForTask(item.task);
        const categoryLabel = (currentSuggestion && currentSuggestion.category) || item.task;
        await saveNewProfileFromWizard(categoryLabel, val);

        if (added) {
            const choices = document.getElementById('wizard-oneliner-choices');
            if (choices) {
                const wrapper = document.createElement('div');
                wrapper.className = 'wizard-choice-btn-wrapper';
                wrapper.innerHTML = `<button class="wizard-choice-btn" data-val="${escapeHtml(val)}">${escapeHtml(val)}</button><button class="wizard-delete-btn" data-delete-oneliner="${escapeHtml(val)}">✕</button>`;
                // "その他" の前に挿入
                const otherBtn = Array.from(choices.children).find(c => c.querySelector('button')?.dataset.val === '__other__');
                if (otherBtn) choices.insertBefore(wrapper, otherBtn);
                else choices.appendChild(wrapper);
            }
            inputEl.value = '';
        }
        inputEl.classList.add('hidden');
        document.getElementById('wizard-save-oneliner-btn').classList.add('hidden');
    });
    document.getElementById('wizard-add-oneliner-input')?.addEventListener('keydown', e => { if (e.key === 'Enter') document.getElementById('wizard-save-oneliner-btn')?.click(); });

    // AIで一言を自動生成するボタン
    document.getElementById('wizard-ai-generate-oneliners-btn')?.addEventListener('click', async (e) => {
        const btn = e.target;
        const taskName = btn.dataset.task;
        btn.textContent = '✨ 生成中...';
        btn.disabled = true;
        try {
            const prompt = `あなたはAIと業務効率化の専門家です。
現在のタスク「${sanitizeForPrompt(taskName, 100)}」において、AI（ChatGPT等）をどのように活用できるか、具体的で短い「一言」を5つ提案してください。文字数は10字〜20字程度にしてください。
以下のJSON配列で返してください：
["一言1", "一言2", "一言3", "一言4", "一言5"]`;
            const raw = await callGeminiRaw(prompt);
            const match = raw.match(/\[[\s\S]*\]/);
            if (match) {
                const arr = JSON.parse(match[0]);
                if (Array.isArray(arr)) {
                    const choices = document.getElementById('wizard-oneliner-choices');
                    if (choices) {
                        for (const s of arr) {
                            if (typeof s === 'string' && s.trim()) {
                                const val = s.trim().slice(0, 100);
                                const added = await saveCustomOneliner(val);
                                if (added) {
                                    if (wizardState.sessionGeneratedOneliners) wizardState.sessionGeneratedOneliners.push(val);
                                    const wrapper = document.createElement('div');
                                    wrapper.className = 'wizard-choice-btn-wrapper';
                                    wrapper.innerHTML = `<button class="wizard-choice-btn wizard-suggest-btn" data-val="${escapeHtml(val)}">${escapeHtml(val)}</button><button class="wizard-delete-btn" data-delete-oneliner="${escapeHtml(val)}">✕</button>`;
                                    const otherBtn = Array.from(choices.children).find(c => c.querySelector('button')?.dataset.val === '__other__');
                                    if (otherBtn) choices.insertBefore(wrapper, otherBtn);
                                    else choices.appendChild(wrapper);
                                }
                            }
                        }
                    }
                }
            }
        } catch (err) {
            console.error('AI生成エラー:', err);
            showInlineError('AI生成に失敗しました: ' + err.message);
        }
        btn.textContent = '✨ AIで生成して追加';
        btn.disabled = false;
    });

    // 削除ボタンのイベントリスナー（ツールと一言）
    el.content.addEventListener('click', async e => {
        if (e.target.dataset.deleteTool) {
            const val = e.target.dataset.deleteTool;
            if (confirm(`ツール「${val}」を削除しますか？`)) {
                await deleteCustomTool(val);
                e.target.closest('.wizard-choice-btn-wrapper')?.remove();
            }
        }
        if (e.target.dataset.deleteOneliner) {
            const val = e.target.dataset.deleteOneliner;
            if (confirm(`一言「${val}」を削除しますか？`)) {
                await deleteCustomOneliner(val);
                e.target.closest('.wizard-choice-btn-wrapper')?.remove();
            }
        }
    });

    document.getElementById('wizard-used-choices').addEventListener('click', e => {
        if (!e.target.classList.contains('wizard-choice-btn')) return;
        document.querySelectorAll('#wizard-used-choices .wizard-choice-btn').forEach(b => b.classList.remove('selected'));
        e.target.classList.add('selected');
        const val = e.target.dataset.val;
        wizardState.aiUsage[idx].used = val;
        const isUsed = val === 'yes' || val === 'failed';
        document.getElementById('wizard-tool-area').classList.toggle('hidden', !isUsed);
        document.getElementById('wizard-oneliner-area').classList.add('hidden');
        document.getElementById('wizard-impact-area').classList.add('hidden');
        document.getElementById('wizard-time-saved-area').classList.add('hidden');
        const failArea = document.getElementById('wizard-failure-area');
        if (failArea) failArea.classList.add('hidden');
    });
    document.getElementById('wizard-tool-choices').addEventListener('click', e => {
        if (!e.target.classList.contains('wizard-choice-btn')) return;
        document.querySelectorAll('#wizard-tool-choices .wizard-choice-btn').forEach(b => b.classList.remove('selected'));
        e.target.classList.add('selected');
        wizardState.aiUsage[idx].tool = e.target.dataset.val;
        populateOneliners(e.target.dataset.val);
        document.getElementById('wizard-oneliner-area').classList.remove('hidden');
        const isFailed = wizardState.aiUsage[idx].used === 'failed';
        document.getElementById('wizard-impact-area').classList.toggle('hidden', isFailed);
        document.getElementById('wizard-time-saved-area').classList.toggle('hidden', isFailed);
        const failArea = document.getElementById('wizard-failure-area');
        if (failArea) failArea.classList.toggle('hidden', !isFailed);
    });
    document.getElementById('wizard-oneliner-choices').addEventListener('click', e => {
        if (!e.target.classList.contains('wizard-choice-btn')) return;
        document.querySelectorAll('#wizard-oneliner-choices .wizard-choice-btn').forEach(b => b.classList.remove('selected'));
        e.target.classList.add('selected');
        const val = e.target.dataset.val;
        const inputEl = document.getElementById('wizard-oneliner-input');
        if (val === '__other__') { inputEl.classList.remove('hidden'); wizardState.aiUsage[idx].oneliner = null; }
        else { inputEl.classList.add('hidden'); wizardState.aiUsage[idx].oneliner = val; }
    });
    document.getElementById('wizard-impact-choices')?.addEventListener('click', e => {
        if (!e.target.classList.contains('wizard-choice-btn')) return;
        document.querySelectorAll('#wizard-impact-choices .wizard-choice-btn').forEach(b => b.classList.remove('selected'));
        e.target.classList.add('selected');
        wizardState.aiUsage[idx].impactScore = parseInt(e.target.dataset.val, 10);
    });

    document.getElementById('wizard-next-task-btn').addEventListener('click', async () => {
        const inputEl = document.getElementById('wizard-oneliner-input');
        if (!inputEl.classList.contains('hidden') && inputEl.value.trim()) wizardState.aiUsage[idx].oneliner = inputEl.value.trim();

        const timeSavedEl = document.getElementById('wizard-time-saved-input');
        if (timeSavedEl && !timeSavedEl.parentElement.classList.contains('hidden')) {
            const val = parseInt(timeSavedEl.value, 10);
            if (!isNaN(val) && val >= 0) wizardState.aiUsage[idx].timeSaved = val;
        }

        const failAreaEl = document.getElementById('wizard-failure-note-input');
        if (failAreaEl && !failAreaEl.parentElement.classList.contains('hidden')) {
            wizardState.aiUsage[idx].note = failAreaEl.value.trim();
        } else {
            wizardState.aiUsage[idx].note = '';
        }

        wizardState.aiCurrentIdx = idx + 1;
        await renderAiQuestion(idx + 1);
    });
}

function renderAiSummary() {
    const used = wizardState.aiUsage.filter(a => a.used === 'yes' || a.used === true);
    const failed = wizardState.aiUsage.filter(a => a.used === 'failed');
    const notUsed = wizardState.aiUsage.filter(a => a.used === 'no' || a.used === false);
    
    const usedHtml = used.map(a => `<li class="wizard-event-item"><span class="wizard-schedule-time">${escapeHtml(a.tool || '—')}</span><span>${escapeHtml(a.task)}<br><small>${escapeHtml(a.oneliner || '—')}</small></span></li>`).join('');
    const failedHtml = failed.map(a => `<li class="wizard-event-item" style="border-left: 3px solid #f43f5e;"><span class="wizard-schedule-time">${escapeHtml(a.tool || '—')}</span><span>${escapeHtml(a.task)}<br><small style="color: #e11d48;">⚠️ ${escapeHtml(a.note || '課題として記録')}</small></span></li>`).join('');
    const combinedHtml = [usedHtml, failedHtml].filter(Boolean).join('') || '<li class="wizard-hint">AI活用なし</li>';

    setContent(`<p class="wizard-step-title">🤖 AI活用まとめ</p><p class="wizard-step-sub">OKで次のステップへ進みます</p><ul class="wizard-event-list">${combinedHtml}</ul>${notUsed.length ? `<p class="wizard-hint">未活用: ${notUsed.map(a => escapeHtml(a.task)).join('、')}</p>` : ''}`);
    el.okBtn.disabled = false;
}

function confirmStep2() { return true; }

// ===== STEP 3: 判断の記録（タグ＋AI質問） =====
async function renderStep3Decision() {
    // AI活用データから判断ポイントを検出
    const points = detectDecisionPoints();

    if (points.length === 0) {
        // 判断ポイントなし → スキップ可能メッセージ
        setContent(`
          <p class="wizard-step-title">🧠 判断の記録</p>
          <p class="wizard-step-sub">今日の予定では特に目立った判断ポイントが見つかりませんでした。<br>次のステップへ進んでください。</p>
        `);
        el.okBtn.disabled = false;
        return;
    }

    // AI に質問を生成させる
    let aiQuestions = {};
    try {
        const summaryText = points.map(p =>
            `・「${sanitizeForPrompt(p.task)}」: ${sanitizeForPrompt(p.reason)}`
        ).join('\n');

        const prompt = `あなたは日報の「判断の記録」を支援するアシスタントです。
以下の業務における判断ポイントについて、ユーザーに「なぜそう判断したか」を引き出す短い質問を1つずつ生成してください。
質問は1行で、親しみやすい口調にしてください。

【判断ポイント一覧】
${summaryText}

以下のJSON形式で返してください：
{${points.map((p, i) => `"${i}": "質問文"`).join(', ')}}
JSONのみ返してください。`;

        const raw = await callGeminiRaw(prompt);
        const match = raw.match(/\{[\s\S]*\}/);
        if (match) {
            aiQuestions = JSON.parse(match[0]);
        }
    } catch (e) {
        console.warn('AI質問生成失敗:', e);
    }

    // 既存のdecisionsから復元
    const existingDecisions = wizardState.decisions || [];

    let cardsHtml = points.map((p, idx) => {
        const existing = existingDecisions[idx] || {};
        const question = aiQuestions[idx] || p.defaultQuestion;
        const selectedTags = existing.tags || [];
        const memo = existing.memo || '';

        const tagsHtml = DECISION_TAGS.map(t => {
            const isActive = selectedTags.includes(t.label);
            return `<button type="button" class="decision-tag${isActive ? ' active' : ''}" data-point="${idx}" data-tag="${escapeHtml(t.label)}">${t.emoji} ${escapeHtml(t.label)}</button>`;
        }).join('');

        return `
          <div class="decision-card" data-point="${idx}">
            <div class="decision-card-header">
              <span class="decision-card-task">📌 ${escapeHtml(p.task)}</span>
              <span class="decision-card-badge">${escapeHtml(p.badge)}</span>
            </div>
            <p class="decision-question">${escapeHtml(question)}</p>
            <div class="decision-tags-row">${tagsHtml}</div>
            <textarea class="wizard-textarea decision-memo" id="decision-memo-${idx}" placeholder="一言メモ（任意）" rows="2">${escapeHtml(memo)}</textarea>
          </div>`;
    }).join('');

    setContent(`
      <p class="wizard-step-title">🧠 今日の判断を記録</p>
      <p class="wizard-step-sub">AIが判断ポイントを検出しました。タグを選んで一言メモを残してください（スキップ可）。</p>
      <div class="decision-cards">${cardsHtml}</div>
    `);

    // タグ toggle イベント
    document.querySelectorAll('.decision-tag').forEach(btn => {
        btn.addEventListener('click', () => {
            btn.classList.toggle('active');
        });
    });

    el.okBtn.disabled = false;
}

/**
 * 判断ポイントの検出ロジック
 */
function detectDecisionPoints() {
    const items = wizardState.actionItems || [];
    const aiUsage = wizardState.aiUsage || [];
    const points = [];

    items.forEach((item, idx) => {
        const usage = aiUsage[idx];
        if (!usage) return;

        // Case 1: AI未活用を選択
        if (usage.used === false || usage.used === 'no') {
            points.push({
                task: item.title,
                reason: 'AI未活用を選択',
                badge: 'AI未活用',
                defaultQuestion: `「${item.title}」でAIを使わなかった理由は？`
            });
        }
        // Case 1.5: うまくいかなかったを選択
        else if (usage.used === 'failed') {
            points.push({
                task: item.title,
                reason: `AI活用課題: ${usage.note || '理由なし'}`,
                badge: 'AI課題',
                defaultQuestion: `「${item.title}」でのAI活用、何が一番のネックでしたか？`
            });
        }
        // Case 2: ツールが途中で変わった可能性（一言に「変更」「切替」等が含まれる）
        else if (usage.oneliner && /変更|切替|切り替|やめ|別の|代わり/.test(usage.oneliner)) {
            points.push({
                task: item.title,
                reason: `${usage.tool}を使用 — 「${usage.oneliner}」`,
                badge: 'ツール判断',
                defaultQuestion: `「${item.title}」でこの選択をした理由は？`
            });
        }
        // Case 3: インパクトスコアが低い（1-2）
        else if (usage.impactScore && usage.impactScore <= 2) {
            points.push({
                task: item.title,
                reason: `${usage.tool}使用、インパクト低(${usage.impactScore}/5)`,
                badge: '低インパクト',
                defaultQuestion: `「${item.title}」でAIの効果が低かった原因は？`
            });
        }
    });

    return points;
}

/**
 * Step 3 確定
 */
function confirmStep3Decision() {
    const cards = document.querySelectorAll('.decision-card');
    const decisions = [];
    cards.forEach((card, idx) => {
        const activeTags = Array.from(card.querySelectorAll('.decision-tag.active'))
            .map(btn => btn.dataset.tag);
        const memo = document.getElementById(`decision-memo-${idx}`)?.value.trim() || '';
        const task = card.querySelector('.decision-card-task')?.textContent.replace('📌 ', '') || '';

        if (activeTags.length > 0 || memo) {
            decisions.push({ task, tags: activeTags, memo });
        }
    });
    wizardState.decisions = decisions;
    return true; // 常にスキップ可能
}

// ===== STEP 3: 気づき・学び（目的×結果ペア + AI校正） =====
async function renderStep3() {
    setContent(`
      <p class="wizard-step-title">💡 得られた気づき・学び</p>
      <p class="wizard-step-sub">「目的」と「結果」をセットで入力してください。雑な内容でもOK — AIが補完・校正します。</p>

      <div class="form-group wizard-mt-sm">
        <label class="wizard-sublabel">📌 目的（何のために行ったか）</label>
        <textarea class="wizard-textarea" id="wizard-insight-purpose" placeholder="例：チーム内の情報共有のためにMTGを実施" rows="3">${escapeHtml(wizardState.insightPurpose || '')}</textarea>
      </div>

      <div class="form-group wizard-mt-sm">
        <label class="wizard-sublabel">📝 結果（何がわかった・得られたか）</label>
        <textarea class="wizard-textarea" id="wizard-insight-result" placeholder="例：来週までにやることが明確になった" rows="3">${escapeHtml(wizardState.insightResult || '')}</textarea>
      </div>

      <div class="form-group wizard-mt-sm">
        <label class="wizard-sublabel">🚀 価値創造アクション（AIで浮いた時間で何をしたか）</label>
        <textarea class="wizard-textarea" id="wizard-value-creation" placeholder="例：浮いた時間でクライアントへの追加提案資料を作成できた" rows="3">${escapeHtml(wizardState.valueCreation || '')}</textarea>
      </div>

      <div class="wizard-choices wizard-mt-sm">
        <button class="wizard-choice-btn wizard-choice-btn--primary" id="wizard-ai-polish-btn">🪄 AIで補完・校正</button>
      </div>

      <div id="wizard-ai-polish-result" class="hidden wizard-mt-sm">
        <label class="wizard-sublabel">✨ AI校正結果</label>
        <div class="form-group">
          <label class="wizard-sublabel"><small>📌 校正後の目的</small></label>
          <textarea class="wizard-textarea" id="wizard-polished-purpose" rows="3"></textarea>
        </div>
        <div class="form-group wizard-mt-sm">
          <label class="wizard-sublabel"><small>📝 校正後の結果</small></label>
          <textarea class="wizard-textarea" id="wizard-polished-result" rows="3"></textarea>
        </div>
        <p class="wizard-hint">校正結果を自由に編集できます。OKを押すと次のステップへ進みます。</p>
      </div>
    `);
    el.okBtn.disabled = false;

    // AI校正ボタン
    document.getElementById('wizard-ai-polish-btn').addEventListener('click', async () => {
        const purposeEl = document.getElementById('wizard-insight-purpose');
        const resultEl = document.getElementById('wizard-insight-result');
        const rawPurpose = purposeEl.value.trim();
        const rawResult = resultEl.value.trim();

        if (!rawPurpose && !rawResult) {
            showInlineError('目的または結果を入力してからAI校正を実行してください');
            return;
        }

        const btn = document.getElementById('wizard-ai-polish-btn');
        btn.disabled = true;
        btn.textContent = '🔄 AI校正中...';

        const actionSummary = wizardState.actionItems.map(a => `・${sanitizeForPrompt(a.title)}`).join('\n');
        const history = await loadInsightHistory();
        const historySection = history.length > 0
            ? `\n【過去の気づき（参考）】\n` +
            history.slice(-5).map(h => `・目的: ${h.purpose || h.text || ''}  結果: ${h.result || ''}`).join('\n') +
            `\n上記の文章スタイルを参考にしてください。`
            : '';

        const prompt = `あなたは日報の「気づき・学び」セクションを校正するアシスタントです。
ユーザーが雑に入力した「目的」と「結果」を、ビジネス日報にふさわしい自然な文章に補完・校正してください。
意味は変えずに、わかりやすく簡潔な文章にしてください。${historySection}

【今日の業務一覧（参考）】
${actionSummary}

【ユーザー入力 - 目的】
${sanitizeForPrompt(rawPurpose, 300) || '（未入力）'}

【ユーザー入力 - 結果】
${sanitizeForPrompt(rawResult, 300) || '（未入力）'}

以下のJSON形式で返してください：
{"purpose": "校正された目的の文章", "result": "校正された結果の文章"}
JSONのみ返してください。`;

        try {
            const raw = await callGeminiRaw(prompt);
            const match = raw.match(/\{[\s\S]*\}/);
            if (match) {
                const parsed = JSON.parse(match[0]);
                if (parsed.purpose && parsed.result) {
                    const resultArea = document.getElementById('wizard-ai-polish-result');
                    resultArea.classList.remove('hidden');
                    document.getElementById('wizard-polished-purpose').value = parsed.purpose;
                    document.getElementById('wizard-polished-result').value = parsed.result;
                } else {
                    showInlineError('AI校正結果の解析に失敗しました。手動で編集してください。');
                }
            } else {
                showInlineError('AI校正結果の解析に失敗しました。手動で編集してください。');
            }
        } catch (e) {
            console.error('AI校正エラー:', e);
            showInlineError('AI校正に失敗しました。手動で入力してください。');
        }

        btn.disabled = false;
        btn.textContent = '🪄 AIで補完・校正';
    });
}

function confirmStep3() {
    // 校正結果が表示されている場合はそちらを優先
    const polishedArea = document.getElementById('wizard-ai-polish-result');
    let purpose, result;
    if (polishedArea && !polishedArea.classList.contains('hidden')) {
        purpose = document.getElementById('wizard-polished-purpose')?.value.trim() || '';
        result = document.getElementById('wizard-polished-result')?.value.trim() || '';
    } else {
        purpose = document.getElementById('wizard-insight-purpose')?.value.trim() || '';
        result = document.getElementById('wizard-insight-result')?.value.trim() || '';
    }

    if (!purpose && !result) {
        showInlineError('目的または結果を入力してください');
        return false;
    }
    wizardState.insightPurpose = purpose;
    wizardState.insightResult = result;
    wizardState.valueCreation = document.getElementById('wizard-value-creation')?.value.trim() || '';
    // 学習のために保存
    saveInsightHistory(purpose, result);
    return true;
}

// ===== STEP 4: 明日挑戦・改善 =====
function renderStep4() {
    setContent(`
      <p class="wizard-step-title">🎯 明日挑戦したいこと・改善したいこと</p>
      <p class="wizard-step-sub">明日の業務で意識したいことや、直したい点を入力してください。</p>
      <div class="form-group wizard-mt-sm">
        <textarea class="wizard-textarea" id="wizard-try-and-improve" placeholder="例：〇〇のタスクの自動化を試す、〇〇の作業時間を半分にする" rows="4">${escapeHtml(wizardState.tryAndImprove || '')}</textarea>
      </div>
    `);
    el.okBtn.disabled = false;
}

function confirmStep4() {
    wizardState.tryAndImprove = document.getElementById('wizard-try-and-improve')?.value.trim() || '';
    return true;
}

// ===== STEP 5: その他 =====
function renderStep5() {
    setContent(`
      <p class="wizard-step-title">📝 その他共有事項・雑記</p>
      <p class="wizard-step-sub">チームに共有したいことや、単なるメモ・つぶやきがあればご記載ください。</p>
      <div class="form-group wizard-mt-sm">
        <textarea class="wizard-textarea" id="wizard-other-notes" placeholder="あれば入力してください（任意）" rows="4">${escapeHtml(wizardState.otherNotes || '')}</textarea>
      </div>
    `);
    el.okBtn.disabled = false;
}

function confirmStep5() {
    wizardState.otherNotes = document.getElementById('wizard-other-notes')?.value.trim() || '';
    return true;
}

// ===== STEP 6: 明日の予定 =====
async function renderStep6() {
    setContent('<div class="wizard-loading"><div class="spinner"></div><p>明日の予定を取得中...</p></div>');
    el.okBtn.disabled = true;
    try {
        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        const events = await getEventsForDate(authTokenRef, tomorrow);
        wizardState.tomorrowEvents = (events || []).filter(e => e.summary && !shouldExclude(e.summary));
    } catch (e) { wizardState.tomorrowEvents = []; }
    renderStep6UI();
}

function renderStep6UI() {
    const tomorrowStr = (() => { const d = new Date(); d.setDate(d.getDate() + 1); return `${d.getMonth() + 1}/${d.getDate()}`; })();
    const existHtml = wizardState.tomorrowEvents.length
        ? wizardState.tomorrowEvents.map(e => `<li class="wizard-event-item"><span class="wizard-event-time">${formatTime(e.start?.dateTime)} - ${formatTime(e.end?.dateTime)}</span><span>${escapeHtml(e.summary || '')}</span></li>`).join('')
        : '<li class="wizard-hint">取得した予定はありません</li>';

    setContent(`
      <p class="wizard-step-title">🗓️ ${tomorrowStr} の予定</p>
      <ul class="wizard-event-list">${existHtml}</ul>
      <p class="wizard-step-title wizard-step-title--mt">➕ 追加タスク（翌日新規）</p>
      <p class="wizard-step-sub">タスク名と所要時間を入力して「➕追加」を押してください</p>
      <div id="wizard-task-list" class="wizard-task-added-list"></div>
      <div class="wizard-add-task-row">
        <input type="text" id="wizard-new-task-name" class="wizard-add-tool-input" placeholder="例：A社さん対応" maxlength="80" />
        <select id="wizard-new-task-time" class="wizard-time-select">
          <option value="30">30分</option>
          <option value="60" selected>1時間</option>
          <option value="90">1.5時間</option>
          <option value="120">2時間</option>
          <option value="180">3時間</option>
          <option value="240">4時間</option>
        </select>
        <button class="wizard-choice-btn wizard-choice-btn--primary" id="wizard-add-task-btn">➕ 追加</button>
      </div>
      <label class="wizard-auto-label">
        <input type="checkbox" id="wizard-auto-schedule" checked>
        空き時間に自動配置する
      </label>
    `);
    el.okBtn.disabled = false;

    if (!wizardState._tomorrowTaskItems) wizardState._tomorrowTaskItems = [];
    const taskList = document.getElementById('wizard-task-list');
    wizardState._tomorrowTaskItems.forEach((item, i) => renderTaskRow(taskList, item, i));

    document.getElementById('wizard-add-task-btn').addEventListener('click', () => {
        const nameEl = document.getElementById('wizard-new-task-name');
        const timeEl = document.getElementById('wizard-new-task-time');
        const name = nameEl.value.trim();
        if (!name) { nameEl.focus(); return; }
        const minutes = parseInt(timeEl.value, 10);
        const item = { title: name, minutes };
        wizardState._tomorrowTaskItems.push(item);
        renderTaskRow(taskList, item, wizardState._tomorrowTaskItems.length - 1);
        nameEl.value = ''; timeEl.value = '60'; nameEl.focus();
    });
    document.getElementById('wizard-new-task-name').addEventListener('keydown', e => { if (e.key === 'Enter') document.getElementById('wizard-add-task-btn').click(); });
}

function renderTaskRow(container, item, idx) {
    const div = document.createElement('div');
    div.className = 'wizard-added-task-row';
    div.dataset.idx = idx;
    const timeLabel = item.minutes >= 60 ? `${item.minutes / 60}時間` : `${item.minutes}分`;
    div.innerHTML = `<span class="wizard-task-name-text"></span><span class="wizard-task-time-badge">${escapeHtml(timeLabel)}</span><button class="wizard-task-delete-btn" data-idx="${idx}">✕</button>`;
    div.querySelector('.wizard-task-name-text').textContent = item.title;
    div.querySelector('.wizard-task-delete-btn').addEventListener('click', e => {
        const i = parseInt(e.target.dataset.idx, 10);
        wizardState._tomorrowTaskItems.splice(i, 1);
        container.innerHTML = '';
        wizardState._tomorrowTaskItems.forEach((it, j) => renderTaskRow(container, it, j));
    });
    container.appendChild(div);
}

function confirmStep6() {
    const autoSchedule = document.getElementById('wizard-auto-schedule')?.checked !== false;
    wizardState._autoSchedule = autoSchedule;
    wizardState.schedule = buildSchedule(wizardState.tomorrowEvents, wizardState._tomorrowTaskItems || [], autoSchedule);
    renderScheduleConfirm();
    return false;
}

function renderScheduleConfirm() {
    const todayHtml = wizardState.actionItems.map(a => `<li class="wizard-schedule-item"><span class="wizard-schedule-time">今日</span><span>${escapeHtml(a.title)}</span></li>`).join('');
    const html = wizardState.schedule.length
        ? wizardState.schedule.map(s => `<li class="wizard-schedule-item ${s.isNew ? 'added' : ''}"><span class="wizard-schedule-time">${formatTime(s.start?.toISOString())} - ${formatTime(s.end?.toISOString())}</span><span>${escapeHtml(s.title)}${s.isNew ? ' <small class="wizard-added-badge">（追加）</small>' : ''}</span></li>`).join('')
        : '<li class="wizard-hint">スケジュールがありません</li>';
    setContent(`
      <p class="wizard-step-title">🗓️ 明日のスケジュール（確認）</p>
      <p class="wizard-step-sub">🟢は新規追加。🔁は今日の未完了繰り越し</p>
      <ul class="wizard-schedule-list">${html}</ul>
      <p class="wizard-step-title wizard-step-title--mt">📌 今日の振り返り</p>
      <ul class="wizard-schedule-list">${todayHtml}</ul>
    `);
    el.okBtn.disabled = false;
    el.okBtn.textContent = '日報確認へ →';
    el.okBtn.onclick = () => { renderFinal(); el.okBtn.onclick = handleOk; };
}

// ===== スケジュール自動配置 =====
function buildSchedule(existingEvents, taskItems, autoSchedule = true) {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const dateStr = `${tomorrow.getFullYear()}-${String(tomorrow.getMonth() + 1).padStart(2, '0')}-${String(tomorrow.getDate()).padStart(2, '0')}`;

    const existing = existingEvents.map(e => ({
        title: e.summary || '', isNew: false,
        start: new Date(e.start?.dateTime || `${dateStr}T09:00:00`),
        end: new Date(e.end?.dateTime || `${dateStr}T10:00:00`),
    }));
    const busy = existing.map(e => ({ start: e.start, end: e.end }));
    busy.sort((a, b) => a.start - b.start);

    // 未完了タスクを先頭にマージ
    const pendingItems = (wizardState.pendingTasks || []).map(t => ({ title: '🔁繰越：' + t, minutes: 60 }));
    const extras = [...pendingItems, ...taskItems];

    if (!autoSchedule) {
        // 自動配置なし：時間未確定として追加
        return [...existing, ...extras.map(e => ({ title: e.title + '（時間調整中）', start: new Date(`${dateStr}T19:00:00`), end: new Date(`${dateStr}T19:00:00`), isNew: true }))].sort((a, b) => a.start - b.start);
    }

    const workStart = new Date(`${dateStr}T10:00:00`);
    const workEnd = new Date(`${dateStr}T19:00:00`);
    const lunchStart = new Date(`${dateStr}T12:00:00`);
    const lunchEnd = new Date(`${dateStr}T13:00:00`);

    const placed = [];
    for (const extra of extras) {
        const dur = extra.minutes * 60 * 1000;
        let cursor = new Date(workStart);
        let placed_flag = false;
        while (cursor.getTime() + dur <= workEnd.getTime()) {
            if (cursor < lunchEnd && new Date(cursor.getTime() + dur) > lunchStart) { cursor = new Date(lunchEnd); continue; }
            const overlap = busy.find(b => cursor < b.end && new Date(cursor.getTime() + dur) > b.start);
            if (overlap) { cursor = new Date(overlap.end); continue; }
            const s = new Date(cursor); const e = new Date(cursor.getTime() + dur);
            placed.push({ title: extra.title, start: s, end: e, isNew: true });
            busy.push({ start: s, end: e }); busy.sort((a, b) => a.start - b.start);
            cursor = new Date(e); placed_flag = true; break;
        }
        if (!placed_flag) placed.push({ title: extra.title + '（時間未確定）', start: workEnd, end: workEnd, isNew: true });
    }
    return [...existing, ...placed].sort((a, b) => a.start - b.start);
}

// ===== 最終出力 =====
function renderFinal() {
    currentStep = 8;
    updateIndicator(8); updateStepLabel(8);
    const actionSection = wizardState.actionItems.map(a => `* ${a.title}`).join('\n');
    const aiSection = wizardState.aiUsage.filter(a => a.used === 'yes' || a.used === true || a.used === 'failed').map(a => {
        let text = `* 業務名：${a.task}\n  * ツール：${a.tool || '—'}\n  * 一言：${a.oneliner || '—'}`;
        if (a.used === 'failed') {
            text += `\n  * ⚠️ 活用結果：うまくいかなかった\n  * 課題・理由：${a.note || '—'}`;
        } else {
            if (a.impactScore) text += `\n  * インパクト：Lv${a.impactScore}`;
            if (a.timeSaved) text += `\n  * 短縮時間：${a.timeSaved}分`;
        }
        return text;
    }).join('\n\n');
    const scheduleSection = wizardState.schedule.map(s => `* ${s.title}（${formatTime(s.start?.toISOString())} - ${formatTime(s.end?.toISOString())}）`).join('\n');
    let insightSection = `【目的】${wizardState.insightPurpose || '—'}\n【結果】${wizardState.insightResult || '—'}`;
    if (wizardState.valueCreation) insightSection += `\n【価値創造】${wizardState.valueCreation}`;
    const report = `◾️具体的な行動内容\n\n1) 具体的な行動内容\n${actionSection}\n\n2) AI活用した行動内容\n${aiSection || '（なし）'}\n\n◾️得られた気づき・学び\n\n${insightSection}\n\n◾️明日挑戦したいこと・改善したいこと\n\n${wizardState.tryAndImprove || '（なし）'}\n\n◾️翌日予定\n\n${scheduleSection}${wizardState.otherNotes ? `\n\n◾️その他共有事項\n\n${wizardState.otherNotes}` : ''}`;
    wizardState.finalReport = report;

    const hasNewItems = wizardState.schedule.some(s => s.isNew);

    setContent(`
      <p class="wizard-step-title">📄 日報（最終確認）</p>
      <pre class="wizard-final-output" id="wizard-final-text">${escapeHtml(report)}</pre>
      <button class="wizard-choice-btn wizard-copy-btn" id="wizard-copy-final">📋 コピー</button>
      <p class="wizard-hint">${hasNewItems ? 'OKを押すと翌日のカレンダー登録へ進みます' : '追加予定がないため、OKで保存してウィザードを終了します'}</p>
    `);
    document.getElementById('wizard-copy-final')?.addEventListener('click', () => {
        const text = document.getElementById('wizard-final-text')?.textContent || '';
        navigator.clipboard.writeText(text).then(() => showToastGlobal('日報をコピーしました', 'success'));
    });
    el.okBtn.disabled = false;
    el.okBtn.textContent = hasNewItems ? 'カレンダー登録へ →' : '✓ 保存して完了';
}

function confirmStep7() {
    renderStep7();
}

// ===== STEP 7: カレンダー登録 =====
function renderStep7() {
    currentStep = 9;
    const newItems = wizardState.schedule.filter(s => s.isNew);
    if (newItems.length === 0) {
        setContent('<p class="wizard-hint">追加タスクがないため、カレンダー登録は不要です。</p>');
        el.okBtn.textContent = '✓ 完了';
        el.okBtn.onclick = () => { closeWizard(); el.okBtn.onclick = handleOk; };
        return;
    }
    const listHtml = newItems.map((s, i) => `
      <li class="wizard-schedule-item added">
        <span class="wizard-schedule-time">${formatTime(s.start?.toISOString())} - ${formatTime(s.end?.toISOString())}</span>
        <span>${escapeHtml(s.title)}</span>
        <button class="wizard-choice-btn wizard-choice-btn--primary" id="wizard-cal-btn-${i}">▶ カレンダーに登録</button>
      </li>`).join('');
    setContent(`
      <p class="wizard-step-title">📅 カレンダー登録</p>
      <p class="wizard-step-sub">各タスクの登録ボタンを押し、開いたGoogleカレンダーで「保存」してください</p>
      <ul class="wizard-schedule-list">${listHtml}</ul>
    `);
    newItems.forEach((s, i) => {
        document.getElementById(`wizard-cal-btn-${i}`)?.addEventListener('click', () => {
            const startStr = toCalUrl(s.start);
            const endStr = toCalUrl(s.end);
            const title = encodeURIComponent(s.title.replace(/（時間未確定）|（時間調整中）/g, '').trim().slice(0, 200));
            const url = `https://calendar.google.com/calendar/r/eventedit?text=${title}&dates=${startStr}/${endStr}&ctz=Asia/Tokyo`;
            chrome.tabs.create({ url });
        });
    });
    el.okBtn.textContent = '✓ 完了';
    el.okBtn.onclick = () => { closeWizard(); el.okBtn.onclick = handleOk; };
}

// ===== ユーティリティ =====
function shouldExclude(title) { return EXCLUDE_KEYWORDS.some(kw => title.toLowerCase().includes(kw.toLowerCase())); }

function formatTime(iso) {
    if (!iso) return '--:--';
    try { const d = new Date(iso); return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`; }
    catch { return '--:--'; }
}

function toCalUrl(date) {
    if (!date) return '';
    const d = new Date(date);
    // ローカル時刻でカレンダーURLを生成（JST維持）
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}T${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function showToastGlobal(msg, type = 'info') { document.dispatchEvent(new CustomEvent('wizard-toast', { detail: { msg, type } })); }

// ===== DOM操作ヘルパー =====
function sanitizeHtmlTemplate(html) {
    return html
        .replace(/<script[\s\S]*?<\/script>/gi, '')
        .replace(/\son\w+\s*=\s*["'][^"']*["']/gi, '')
        .replace(/\son\w+\s*=\s*`[^`]*`/gi, '')
        .replace(/javascript\s*:/gi, '#');
}

function setContent(html) {
    const content = el.content;
    Array.from(content.children).forEach(child => { if (!child.id || child.id !== 'wizard-loading') child.remove(); });
    el.loading.classList.add('hidden');
    const div = document.createElement('div');
    div.innerHTML = sanitizeHtmlTemplate(html);
    content.appendChild(div);
}

function showInlineError(msg) {
    const existing = el.content.querySelector('.wizard-error');
    if (existing) existing.remove();
    const err = document.createElement('p');
    err.className = 'wizard-hint wizard-error';
    err.style.color = '#ef4444';
    err.textContent = msg;
    el.content.appendChild(err);
}
