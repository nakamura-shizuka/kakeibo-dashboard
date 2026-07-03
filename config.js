/**
 * 🔧 設定・定数
 * =========================================================
 * GASのスクリプトプロパティとアプリ全体の定数を管理する
 */

// ===== スクリプトプロパティ =====
const PROPERTIES = PropertiesService.getScriptProperties();
const LINE_ACCESS_TOKEN = (PROPERTIES.getProperty('LINE_ACCESS_TOKEN') || "").trim();
const LINE_CHANNEL_SECRET = (PROPERTIES.getProperty('LINE_CHANNEL_SECRET') || "").trim();
const SPREADSHEET_ID = (PROPERTIES.getProperty('SPREADSHEET_ID') || "").trim();
const GEMINI_API_KEY = (PROPERTIES.getProperty('GEMINI_API_KEY') || "").trim();

// Gemini APIモデル名（全AI関数で共通使用）
const GEMINI_MODEL = 'gemini-2.5-flash';

// 月間予算（デフォルト値。設定シートで上書き可能）
const DEFAULT_MONTHLY_BUDGET = 120000;

// ===== Layer 1（月次確定ベース集計）の定数 =====
// 支出総額KPIは「月次確定シート＋家計簿シートの非カード明細」から算出する。
// 家計簿シートのカード明細（下記Method）は参考情報（カテゴリ内訳）にのみ使用し、
// 総額には算入しない（請求確定額との二重計上を防ぐ）。
const CARD_METHODS = ['三井住友カード', 'PayPayカード'];

// 月次確定シートの Card 列に使う正規名
const BILLING_CARDS = ['三井住友', 'PayPay', 'イオン'];

// Layer 1 シート名
const SHEET_MONTHLY_CONFIRMED = '月次確定';
const SHEET_BALANCE = '残高';

// 予備費率のデフォルト（月収に対する%）
const DEFAULT_RESERVE_RATE = 5;
