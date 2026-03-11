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
