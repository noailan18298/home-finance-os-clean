// @ts-nocheck

'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

// LocalStorage keys are versioned so structural changes do not collide with older saved data.
// Monthly data is stored separately from global settings so switching months does not reset Supabase, theme, users, or dashboard preferences.
const STORAGE_KEY = 'family-finance-os-stable-v14';
const SETTINGS_STORAGE_KEY = 'family-finance-os-global-settings-v1';
const AUTH_STORAGE_KEY = 'family-finance-os-auth-session-v1';
const FX_RATES_STORAGE_KEY = 'family-finance-os-fx-rates-v1';
const DEFAULT_SUPABASE_PROFILE_ID = 'default-household';
const APP_BUILD_MARKER = 'finance-dashboard-build-v14';

// Monthly compare periods determine how many previous months are averaged against the current month.
// Controls the Monthly Compare ranges. Each option compares the current month against an average of earlier months.
const COMPARE_PERIODS = [
  { id: 'previous', label: 'חודש קודם', months: 1 },
  { id: 'quarter', label: '3 חודשים', months: 3 },
  { id: 'halfYear', label: '6 חודשים', months: 6 },
  { id: 'year', label: 'שנה', months: 12 },
  { id: 'all', label: 'כל התקופה', months: Infinity },
];

const TABS = [
  { id: 'dashboard', label: 'Dashboard' },
  { id: 'income', label: 'הכנסות' },
  { id: 'credit', label: 'הוצאות' },
  { id: 'savings', label: 'חיסכון' },
  { id: 'insights', label: 'תובנות חכמות' },
  { id: 'settings', label: 'הגדרות' },
];

const THEME_STYLES = {
  Sage: { accent: '#A8B59A', accentHover: '#97A788', soft: '#F4F6F1', text: '#66725E', page: 'bg-white text-neutral-800' },
  Warm: { accent: '#C49A6C', accentHover: '#B3875A', soft: '#FBF4EC', text: '#8A6742', page: 'bg-[#FCFAF7] text-neutral-800' },
  Minimal: { accent: '#111111', accentHover: '#2B2B2B', soft: '#F5F5F5', text: '#404040', page: 'bg-[#FAFAFA] text-neutral-900' },
  Dark: { accent: '#4E5B52', accentHover: '#647267', soft: '#1F1F1F', text: '#D1D5DB', page: 'bg-[#111111] text-white' },
};

// Each financial mode adjusts targets, warnings, and insight priority without changing raw data.
// Financial modes change thresholds and priorities across insights, budget warnings, and savings targets.
const FINANCIAL_MODES = {
  Survival: {
    label: 'Survival',
    savingsTarget: 5,
    budgetWarningAt: 65,
    strictness: 1.35,
    focus: 'קיצוץ הוצאות ושמירה על תזרים חיובי',
    priorityMetric: 'burnRate',
    notificationTone: 'אגרסיבי',
  },
  Stable: {
    label: 'Stable',
    savingsTarget: 20,
    budgetWarningAt: 80,
    strictness: 1,
    focus: 'איזון בין איכות חיים לחיסכון יציב',
    priorityMetric: 'savingsRate',
    notificationTone: 'מאוזן',
  },
  Growth: {
    label: 'Growth',
    savingsTarget: 25,
    budgetWarningAt: 90,
    strictness: 0.85,
    focus: 'הגדלת הכנסות, השקעה בצמיחה ושיפור Cash Flow',
    priorityMetric: 'cashFlow',
    notificationTone: 'צמיחה',
  },
  'Wealth Building': {
    label: 'Wealth Building',
    savingsTarget: 35,
    budgetWarningAt: 95,
    strictness: 0.75,
    focus: 'בניית הון, הגדלת נכסים ואופטימיזציה פיננסית',
    priorityMetric: 'netWorth',
    notificationTone: 'אופטימיזציה',
  },
};

const EXPENSE_CATEGORIES = [
  'מזון וצריכה',
  'מסעדות ובתי קפה',
  'פנאי, בידור וספורט',
  'תחבורה ורכבים',
  'טיסות ותיירות',
  'בריאות ורפואה',
  'פארם וקוסמטיקה',
  'אופנה והלבשה',
  'העברת כספים',
  'משיכת מזומן',
  'ספרים ודפוס',
  'ביטוחים',
  'מיסים ותשלומים',
  'דיור וחשבונות',
  'חיסכון והשקעות',
  'העברה פנימית',
  'מט״ח / ארנק אשראי',
  'הוצאות עסקיות',
  'שונות',
  'אחר',
];

const MAX_CATEGORY_MAP = {
  'מזון וצריכה': 'מזון וצריכה',
  'מסעדות ובתי קפה': 'מסעדות ובתי קפה',
  'פנאי, בידור וספורט': 'פנאי, בידור וספורט',
  'תחבורה ורכבים': 'תחבורה ורכבים',
  'טיסות ותיירות': 'טיסות ותיירות',
  'בריאות ורפואה': 'בריאות ורפואה',
  'פארם וקוסמטיקה': 'פארם וקוסמטיקה',
  'אופנה והלבשה': 'אופנה והלבשה',
  'העברת כספים': 'העברת כספים',
  'משיכת מזומן': 'משיכת מזומן',
  'ספרים ודפוס': 'ספרים ודפוס',
  'שונות': 'שונות',
};

const CATEGORY_BUDGETS = {
  'מזון וצריכה': 4000,
  'מסעדות ובתי קפה': 800,
  'תחבורה ורכבים': 1800,
  'אופנה והלבשה': 1200,
  'בריאות ורפואה': 800,
  'פארם וקוסמטיקה': 700,
  'פנאי, בידור וספורט': 600,
  'טיסות ותיירות': 1500,
  'העברת כספים': 1000,
  'שונות': 1000,
  'אחר': 1000,
};

// Merchant keywords are intentionally simple and editable: user corrections are saved in learnedRules.
// First-pass categorization rules for imported card transactions. User edits later become learnedRules.
const MERCHANT_CATEGORY_MAP = {
  wolt: 'מסעדות ובתי קפה',
  tenbis: 'מסעדות ובתי קפה',
  shufersal: 'מזון וצריכה',
  שופרסל: 'מזון וצריכה',
  רמי: 'מזון וצריכה',
  victory: 'מזון וצריכה',
  ויקטורי: 'מזון וצריכה',
  yellow: 'תחבורה ורכבים',
  דור: 'תחבורה ורכבים',
  פז: 'תחבורה ורכבים',
  fox: 'אופנה והלבשה',
  zara: 'אופנה והלבשה',
  superpharm: 'פארם וקוסמטיקה',
  סופרפארם: 'פארם וקוסמטיקה',
  כללית: 'בריאות ורפואה',
  netflix: 'פנאי, בידור וספורט',
  spotify: 'פנאי, בידור וספורט',
  icloud: 'פנאי, בידור וספורט',
  google: 'פנאי, בידור וספורט',
  apple: 'פנאי, בידור וספורט',
};

const RECURRING_KEYWORDS = [
  'netflix', 'spotify', 'icloud', 'google', 'apple', 'cellcom', 'partner', 'pelephone', 'hot', 'yes',
  'ביטוח', 'הראל', 'מגדל', 'כלל', 'סלקום', 'פרטנר', 'פלאפון', 'שכירות',
];
const INTERNAL_TRANSFER_CATEGORIES = [
  'העברת כספים',
  'משיכת מזומן',
  'חיסכון והשקעות',
  'העברה פנימית',
  'מט״ח / ארנק אשראי',
];

function isInternalTransferTransaction(transaction) {
  const category = transaction?.category || '';
  const text = normalizeMerchantName(`${transaction?.merchant || ''} ${transaction?.description || ''} ${transaction?.notes || ''} ${transaction?.sourceSheet || ''}`);

  const isWalletOrCardSettlement =
    text.includes('רכישת מטח') ||
    text.includes('רכישת מט״ח') ||
    text.includes('ארנק מטח') ||
    text.includes('ארנק מט״ח') ||
    text.includes('עסקאות בחיוב מיידי') ||
    text.includes('חיוב חודשי') ||
    text.includes('חיוב כרטיס') ||
    text.includes('כרטיס אשראי') ||
    text.includes('סיכום חיובים') ||
    text.includes('חיוב מקס') ||
    text.includes('מקס איט פיננ');

  return INTERNAL_TRANSFER_CATEGORIES.includes(category) || isWalletOrCardSettlement;
}

const DEFAULT_FX_RATES_TO_ILS = {
  ILS: 1,
  USD: 3.7,
  EUR: 4.0,
  GBP: 4.7,
};

function getStoredFxRates() {
  const saved = safeJsonParse(getStorageItem(FX_RATES_STORAGE_KEY), null);
  if (!saved || typeof saved !== 'object') return DEFAULT_FX_RATES_TO_ILS;
  return { ...DEFAULT_FX_RATES_TO_ILS, ...(saved.rates || saved) };
}

function setStoredFxRates(rates) {
  setStorageItem(FX_RATES_STORAGE_KEY, JSON.stringify({
    updatedAt: new Date().toISOString(),
    rates: { ...DEFAULT_FX_RATES_TO_ILS, ...rates },
  }));
}

async function fetchFxRatesToIls() {
  try {
    const response = await fetch('https://api.frankfurter.dev/v1/latest?base=ILS&symbols=USD,EUR,GBP');
    if (!response.ok) throw new Error('FX request failed');

    const data = await response.json();
    const ratesFromIls = data?.rates || {};

    const ratesToIls = {
      ILS: 1,
      USD: ratesFromIls.USD ? 1 / ratesFromIls.USD : DEFAULT_FX_RATES_TO_ILS.USD,
      EUR: ratesFromIls.EUR ? 1 / ratesFromIls.EUR : DEFAULT_FX_RATES_TO_ILS.EUR,
      GBP: ratesFromIls.GBP ? 1 / ratesFromIls.GBP : DEFAULT_FX_RATES_TO_ILS.GBP,
    };

    setStoredFxRates(ratesToIls);
    return ratesToIls;
  } catch {
    return getStoredFxRates();
  }
}

function detectCurrencyFromRow(row) {
  const rawText = Array.isArray(row) ? row.join(' ') : String(row || '');
  const text = normalizeMerchantName(rawText);

  if (rawText.includes('$') || text.includes('usd') || text.includes('דולר')) return 'USD';
  if (rawText.includes('€') || text.includes('eur') || text.includes('יורו')) return 'EUR';
  if (rawText.includes('£') || text.includes('gbp') || text.includes('לישט')) return 'GBP';

  return 'ILS';
}

function convertToIls(amount, currency, fxRates = DEFAULT_FX_RATES_TO_ILS) {
  const rate = fxRates?.[currency] || 1;
  return toNumber(amount) * rate;
}

const SHEKEL = new Intl.NumberFormat('he-IL', { style: 'currency', currency: 'ILS', maximumFractionDigits: 0 });

function getPublicEnv(key) {
  if (typeof process !== 'undefined' && process.env && process.env[key]) return process.env[key];
  return '';
}

// Safe parser prevents broken localStorage/backup JSON from crashing the app during startup.
function safeJsonParse(value, fallback) {
  try {
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function getStorageItem(key) {
  if (typeof window === 'undefined' || !window.localStorage) return null;
  try {
    return window.localStorage.getItem(key);
  } catch {
    return null;
  }
}

function setStorageItem(key, value) {
  if (typeof window === 'undefined' || !window.localStorage) return false;
  try {
    window.localStorage.setItem(key, value);
    return true;
  } catch {
    return false;
  }
}

// These defaults stay empty because Supabase connection details are now stored in global settings.
// These defaults stay empty because Supabase connection details are now stored in global settings.
const DEFAULT_SUPABASE_URL = 'https://weqrtoovivzbunaakpca.supabase.co';
const DEFAULT_SUPABASE_ANON_KEY = 'sb_publishable_5wCORNNAyljwtOZHY67oew_4iQz1o4_';

const SUPABASE_URL = DEFAULT_SUPABASE_URL;
const SUPABASE_ANON_KEY = DEFAULT_SUPABASE_ANON_KEY;

function getFinancialModeConfig(mode) {
  return FINANCIAL_MODES[mode] || FINANCIAL_MODES.Stable;
}

function getSafeTheme(themeName) {
  return THEME_STYLES[themeName] || THEME_STYLES.Sage;
}

function makeId(prefix = 'id') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function getCurrentMonthKey() {
  const now = new Date();
  try {
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Jerusalem',
      year: 'numeric',
      month: '2-digit',
    }).formatToParts(now);
    const year = parts.find((part) => part.type === 'year')?.value;
    const month = parts.find((part) => part.type === 'month')?.value;
    if (year && month) return `${year}-${month}`;
  } catch {
    // Fallback only. The main path uses Israel time so month selection does not drift around midnight UTC.
  }
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

// Normalizes currency-like values from manual inputs and imported bank/card files into numbers.
function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  let raw = String(value || '').trim();
  if (!raw) return 0;
  const isNegative = raw.includes('-') || (raw.includes('(') && raw.includes(')'));
  raw = raw
    .split('₪').join('')
    .split('(').join('')
    .split(')').join('')
    .split(' ').join('')
    .split(String.fromCharCode(160)).join('')
    .split('−').join('-')
    .split('–').join('-')
    .split('—').join('-')
    .trim();

  if (raw.includes(',') && raw.includes('.')) {
    const lastComma = raw.lastIndexOf(',');
    const lastDot = raw.lastIndexOf('.');
    raw = lastComma > lastDot ? raw.split('.').join('').replace(',', '.') : raw.split(',').join('');
  } else if (raw.includes(',')) {
    const parts = raw.split(',');
    raw = parts[parts.length - 1]?.length === 2 ? raw.replace(',', '.') : raw.split(',').join('');
  }

  let numeric = '';
  for (let index = 0; index < raw.length; index += 1) {
    const char = raw[index];
    if ((char >= '0' && char <= '9') || char === '.' || char === '-') numeric += char;
  }
  const number = Number(numeric);
  if (!Number.isFinite(number)) return 0;
  return isNegative ? -Math.abs(number) : number;
}

function formatPercent(value) {
  return `${Math.round(value || 0)}%`;
}

function monthLabel(monthKey) {
  if (!monthKey) return '';
  const [year, month] = monthKey.split('-');
  return `${month}/${year}`;
}

// Keeps short Hebrew text from leaving a single word orphaned on its own line.
// Keeps Hebrew UI copy from leaving a single orphan word on its own line.
function noSingleWordLine(text) {
  const cleanText = String(text || '').replace(/[ \t]+/g, ' ').trim();
  const parts = cleanText.split(' ').filter(Boolean);
  if (parts.length < 4) return cleanText.replace(/ /g, String.fromCharCode(160));
  const chunks = [];
  for (let index = 0; index < parts.length; index += 2) {
    chunks.push(parts.slice(index, index + 2).join(String.fromCharCode(160)));
  }
  if (chunks.length >= 2 && !chunks[chunks.length - 1].includes(String.fromCharCode(160))) {
    chunks[chunks.length - 2] = `${chunks[chunks.length - 2]}${String.fromCharCode(160)}${chunks.pop()}`;
  }
  return chunks.join(' ');
}

function normalizeMerchantName(merchant = '') {
  return String(merchant).toLowerCase().replace(/\s+/g, ' ').replace(/[.,:;|()\[\]{}]/g, '').trim();
}

// Category detection first uses user-learned merchant rules, then falls back to the built-in merchant map.
function detectCategory(merchant = '', learnedRules = {}, importedCategory = '') {
  const normalizedImportedCategory = normalizeMerchantName(importedCategory);
  for (const [maxCategory, appCategory] of Object.entries(MAX_CATEGORY_MAP)) {
    if (normalizedImportedCategory && normalizedImportedCategory.includes(normalizeMerchantName(maxCategory))) return appCategory;
  }

  const normalized = normalizeMerchantName(merchant);
  for (const [key, category] of Object.entries(learnedRules || {})) {
    if (normalized.includes(normalizeMerchantName(key))) return category;
  }
  for (const [key, category] of Object.entries(MERCHANT_CATEGORY_MAP)) {
    if (normalized.includes(normalizeMerchantName(key))) return category;
  }
  return 'אחר';
}
function detectNecessity(category = '', merchant = '') {
  const text = normalizeMerchantName(`${category} ${merchant}`);

  if (
    text.includes('דיור') ||
    text.includes('חשבונות') ||
    text.includes('מזון') ||
    text.includes('בריאות') ||
    text.includes('רפואה') ||
    text.includes('ביטוחים') ||
    text.includes('מיסים')
  ) {
    return 'חיוני';
  }

  if (
    text.includes('תחבורה') ||
    text.includes('רכבים') ||
    text.includes('פארם') ||
    text.includes('ספרים') ||
    text.includes('דפוס')
  ) {
    return 'חשוב';
  }

  return 'מותרות';
}

// A tiny CSV parser that supports quoted fields, commas, semicolons, and tabs.
// Parses CSV rows while respecting quoted cells, commas, semicolons, tabs, and escaped quotes.
function splitCsvLine(line) {
  const result = [];
  let current = '';
  let insideQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];
    if (char === '"' && nextChar === '"') {
      current += '"';
      index += 1;
    } else if (char === '"') {
      insideQuotes = !insideQuotes;
    } else if ((char === ',' || char === ';' || char === '\t') && !insideQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }

  result.push(current.trim());
  return result.map((cell) => cell.replace(/^"|"$/g, '').trim());
}

function findHeaderIndex(headers, keywords, fallbackIndex = -1) {
  const normalizedHeaders = headers.map((header) => normalizeMerchantName(header));
  const foundIndex = normalizedHeaders.findIndex((header) => keywords.some((keyword) => header.includes(normalizeMerchantName(keyword))));
  return foundIndex >= 0 ? foundIndex : fallbackIndex;
}

// When the amount column is not clearly named, pick the column with the most numeric-looking values.
// Finds the amount column by header names first, then falls back to the most numeric-looking column.
function findAmountIndex(headers, sampleRows) {
  const headerIndex = findHeaderIndex(headers, ['amount', 'סכום', 'חיוב', 'חובה', 'זכות', 'עסקה', 'debit', 'credit', 'charge', 'total', 'נטו', 'לתשלום'], -1);
  if (headerIndex >= 0) return headerIndex;
  const width = Math.max(headers.length, ...sampleRows.map((row) => row.length));
  const rowScores = Array.from({ length: width }).map((_, index) => {
    const numericCount = sampleRows.slice(0, 20).filter((row) => Math.abs(toNumber(row[index])) > 0).length;
    const numericSum = sampleRows.slice(0, 20).reduce((sum, row) => sum + Math.abs(toNumber(row[index])), 0);
    return { index, numericCount, numericSum };
  });
  rowScores.sort((a, b) => (b.numericCount - a.numericCount) || (b.numericSum - a.numericSum));
  return rowScores[0]?.numericCount > 0 ? rowScores[0].index : 2;
}

// Normalizes CSV/Excel rows into pending credit-card transactions.
// Converts raw CSV/Excel rows into pending transactions, even when bank exports use different Hebrew/English column names.
function normalizeImportedRows(rows, learnedRules = {}, fxRates = DEFAULT_FX_RATES_TO_ILS) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const cleanedRows = rows.map((row) => (Array.isArray(row) ? row.map((cell) => String(cell || '').trim()) : [])).filter((row) => row.some(Boolean));
  if (!cleanedRows.length) return [];

  const headerCandidates = cleanedRows.slice(0, 25);
  const headerRowIndex = headerCandidates.findIndex((row) => {
    const joined = row.join(' ').toLowerCase();
    return ['date', 'תאריך', 'amount', 'סכום', 'merchant', 'בית עסק', 'שם בית העסק', 'תיאור', 'פירוט', 'חיוב', 'זכות', 'חובה'].some((word) => joined.includes(word));
  });
  const hasHeader = headerRowIndex >= 0;
  const headers = hasHeader ? cleanedRows[headerRowIndex] : cleanedRows[0] || [];
  const dataRows = hasHeader ? cleanedRows.slice(headerRowIndex + 1) : cleanedRows;
  const sampleRows = dataRows.slice(0, 30);

  const dateIndex = hasHeader ? findHeaderIndex(headers, ['date', 'תאריך', 'תאריך עסקה', 'תאריך רכישה', 'תאריך חיוב'], 0) : 0;
  const merchantIndex = hasHeader ? findHeaderIndex(headers, ['merchant', 'בית עסק', 'שם בית העסק', 'שם בית עסק', 'ספק', 'תיאור', 'פירוט', 'שם', 'פרטים'], 1) : 1;
const importedCategoryIndex = hasHeader ? findHeaderIndex(headers, ['קטגוריה', 'category'], -1) : -1;
const currencyIndex = hasHeader ? findHeaderIndex(headers, ['מטבע', 'currency', 'סוג מטבע', 'curr'], -1) : -1;
  const amountIndex = hasHeader ? findHeaderIndex(headers, ['סכום חיוב', 'amount charged', 'חיוב', 'סכום', 'חובה', 'זכות', 'amount', 'charge', 'total'], -1) : findAmountIndex(headers, sampleRows);
  const finalAmountIndex = amountIndex >= 0 ? amountIndex : findAmountIndex(headers, sampleRows);

  return dataRows
    .map((row) => {
      const amountCell = finalAmountIndex >= 0 ? row[finalAmountIndex] : [...row].reverse().find((cell) => Math.abs(toNumber(cell)) > 0);
      const rawAmount = toNumber(amountCell);
const currency = detectCurrencyFromRow(currencyIndex >= 0 ? `${row[currencyIndex]} ${row.join(' ')}` : row);
const amountIls = convertToIls(rawAmount, currency, fxRates);
const rowText = row.join(' ');
const normalizedRowText = normalizeMerchantName(rowText);

const isCreditRefund =
  normalizedRowText.includes('זיכוי') ||
  normalizedRowText.includes('זכות') ||
  normalizedRowText.includes('החזר') ||
  normalizedRowText.includes('refund') ||
  normalizedRowText.includes('credit') ||
  String(amountCell || '').includes('-') ||
  rawAmount < 0;

const amount = isCreditRefund ? -Math.abs(amountIls) : Math.abs(amountIls);
      const date = row[dateIndex] || row.find((cell) => String(cell || '').includes('/')) || row.find((cell) => String(cell || '').includes('-')) || '';
      const merchant = row[merchantIndex] || row.find((cell, index) => index !== dateIndex && index !== finalAmountIndex && String(cell || '').trim() && Math.abs(toNumber(cell)) === 0) || 'עסקה';
      const importedCategory = importedCategoryIndex >= 0 ? row[importedCategoryIndex] : '';
      const normalizedMerchant = normalizeMerchantName(merchant);
      const isSummaryRow = normalizedMerchant.includes('סך הכל') || normalizedMerchant.includes('total') || normalizedMerchant.includes('סהכ');
      const detectedCategory = normalizeMerchantName(`${merchant} ${rowText}`).includes('רכישת מטח') || normalizeMerchantName(`${merchant} ${rowText}`).includes('רכישת מט״ח')
  ? 'מט״ח / ארנק אשראי'
  : detectCategory(merchant, learnedRules, importedCategory);

return {
  id: makeId('tx'),
  date,
  merchant,
  amount,
  originalAmount: Math.abs(rawAmount),
  currency,
  category: detectedCategory,
  necessity: detectNecessity(detectedCategory, merchant),
};
    })
    .filter((transaction) => Math.abs(transaction.amount) > 0 && normalizeMerchantName(transaction.merchant) !== normalizeMerchantName('עסקה') && !normalizeMerchantName(transaction.merchant).includes('סך הכל'));
}

function normalizeBankRows(rows) {
      
function normalizeBankRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const cleanedRows = rows.map((row) => (Array.isArray(row) ? row.map((cell) => String(cell || '').trim()) : [])).filter((row) => row.some(Boolean));
  if (!cleanedRows.length) return [];

  const headerCandidates = cleanedRows.slice(0, 25);
  const headerRowIndex = headerCandidates.findIndex((row) => {
    const joined = row.join(' ').toLowerCase();
    return ['תאריך', 'date', 'תיאור', 'פירוט', 'אסמכתא', 'חובה', 'זכות', 'יתרה', 'balance', 'debit', 'credit'].some((word) => joined.includes(word));
  });
  const hasHeader = headerRowIndex >= 0;
  const headers = hasHeader ? cleanedRows[headerRowIndex] : cleanedRows[0] || [];
  const dataRows = hasHeader ? cleanedRows.slice(headerRowIndex + 1) : cleanedRows;
  const sampleRows = dataRows.slice(0, 40);

  const dateIndex = hasHeader ? findHeaderIndex(headers, ['תאריך', 'date', 'תאריך פעולה', 'תאריך ערך'], 0) : 0;
  const descriptionIndex = hasHeader ? findHeaderIndex(headers, ['תיאור', 'פירוט', 'פרטים', 'שם', 'פעולה', 'description'], 1) : 1;
  const debitIndex = hasHeader ? findHeaderIndex(headers, ['חובה', 'debit', 'חיוב', 'משיכה'], -1) : -1;
  const creditIndex = hasHeader ? findHeaderIndex(headers, ['זכות', 'credit', 'הפקדה'], -1) : -1;
  const balanceIndex = hasHeader ? findHeaderIndex(headers, ['יתרה', 'balance', 'יתרה בשח', 'יתרה נוכחית'], -1) : -1;
  const amountIndex = debitIndex < 0 && creditIndex < 0 ? findAmountIndex(headers, sampleRows) : -1;

  return dataRows
    .map((row) => {
      const debit = debitIndex >= 0 ? Math.abs(toNumber(row[debitIndex])) : 0;
      const credit = creditIndex >= 0 ? Math.abs(toNumber(row[creditIndex])) : 0;
      const fallbackAmount = amountIndex >= 0 ? toNumber(row[amountIndex]) : 0;
      const amount = credit || debit ? credit - debit : fallbackAmount;
      const balance = balanceIndex >= 0 ? toNumber(row[balanceIndex]) : 0;
      const description = row[descriptionIndex] || row.find((cell, index) => index !== dateIndex && index !== debitIndex && index !== creditIndex && index !== balanceIndex && String(cell || '').trim() && Math.abs(toNumber(cell)) === 0) || 'תנועה בבנק';
      const date = row[dateIndex] || row.find((cell) => String(cell || '').includes('/') || String(cell || '').includes('-')) || '';
      const normalizedDescription = normalizeMerchantName(description);
      const isSummaryRow = normalizedDescription.includes('סך הכל') || normalizedDescription.includes('סהכ') || normalizedDescription.includes('total');
      return { id: makeId('banktx'), date, description, amount, debit, credit, balance };
    })
    .filter((transaction) => Math.abs(toNumber(transaction.amount)) > 0 && !normalizeMerchantName(transaction.description).includes('סך הכל'));
}

function parseBankCsvText(text) {
  const carriageReturn = String.fromCharCode(13);
  const lineFeed = String.fromCharCode(10);
  const normalizedText = String(text || '').split(carriageReturn).join('');
  const rawLines = normalizedText.split(lineFeed);
  const lines = rawLines.map((line) => line.trim()).filter(Boolean);
  return normalizeBankRows(lines.map((line) => splitCsvLine(line)));
}

function parseBankExcelArrayBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  return workbook.SheetNames.flatMap((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: '' });
    return normalizeBankRows(rows).map((transaction) => ({ ...transaction, sourceSheet: sheetName }));
  });
}

function getBankBalancesFromTransactions(transactions, fallbackOpening = 0, fallbackClosing = 0) {
  const rowsWithBalance = (transactions || []).filter((transaction) => Math.abs(toNumber(transaction.balance)) > 0);
  if (rowsWithBalance.length >= 2) {
    return {
      openingBalance: toNumber(rowsWithBalance[0].balance) - toNumber(rowsWithBalance[0].amount),
      closingBalance: toNumber(rowsWithBalance[rowsWithBalance.length - 1].balance),
    };
  }
  if (rowsWithBalance.length === 1) {
    const closingBalance = toNumber(rowsWithBalance[0].balance);
    const movement = (transactions || []).reduce((sum, transaction) => sum + toNumber(transaction.amount), 0);
    return { openingBalance: closingBalance - movement, closingBalance };
  }
  const movement = (transactions || []).reduce((sum, transaction) => sum + toNumber(transaction.amount), 0);
  const openingBalance = toNumber(fallbackOpening);
  const closingBalance = fallbackClosing ? toNumber(fallbackClosing) : openingBalance + movement;
  return { openingBalance, closingBalance };
}

function parseCsvText(text, learnedRules = {}, fxRates = DEFAULT_FX_RATES_TO_ILS) {
  const carriageReturn = String.fromCharCode(13);
  const lineFeed = String.fromCharCode(10);
  const normalizedText = String(text || '').split(carriageReturn).join('');
  const rawLines = normalizedText.split(lineFeed);
  const lines = rawLines.map((line) => line.trim()).filter(Boolean);
  return normalizeImportedRows(lines.map((line) => splitCsvLine(line)), learnedRules, fxRates);
}

// Reads the first sheet from an uploaded Excel file and sends it through the same transaction normalizer as CSV.
function parseExcelArrayBuffer(buffer, learnedRules = {}, fxRates = DEFAULT_FX_RATES_TO_ILS) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const allTransactions = workbook.SheetNames.flatMap((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: '' });
    return normalizeImportedRows(rows, learnedRules, fxRates).map((transaction) => ({
      ...transaction,
      sourceSheet: sheetName,
    }));
  });

  const seen = new Set();
  return allTransactions.filter((transaction) => {
    const key = [transaction.date, normalizeMerchantName(transaction.merchant), transaction.amount, transaction.sourceSheet].join('|');
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function normalizeIncomeRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const cleanedRows = rows.map((row) => (Array.isArray(row) ? row.map((cell) => String(cell || '').trim()) : [])).filter((row) => row.some(Boolean));
  if (!cleanedRows.length) return [];

  const headerCandidates = cleanedRows.slice(0, 15);
  const headerRowIndex = headerCandidates.findIndex((row) => {
    const joined = row.join(' ').toLowerCase();
    return ['income', 'salary', 'net', 'amount', 'הכנסה', 'שכר', 'נטו', 'סכום', 'שם', 'לתשלום', 'זכות'].some((word) => joined.includes(word));
  });
  const hasHeader = headerRowIndex >= 0;
  const headers = hasHeader ? cleanedRows[headerRowIndex] : cleanedRows[0] || [];
  const dataRows = hasHeader ? cleanedRows.slice(headerRowIndex + 1) : cleanedRows;
  const sampleRows = dataRows.slice(0, 25);

  const nameIndex = hasHeader ? findHeaderIndex(headers, ['name', 'שם', 'עובד', 'מקור', 'תיאור', 'פירוט', 'income', 'salary', 'הכנסה', 'שכר', 'מעסיק'], 0) : 0;
  const amountIndex = findAmountIndex(headers, sampleRows);

  return dataRows
    .map((row) => {
      const amountCell = amountIndex >= 0 ? row[amountIndex] : [...row].reverse().find((cell) => Math.abs(toNumber(cell)) > 0);
      const amount = Math.abs(toNumber(amountCell));
      const name = row[nameIndex] || row.find((cell, index) => index !== amountIndex && String(cell || '').trim() && Math.abs(toNumber(cell)) === 0) || 'הכנסה מיובאת';
      return { id: makeId('income'), name, amount };
    })
    .filter((income) => income.amount > 0);
}

function parseIncomeCsvText(text) {
  const carriageReturn = String.fromCharCode(13);
  const lineFeed = String.fromCharCode(10);
  const normalizedText = String(text || '').split(carriageReturn).join('');
  const rawLines = normalizedText.split(lineFeed);
  const lines = rawLines.map((line) => line.trim()).filter(Boolean);
  return normalizeIncomeRows(lines.map((line) => splitCsvLine(line)));
}

function parseIncomeExcelArrayBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) return [];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1, raw: false, defval: '' });
  return normalizeIncomeRows(rows);
}

function normalizeLooseText(text) {
  return String(text || '').split(String.fromCharCode(10)).join(' ').split(String.fromCharCode(13)).join(' ').split(String.fromCharCode(9)).join(' ').split('  ').join(' ').trim();
}

function extractLooseTextFromPdfArrayBuffer(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 8192;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.slice(index, index + chunkSize));
  }

  const textParts = [];
  let searchFrom = 0;
  while (searchFrom < binary.length) {
    const markerIndex = binary.indexOf(') Tj', searchFrom);
    if (markerIndex < 0) break;
    const openIndex = binary.lastIndexOf('(', markerIndex);
    if (openIndex >= 0 && markerIndex > openIndex) textParts.push(binary.slice(openIndex + 1, markerIndex));
    searchFrom = markerIndex + 4;
  }

  const extracted = normalizeLooseText(textParts.join(' '));
  return extracted || normalizeLooseText(binary.slice(0, 25000));
}

function extractFirstMoneyValueNear(text, keyword) {
  const source = normalizeLooseText(text);
  const keywordIndex = source.indexOf(keyword);
  if (keywordIndex < 0) return 0;
  const windowText = source.slice(keywordIndex, keywordIndex + 180);
  let current = '';
  const values = [];
  for (let index = 0; index < windowText.length; index += 1) {
    const char = windowText[index];
    const isDigit = char >= '0' && char <= '9';
    if (isDigit || char === ',' || char === '.') current += char;
    else if (current) {
      const value = toNumber(current);
      if (value > 1000) values.push(value);
      current = '';
    }
  }
  if (current) {
    const value = toNumber(current);
    if (value > 1000) values.push(value);
  }
  return values[0] || 0;
}

function extractNetSalaryFromPdfText(text) {
  const keywords = ['נטו לתשלום', 'לתשלום בבנק', 'סהכ לתשלום', 'סה״כ לתשלום', 'שכר נטו', 'נטו'];
  for (const keyword of keywords) {
    const value = extractFirstMoneyValueNear(text, keyword);
    if (value > 1000) return value;
  }
  return 0;
}

async function parseIncomePdfFile(file) {
  const buffer = await file.arrayBuffer();
  const text = extractLooseTextFromPdfArrayBuffer(buffer);
  const netSalary = extractNetSalaryFromPdfText(text);
  if (!netSalary) return [];
  return [{ id: makeId('income'), name: `נטו מתלוש ${file.name}`, amount: netSalary }];
}

function getCategoryTotals(transactions) {
  return transactions.reduce((acc, transaction) => {
    const category = transaction.category || 'אחר';
    acc[category] = (acc[category] || 0) + toNumber(transaction.amount);
    return acc;
  }, {});
}

function getMerchantTotals(transactions) {
  return transactions.reduce((acc, transaction) => {
    const merchant = transaction.merchant || 'ללא שם';
    acc[merchant] = (acc[merchant] || 0) + toNumber(transaction.amount);
    return acc;
  }, {});
}

// This score is heuristic only: it highlights concentration, large transactions, and budget pressure.
// Produces a simple 0-100 score based on concentration, uncategorized spend, large transactions, and mode strictness.
function calculateFinancialHealthScore(transactions, modeConfig = FINANCIAL_MODES.Stable) {
  if (!transactions.length) return null;
  const total = transactions.reduce((sum, item) => sum + toNumber(item.amount), 0) || 1;
  const categoryTotals = getCategoryTotals(transactions);
  const merchantTotals = getMerchantTotals(transactions);
  const uncategorizedAmount = categoryTotals['אחר'] || 0;
  const largestTransaction = Math.max(...transactions.map((item) => toNumber(item.amount)));
  const largestMerchantAmount = Math.max(...Object.values(merchantTotals).map(toNumber));
  const strictness = modeConfig?.strictness || 1;
  let score = 100;

  if (uncategorizedAmount / total > 0.2) score -= Math.round(15 * strictness);
  if (largestTransaction / total > 0.25) score -= Math.round(12 * strictness);
  if (largestMerchantAmount / total > 0.35) score -= Math.round(10 * strictness);

  Object.entries(CATEGORY_BUDGETS).forEach(([category, budget]) => {
    const spent = categoryTotals[category] || 0;
    const adjustedBudget = budget / strictness;
    if (spent > adjustedBudget) score -= Math.round(8 * strictness);
    else if (spent >= adjustedBudget * ((modeConfig?.budgetWarningAt || 80) / 100)) score -= Math.round(4 * strictness);
  });

  return Math.max(0, Math.min(100, score));
}

// Detects recurring payments using recurring keywords plus merchants seen in previous months.
function detectRecurringTransactions(transactions, historicalMonths = {}, selectedMonth = '') {
  const historicalMerchants = new Set();
  Object.entries(historicalMonths || {}).forEach(([month, data]) => {
    if (month === selectedMonth) return;
    const monthData = normalizeMonthData(data);
    monthData.creditCards.forEach((card) => {
      (card.transactions || []).forEach((transaction) => historicalMerchants.add(normalizeMerchantName(transaction.merchant)));
    });
  });

  return transactions.filter((transaction) => {
    const normalized = normalizeMerchantName(transaction.merchant);
    const keywordHit = RECURRING_KEYWORDS.some((keyword) => normalized.includes(normalizeMerchantName(keyword)));
    return keywordHit || historicalMerchants.has(normalized);
  });
}

// Month totals include only month-specific financial data. Global settings are handled through preferences.
function getMonthTotals(data) {
  const safeData = normalizeMonthData(data);
  const includeSelfEmployed = Boolean(safeData.preferences.includeSelfEmployed);
  const income = safeData.incomes.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const credit = safeData.creditCards.reduce((sum, card) => sum + (card.transactions || []).reduce((inner, item) => inner + toNumber(item.amount), 0), 0);
  const manual = safeData.manualExpenses.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const savingsProducts = safeData.savingsProducts.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const savingGoals = safeData.savingGoals.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const selfEmployedVatDue = Math.max(0, toNumber(safeData.selfEmployed.vatCollected) - toNumber(safeData.selfEmployed.vatPaidOnExpenses));
  const selfEmployedRaw = selfEmployedVatDue + toNumber(safeData.selfEmployed.incomeTaxAdvance) + toNumber(safeData.selfEmployed.nationalInsurance) + toNumber(safeData.selfEmployed.businessExpenses);
  const selfEmployed = includeSelfEmployed ? selfEmployedRaw : 0;
const savingsTransfers = savingsProducts + savingGoals;
const realExpenses = credit + manual + selfEmployed;
const remainingCashFlow = income - realExpenses - savingsTransfers;
const savingsRate = income ? (savingsTransfers / income) * 100 : 0;

return {
  income,
  credit,
  manual,
  savings: savingsTransfers,
  selfEmployed,
  selfEmployedRaw,
  expenses: realExpenses,
  net: remainingCashFlow,
  savingsRate,
};
}

function getPreviousMonthKey(monthKey) {
  if (!monthKey || !monthKey.includes('-')) return '';
  const [year, month] = monthKey.split('-').map(Number);
  const date = new Date(year, month - 2, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function getDataWeight(value) {
  if (Array.isArray(value)) return value.length;
  if (value && typeof value === 'object') return Object.keys(value).length;
  if (typeof value === 'number') return Math.abs(value) > 0 ? 1 : 0;
  return value ? 1 : 0;
}

function getMonthDataWeight(monthData) {
  const month = normalizeMonthData(monthData);
  return [
    month.incomes,
    month.manualExpenses,
    month.savingsProducts,
    month.savingGoals,
    month.creditCards.flatMap((card) => [...(card.transactions || []), ...(card.pendingTransactions || [])]),
    month.bankAccounts.flatMap((account) => account.transactions || []),
    month.bankAccounts.map((account) => [account.openingBalance, account.closingBalance]).flat(),
    month.attachedDocuments,
    month.emergencyFund,
  ].reduce((sum, value) => sum + getDataWeight(value), 0);
}

function mergeMonthKeepingRicher(localMonth, cloudMonth) {
  if (!localMonth) return normalizeMonthData(cloudMonth);
  if (!cloudMonth) return normalizeMonthData(localMonth);

  const localHasData = hasMeaningfulMonths({ localMonth });
  const cloudHasData = hasMeaningfulMonths({ cloudMonth });

  if (cloudHasData && !localHasData) return normalizeMonthData(cloudMonth);
  if (localHasData && !cloudHasData) return normalizeMonthData(localMonth);

  const localWeight = getMonthDataWeight(localMonth);
  const cloudWeight = getMonthDataWeight(cloudMonth);

  return normalizeMonthData(cloudWeight >= localWeight ? cloudMonth : localMonth);
}

function mergeMonthsKeepingRicher(localMonths = {}, cloudMonths = {}) {
  const merged = { ...localMonths };
  Object.keys(cloudMonths || {}).forEach((monthKey) => {
    merged[monthKey] = mergeMonthKeepingRicher(localMonths[monthKey], cloudMonths[monthKey]);
  });
  return merged;
}

function hasRollingData(monthData) {
  if (!monthData) return false;
  const month = normalizeMonthData(monthData);
  const bankHasData = month.bankAccounts.some((account) =>
    Math.abs(toNumber(account.openingBalance)) > 0 ||
    Math.abs(toNumber(account.closingBalance)) > 0 ||
    (account.transactions || []).length > 0
  );
  const savingsHasRealBalances = month.savingsProducts.some((product) =>
    Math.abs(toNumber(product.currentBalance)) > 0 ||
    Math.abs(toNumber(product.monthlyDeposit)) > 0
  );
  const goalsHaveRealBalances = month.savingGoals.some((goal) =>
    Math.abs(toNumber(goal.currentAmount)) > 0 ||
    Math.abs(toNumber(goal.monthlyDeposit)) > 0
  );
  return bankHasData || savingsHasRealBalances || goalsHaveRealBalances || Math.abs(toNumber(month.emergencyFund)) > 0;
}

function createRollingMonthFromPrevious(previousMonthData) {
  const base = createMonthFromPrevious(previousMonthData);
  if (!previousMonthData) return base;
  const previous = normalizeMonthData(previousMonthData);

  return {
    ...base,
    bankAccounts: previous.bankAccounts.map((account) => {
      const previousClosing = toNumber(account.closingBalance);
      return {
        ...account,
        id: makeId('bank'),
        // Bank balance is a real bank value only. It does not include calculated monthly net / "יתרה אחרי הכול".
        openingBalance: previousClosing,
        closingBalance: previousClosing,
        importedFile: '',
        transactions: [],
      };
    }),
    savingsProducts: previous.savingsProducts.map((product) => ({
      ...product,
      id: makeId('saving'),
      currentBalance: toNumber(product.currentBalance) + toNumber(product.monthlyDeposit),
    })),
    savingGoals: previous.savingGoals.map((goal) => ({
      ...goal,
      id: makeId('goal'),
      currentAmount: Math.min(
        toNumber(goal.targetAmount) || Number.POSITIVE_INFINITY,
        toNumber(goal.currentAmount) + toNumber(goal.monthlyDeposit)
      ),
    })),
    emergencyFund: toNumber(previous.emergencyFund),
  };
}

function repairCarriedBankBalances(targetMonthData, previousMonthData) {
  if (!previousMonthData) return normalizeMonthData(targetMonthData);
  const target = normalizeMonthData(targetMonthData);
  const previous = normalizeMonthData(previousMonthData);
  return {
    ...target,
    bankAccounts: target.bankAccounts.map((account, index) => {
      const previousAccount = previous.bankAccounts[index];
      const previousClosing = toNumber(previousAccount?.closingBalance);
      const hasBankImport = Boolean(account.importedFile) || (account.transactions || []).length > 0;
      const looksCarriedFromPrevious = Math.abs(toNumber(account.openingBalance) - previousClosing) < 1;
      if (!hasBankImport && looksCarriedFromPrevious) {
        return { ...account, closingBalance: toNumber(account.openingBalance) };
      }
      return account;
    }),
  };
}

function carryForwardRollingFields(targetMonthData, previousMonthData) {
  if (!previousMonthData) return normalizeMonthData(targetMonthData);
  const target = repairCarriedBankBalances(targetMonthData, previousMonthData);
  const carried = createRollingMonthFromPrevious(previousMonthData);
  if (hasRollingData(target)) return target;
  return {
    ...target,
    emergencyFund: carried.emergencyFund,
    bankAccounts: carried.bankAccounts,
    savingsProducts: carried.savingsProducts,
    savingGoals: carried.savingGoals,
  };
}

function createMonthFromPrevious(previousMonthData) {
  const base = createDefaultMonth();
  if (!previousMonthData) return base;
  const previous = normalizeMonthData(previousMonthData);

  return {
    ...base,
    dashboardTitle: previous.dashboardTitle || base.dashboardTitle,
    emergencyFund: previous.emergencyFund || 0,
    bankAccounts: previous.bankAccounts.map((account) => ({
      ...account,
      id: makeId('bank'),
      openingBalance: toNumber(account.closingBalance),
      closingBalance: toNumber(account.closingBalance),
      importedFile: account.importedFile || '',
      transactions: account.transactions || [],
    })),
    savingsProducts: previous.savingsProducts.map((product) => ({
      ...product,
      id: makeId('saving'),
    })),
    savingGoals: previous.savingGoals.map((goal) => ({
      ...goal,
      id: makeId('goal'),
    })),
    selfEmployed: { ...base.selfEmployed, owner: previous.selfEmployed.owner || base.selfEmployed.owner },
  };
}

// Collects earlier months for comparison, newest first, according to the selected compare period.
function getCompareMonthKeys(months, selectedMonth, periodId = 'previous') {
  const period = COMPARE_PERIODS.find((item) => item.id === periodId) || COMPARE_PERIODS[0];
  const sortedMonths = Object.keys(months || {}).filter((month) => month < selectedMonth).sort((a, b) => b.localeCompare(a));
  if (period.id === 'all') return sortedMonths;
  return sortedMonths.slice(0, period.months);
}

// Averages totals across selected months so compare can work against 3 months, 6 months, a year, or all history.
function averageTotals(months, monthKeys) {
  const empty = { income: 0, credit: 0, manual: 0, savings: 0, selfEmployed: 0, expenses: 0, net: 0, savingsRate: 0 };
  if (!monthKeys.length) return empty;
  const totals = monthKeys.reduce((acc, monthKey) => {
    const monthTotals = getMonthTotals(months[monthKey]);
    Object.keys(empty).forEach((key) => {
      acc[key] += monthTotals[key] || 0;
    });
    return acc;
  }, { ...empty });
  Object.keys(totals).forEach((key) => {
    totals[key] = totals[key] / monthKeys.length;
  });
  return totals;
}

function getMonthlyCompare(months, selectedMonth, periodId = 'previous') {
  const current = getMonthTotals(months[selectedMonth]);
  const compareMonthKeys = getCompareMonthKeys(months, selectedMonth, periodId);
  const period = COMPARE_PERIODS.find((item) => item.id === periodId) || COMPARE_PERIODS[0];
  if (!compareMonthKeys.length) {
    return {
      period,
      compareMonthKeys,
      hasPrevious: false,
      previousMonth: getPreviousMonthKey(selectedMonth),
      current,
      previous: null,
      rows: [],
    };
  }
  const previous = averageTotals(months, compareMonthKeys);
  const buildRow = (label, key, type = 'currency') => {
    const currentValue = current[key] || 0;
    const previousValue = previous[key] || 0;
    const diff = currentValue - previousValue;
    const percent = previousValue ? (diff / previousValue) * 100 : null;
    return { label, key, type, currentValue, previousValue, diff, percent };
  };
  return {
    period,
    compareMonthKeys,
    previousMonth: compareMonthKeys[0] || '',
    hasPrevious: true,
    current,
    previous,
    rows: [
      buildRow('הכנסות', 'income'),
      buildRow('הוצאות', 'expenses'),
      buildRow('אשראי', 'credit'),
      buildRow('הוצאות ידניות', 'manual'),
      buildRow('חיסכון', 'savings'),
      buildRow('עצמאי', 'selfEmployed'),
      buildRow('יתרה', 'net'),
      buildRow('שיעור חיסכון', 'savingsRate', 'percent'),
    ],
  };
}

function getMonthlyTrend(months) {
  return Object.entries(months || {})
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([month, data]) => {
      const totals = getMonthTotals(data);
      return {
        month,
        total: totals.credit + totals.manual,
        income: totals.income,
        expenses: totals.expenses,
        savings: totals.net,
      };
    });
}

// Builds deterministic insights from real user-entered/imported data without calling an external AI service.
function buildRealInsights(transactions, recurringTransactions = [], totalIncome = 0, financialMode = 'Stable', context = {}) {
  const modeConfig = getFinancialModeConfig(financialMode);
  if (!transactions.length) return [`מצב ${modeConfig.label}: ${modeConfig.focus}. העלו CSV או Excel כדי לקבל תובנות.`];

  const total = transactions.reduce((sum, item) => sum + toNumber(item.amount), 0) || 1;
  const categoryTotals = getCategoryTotals(transactions);
  const merchantTotals = getMerchantTotals(transactions);
  const sortedCategories = Object.entries(categoryTotals).sort((a, b) => toNumber(b[1]) - toNumber(a[1]));
  const sortedMerchants = Object.entries(merchantTotals).sort((a, b) => toNumber(b[1]) - toNumber(a[1]));
  const healthScore = calculateFinancialHealthScore(transactions, modeConfig);
  const insights = [
    `מצב ${modeConfig.label}: ${modeConfig.focus}. רמת התראות: ${modeConfig.notificationTone}.`,
    `ציון בריאות הוצאות אשראי לפי מצב ${modeConfig.label}: ${healthScore}/100.`,
  ];

  if (typeof context.savingsRate === 'number') {
    const gap = modeConfig.savingsTarget - context.savingsRate;
    if (gap > 0) insights.push(`שיעור החיסכון נמוך מהיעד של מצב ${modeConfig.label} ב־${formatPercent(gap)}.`);
    else insights.push(`שיעור החיסכון עומד ביעד של מצב ${modeConfig.label} ואף גבוה ממנו ב־${formatPercent(Math.abs(gap))}.`);
  }

  const [topCategory, topCategoryAmount] = sortedCategories[0] || [];
  if (topCategory) insights.push(`הקטגוריה הגדולה ביותר באשראי היא ${topCategory}: ${SHEKEL.format(topCategoryAmount)}, שהם ${Math.round((topCategoryAmount / total) * 100)}% מהחיובים.`);

  const [topMerchant, topMerchantAmount] = sortedMerchants[0] || [];
  if (topMerchant) insights.push(`בית העסק הדומיננטי ביותר הוא ${topMerchant}: ${SHEKEL.format(topMerchantAmount)}.`);

  insights.push(`גובה עסקת אשראי ממוצעת: ${SHEKEL.format(total / transactions.length)}.`);
  if (totalIncome > 0) insights.push(`חיובי האשראי הם ${formatPercent((total / totalIncome) * 100)} מההכנסה שהוזנה החודש.`);

  Object.entries(CATEGORY_BUDGETS).forEach(([category, budget]) => {
    const spent = categoryTotals[category] || 0;
    const adjustedBudget = budget / modeConfig.strictness;
    const warningPoint = adjustedBudget * (modeConfig.budgetWarningAt / 100);
    if (spent > adjustedBudget) insights.push(`${category} חרגה מהתקציב המותאם למצב ${modeConfig.label} ב־${SHEKEL.format(spent - adjustedBudget)}.`);
    else if (spent >= warningPoint) insights.push(`${category} מתקרבת לתקציב לפי מצב ${modeConfig.label}: ${SHEKEL.format(spent)} מתוך ${SHEKEL.format(adjustedBudget)}.`);
  });

  const uncategorized = categoryTotals['אחר'] || 0;
  if (uncategorized > 0) insights.push(`${SHEKEL.format(uncategorized)} עדיין מסווגים כ״אחר״. שינוי ידני של קטגוריה ילמד את המערכת לפעמים הבאות.`);

  const largeTransactions = transactions
    .filter((transaction) => toNumber(transaction.amount) >= Math.max(500, total * 0.08))
    .sort((a, b) => toNumber(b.amount) - toNumber(a.amount));
  if (largeTransactions.length > 0) insights.push(`זוהו ${largeTransactions.length} עסקאות גדולות יחסית. הגדולה ביותר: ${largeTransactions[0].merchant} בסך ${SHEKEL.format(largeTransactions[0].amount)}.`);

  if (recurringTransactions.length > 0) {
    const recurringTotal = recurringTransactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
    insights.push(`זוהו ${recurringTransactions.length} עסקאות חוזרות/מנויים בסך כולל של ${SHEKEL.format(recurringTotal)}.`);
  }

  if (modeConfig.priorityMetric === 'burnRate' && context.burnRate) insights.push(`במצב Survival כדאי להוריד Burn Rate מתחת ל־${SHEKEL.format(context.burnRate * 0.9)} בחודש הבא.`);
  if (modeConfig.priorityMetric === 'cashFlow' && context.cashFlow) insights.push(`במצב Growth הדגש הוא להגדיל Cash Flow פנוי מעל ${SHEKEL.format(context.cashFlow + 1000)}.`);
  if (modeConfig.priorityMetric === 'netWorth' && context.totalAssets) insights.push(`במצב Wealth Building הדגש הוא להגדיל שווי שהוזן מעבר ל־${SHEKEL.format(context.totalAssets * 1.05)}.`);

  return insights;
}

// Loads one household state row from Supabase. The optional accessToken is used later for authenticated mode.
async function loadFinanceStateFromSupabase(profileId = DEFAULT_SUPABASE_PROFILE_ID, config = {}) {
  const supabaseUrl = config.url || SUPABASE_URL;
  const supabaseKey = config.key || SUPABASE_ANON_KEY;
  if (!supabaseUrl || !supabaseKey) return null;
  const safeProfileId = profileId || DEFAULT_SUPABASE_PROFILE_ID;
  const url = `${supabaseUrl}/rest/v1/finance_app_state?profile_id=eq.${encodeURIComponent(safeProfileId)}&select=months,learned_rules,global_preferences`;
  const response = await fetch(url, { headers: { apikey: supabaseKey, Authorization: `Bearer ${config.accessToken || supabaseKey}` } });
  if (!response.ok) {
    const details = await response.text().catch(() => '');
    throw new Error(`Supabase load failed: ${response.status} ${details}`);
  }
  const rows = await response.json();
  return Array.isArray(rows) ? rows[0] || null : null;
}

// Upserts the full app state into Supabase by profile_id so one household keeps one cloud row.
function hasMeaningfulMonths(months) {
  if (!months || typeof months !== 'object') return false;

  return Object.values(months).some((data) => {
    const month = normalizeMonthData(data);

    return (
      month.incomes.some((item) => toNumber(item.amount) > 0) ||
      month.manualExpenses.some((item) => toNumber(item.amount) > 0) ||
      month.creditCards.some((card) =>
        (card.transactions || []).length > 0 ||
        (card.pendingTransactions || []).length > 0 ||
        Boolean(card.importedFile)
      ) ||
      month.bankAccounts.some((account) =>
        toNumber(account.openingBalance) !== 0 ||
        toNumber(account.closingBalance) !== 0 ||
        (account.transactions || []).length > 0 ||
        Boolean(account.importedFile)
      ) ||
      month.savingsProducts.some((item) =>
        toNumber(item.currentBalance) > 0 ||
        toNumber(item.monthlyDeposit) > 0
      ) ||
      month.savingGoals.some((item) =>
        toNumber(item.currentAmount) > 0 ||
        toNumber(item.monthlyDeposit) > 0
      ) ||
      toNumber(month.emergencyFund) > 0 ||
      (month.attachedDocuments || []).length > 0
    );
  });
}

async function saveFinanceStateToSupabase(months, learnedRules, globalPreferences, profileId = DEFAULT_SUPABASE_PROFILE_ID, config = {}) {
  const supabaseUrl = config.url || SUPABASE_URL;
  const supabaseKey = config.key || SUPABASE_ANON_KEY;
  if (!supabaseUrl || !supabaseKey) return;

  const existingState = await loadFinanceStateFromSupabase(profileId, config).catch(() => null);

  const incomingHasData = hasMeaningfulMonths(months);
  const cloudHasData = hasMeaningfulMonths(existingState?.months);

  const safeMonths = !incomingHasData && cloudHasData
    ? existingState.months
    : months;

  const safePreferences = mergeCloudPreferences(
    existingState?.global_preferences,
    globalPreferences
  );

  const response = await fetch(`${supabaseUrl}/rest/v1/finance_app_state?on_conflict=profile_id`, {
    method: 'POST',
    headers: {
      apikey: supabaseKey,
      Authorization: `Bearer ${config.accessToken || supabaseKey}`,
      'Content-Type': 'application/json',
      Prefer: 'resolution=merge-duplicates',
    },
    body: JSON.stringify({
      profile_id: profileId || DEFAULT_SUPABASE_PROFILE_ID,
      months: safeMonths,
      learned_rules: learnedRules,
      global_preferences: normalizePreferences(safePreferences),
      updated_at: new Date().toISOString(),
    }),
  });

  if (!response.ok) {
    const details = await response.text().catch(() => '');
    throw new Error(`Supabase save failed: ${response.status} ${details}`);
  }
}

// Global preferences must live above months, otherwise changing month would reset Supabase and dashboard settings.
function createDefaultPreferences() {
  return {
    primaryPerson: 'נועה',
    householdProfileId: DEFAULT_SUPABASE_PROFILE_ID,
    secondaryPerson: 'אורן',
    includeSelfEmployed: false,
    monthlyBudgetTarget: 0,
    savingsRateTarget: 20,
    showMonthlyStory: true,
    showFinancialHealth: true,
    showCategoryChart: true,
    showTrendChart: true,
    showSmartInsightCards: true,
    showRecurringDetection: true,
    themeMood: 'Sage',
    financialMode: 'Stable',
    syncMode: 'Cloud Sync',
    supabaseUrl: DEFAULT_SUPABASE_URL,
    supabaseAnonKey: DEFAULT_SUPABASE_ANON_KEY,
    notifications: { budget80: true, woltSpike: true, savingsDrop: true },
  };
}

// Month defaults contain financial records. The preferences field remains for backward compatibility with older saved data.
function createDefaultMonth() {
  return {
    dashboardTitle: 'מערכת פיננסית משפחתית',
    emergencyFund: 0,
    lastSalaryImport: '',
    attachedDocuments: [],
    bankAccounts: [
      { id: makeId('bank'), name: 'עו״ש משותף', owner: 'משפחה', openingBalance: 0, closingBalance: 0, importedFile: '', transactions: [] },
      { id: makeId('bank'), name: 'עו״ש נועה', owner: 'נועה', openingBalance: 0, closingBalance: 0, importedFile: '', transactions: [] },
      { id: makeId('bank'), name: 'עו״ש אורן', owner: 'אורן', openingBalance: 0, closingBalance: 0, importedFile: '', transactions: [] },
    ],
    incomes: [
      { id: makeId('income'), name: 'משכורת נועה', amount: 0 },
      { id: makeId('income'), name: 'הכנסה אורן', amount: 0 },
      { id: makeId('income'), name: 'הכנסה נוספת', amount: 0 },
    ],
    manualExpenses: [
      { id: makeId('expense'), category: 'משכנתא / שכירות', amount: 0, type: 'קבועה' },
      { id: makeId('expense'), category: 'ארנונה', amount: 0, type: 'קבועה' },
      { id: makeId('expense'), category: 'חשמל', amount: 0, type: 'קבועה' },
      { id: makeId('expense'), category: 'מים', amount: 0, type: 'קבועה' },
      { id: makeId('expense'), category: 'אינטרנט + סלולר', amount: 0, type: 'קבועה' },
      { id: makeId('expense'), category: 'ביטוחים', amount: 0, type: 'קבועה' },
    ],
    savingsProducts: [
      { id: makeId('saving'), name: 'קרן השתלמות נועה', type: 'קרן השתלמות', owner: 'נועה', monthlyDeposit: 0, currentBalance: 0 },
      { id: makeId('saving'), name: 'קרן השתלמות אורן', type: 'קרן השתלמות', owner: 'אורן', monthlyDeposit: 0, currentBalance: 0 },
      { id: makeId('saving'), name: 'פנסיה נועה', type: 'פנסיה', owner: 'נועה', monthlyDeposit: 0, currentBalance: 0 },
      { id: makeId('saving'), name: 'פנסיה אורן', type: 'פנסיה', owner: 'אורן', monthlyDeposit: 0, currentBalance: 0 },
    ],
    savingGoals: [
      { id: makeId('goal'), name: 'טיסה ליפן', targetAmount: 30000, currentAmount: 0, monthlyDeposit: 0 },
      { id: makeId('goal'), name: 'חתונה', targetAmount: 100000, currentAmount: 0, monthlyDeposit: 0 },
      { id: makeId('goal'), name: 'קרן חירום', targetAmount: 60000, currentAmount: 0, monthlyDeposit: 0 },
    ],
    creditCards: [
      { id: makeId('card'), name: 'כרטיס אשראי נועה', owner: 'נועה', importedFile: '', transactions: [], pendingTransactions: [] },
      { id: makeId('card'), name: 'כרטיס אשראי אורן', owner: 'אורן', importedFile: '', transactions: [], pendingTransactions: [] },
    ],
    selfEmployed: {
      owner: 'אורן',
      salaryTransferToHousehold: 0,
      grossRevenue: 0,
      vatCollected: 0,
      vatPaidOnExpenses: 0,
      incomeTaxAdvance: 0,
      nationalInsurance: 0,
      businessExpenses: 0,
    },
    preferences: createDefaultPreferences(),
  };
}

// Monthly data is normalized defensively because old saved months may miss newly added fields.
// Normalizes older saved months so missing arrays/objects never break rendering after schema changes.
function normalizeMonthData(data) {
  const base = createDefaultMonth();
  const safe = data || {};
  return {
    ...base,
    ...safe,
    bankAccounts: (Array.isArray(safe.bankAccounts) ? safe.bankAccounts : base.bankAccounts).map((account) => ({ importedFile: '', transactions: [], ...account })),
    incomes: Array.isArray(safe.incomes) ? safe.incomes : base.incomes,
    manualExpenses: Array.isArray(safe.manualExpenses) ? safe.manualExpenses : base.manualExpenses,
    savingsProducts: Array.isArray(safe.savingsProducts) ? safe.savingsProducts : base.savingsProducts,
    savingGoals: Array.isArray(safe.savingGoals) ? safe.savingGoals : base.savingGoals,
    creditCards: (Array.isArray(safe.creditCards) ? safe.creditCards : base.creditCards).map((card) => ({ transactions: [], pendingTransactions: [], importedFile: '', ...card })),
    attachedDocuments: Array.isArray(safe.attachedDocuments) ? safe.attachedDocuments : base.attachedDocuments,
    selfEmployed: { ...base.selfEmployed, ...(safe.selfEmployed || {}) },
    preferences: normalizePreferences(safe.preferences),
  };
}

function getInitialMonths() {
  const currentMonth = getCurrentMonthKey();
  const saved = safeJsonParse(getStorageItem(STORAGE_KEY), null);
  if (saved && typeof saved === 'object' && !Array.isArray(saved)) {
    if (saved[currentMonth]) return saved;
    const previousMonthKey = getPreviousMonthKey(currentMonth);
    const previousMonthData = saved[previousMonthKey] || Object.entries(saved).filter(([key]) => key < currentMonth).sort(([a], [b]) => b.localeCompare(a))[0]?.[1];
    return { ...saved, [currentMonth]: createRollingMonthFromPrevious(previousMonthData) };
  }
  return { [currentMonth]: createDefaultMonth() };
}

function getInitialLearnedRules() {
  const saved = safeJsonParse(getStorageItem(`${STORAGE_KEY}-rules`), {});
  return saved && typeof saved === 'object' && !Array.isArray(saved) ? saved : {};
}

// Normalizes global settings and preserves nested notification defaults.
function normalizePreferences(preferences) {
  const base = createDefaultPreferences();
  const safe = preferences && typeof preferences === 'object' && !Array.isArray(preferences) ? preferences : {};
  return {
    ...base,
    ...safe,
    notifications: { ...base.notifications, ...(safe.notifications || {}) },
  };
}

function hasMeaningfulPreferences(preferences) {
  if (!preferences || typeof preferences !== 'object' || Array.isArray(preferences)) return false;
  return Object.keys(preferences).some((key) => preferences[key] !== undefined && preferences[key] !== null && preferences[key] !== '');
}

// Cloud preferences are merged, not blindly applied, so an empty cloud row cannot erase a working local Supabase setup.
// Merges cloud settings carefully so an empty cloud row cannot wipe a working local Supabase configuration.
function mergeCloudPreferences(current, cloud) {
  const currentSafe = normalizePreferences(current);
  if (!hasMeaningfulPreferences(cloud)) return currentSafe;
  const cloudSafe = normalizePreferences(cloud);
  return normalizePreferences({
    ...currentSafe,
    ...cloudSafe,
    supabaseUrl: cloud.supabaseUrl || currentSafe.supabaseUrl,
    supabaseAnonKey: cloud.supabaseAnonKey || currentSafe.supabaseAnonKey,
    householdProfileId: cloud.householdProfileId || currentSafe.householdProfileId,
  });
}

function getInitialGlobalPreferences() {
  const saved = safeJsonParse(getStorageItem(SETTINGS_STORAGE_KEY), null);
  return normalizePreferences(saved);
}

function getInitialAuthSession() {
  const saved = safeJsonParse(getStorageItem(AUTH_STORAGE_KEY), null);
  if (!saved || typeof saved !== 'object') return null;
  if (!saved.access_token || !saved.email) return null;
  if (saved.expires_at && Date.now() > saved.expires_at) return null;
  return saved;
}

function clearAuthSession() {
  setStorageItem(AUTH_STORAGE_KEY, '');
}

// Password login for Supabase Auth.
async function signInWithSupabasePassword(email, password, config = {}) {
  const supabaseUrl = config.url || SUPABASE_URL;
  const supabaseKey = config.key || SUPABASE_ANON_KEY;

  if (!supabaseUrl || !supabaseKey) throw new Error('חסרים Supabase URL או Publishable Key');
  if (!email || !password) throw new Error('צריך להזין אימייל וסיסמה');

  const response = await fetch(`${supabaseUrl}/auth/v1/token?grant_type=password`, {
    method: 'POST',
    headers: {
      apikey: supabaseKey,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email, password }),
  });

  if (!response.ok) {
    const details = await response.text().catch(() => '');
    throw new Error(`הכניסה נכשלה: ${response.status} ${details}`);
  }

  const data = await response.json();
  const session = {
    email,
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    user_id: data.user?.id || '',
    expires_at: Date.now() + Math.max(1, Number(data.expires_in || 3600) - 60) * 1000,
  };

  setStorageItem(AUTH_STORAGE_KEY, JSON.stringify(session));
  return session;
}

async function signUpWithSupabasePassword(email, password, config = {}) {
  const supabaseUrl = config.url || SUPABASE_URL;
  const supabaseKey = config.key || SUPABASE_ANON_KEY;

  if (!supabaseUrl || !supabaseKey) throw new Error('חסרים Supabase URL או Publishable Key');
  if (!email || !password) throw new Error('צריך להזין אימייל וסיסמה');

  const response = await fetch(`${supabaseUrl}/auth/v1/signup`, {
    method: 'POST',
    headers: {
      apikey: supabaseKey,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ email, password }),
  });

  if (!response.ok) {
    const details = await response.text().catch(() => '');
    throw new Error(`ההרשמה נכשלה: ${response.status} ${details}`);
  }

  return signInWithSupabasePassword(email, password, config);
}
// Lightweight runtime checks catch common parser/calculation regressions while developing in the browser canvas.
function runSmokeTests() {
  console.assert(APP_BUILD_MARKER === 'finance-dashboard-build-v14', 'build marker failed');
  console.assert(getPublicEnv('THIS_ENV_SHOULD_NOT_EXIST') === '', 'safe env fallback failed');
  console.assert(toNumber('₪1,250') === 1250, 'currency parsing failed');
  console.assert(detectCategory('Wolt TLV') === 'מסעדות ובתי קפה', 'wolt category failed');
  console.assert(detectCategory('My Shop', { shop: 'קניות' }) === 'קניות', 'learned rule failed');
  console.assert(splitCsvLine('a,b,c').length === 3, 'csv split failed');
  console.assert(parseCsvText(['date,merchant,amount', '2026-01-01,Wolt,55'].join(String.fromCharCode(10))).length === 1, 'csv parse failed');
  console.assert(parseCsvText(['date,merchant,amount', '2026-01-01,Wolt,55'].join(String.fromCharCode(13) + String.fromCharCode(10))).length === 1, 'csv CRLF parse failed');
  console.assert(splitCsvLine('"a,b",c').length === 2, 'quoted csv parsing failed');
  console.assert(getCategoryTotals([{ category: 'קניות', amount: 10 }, { category: 'קניות', amount: 20 }]).קניות === 30, 'category totals failed');
  console.assert(buildRealInsights([{ merchant: 'Wolt', category: 'מסעדות ובתי קפה', amount: 900 }], [], 0, 'Survival').some((insight) => insight.includes('Survival')), 'real budget insight failed');
  console.assert(getFinancialModeConfig('Growth').savingsTarget === 25, 'financial mode config failed');
  console.assert(detectRecurringTransactions([{ merchant: 'Netflix', amount: 50 }]).length === 1, 'recurring detection failed');
  console.assert(normalizeMonthData({}).creditCards.length === 2, 'month normalizer failed');
  console.assert(normalizeMonthData({ creditCards: null }).creditCards.length === 2, 'month normalizer array fallback failed');
  console.assert(getMonthTotals(undefined).expenses === 0, 'month totals fallback failed');
  console.assert(getMonthTotals({ selfEmployed: { incomeTaxAdvance: 100 }, preferences: { includeSelfEmployed: false } }).selfEmployed === 0, 'self employed disabled failed');
  console.assert(getMonthTotals({ selfEmployed: { incomeTaxAdvance: 100 }, preferences: { includeSelfEmployed: true } }).selfEmployed === 100, 'self employed enabled failed');
  console.assert(Array.isArray(normalizeMonthData({}).attachedDocuments), 'attached documents normalizer failed');
  console.assert(getMonthlyTrend({ '2026-01': createDefaultMonth() }).length === 1, 'monthly trend failed');
  console.assert(getPreviousMonthKey('2026-03') === '2026-02', 'previous month helper failed');
  console.assert(getMonthlyCompare({ '2026-01': createDefaultMonth(), '2026-02': createDefaultMonth() }, '2026-02').hasPrevious === true, 'monthly compare failed');
  console.assert(getCompareMonthKeys({ '2026-01': {}, '2026-02': {}, '2026-03': {}, '2026-04': {} }, '2026-04', 'quarter').length === 3, 'quarter compare failed');
  console.assert(getMonthlyCompare({ '2026-01': createDefaultMonth(), '2026-02': createDefaultMonth(), '2026-03': createDefaultMonth() }, '2026-03', 'all').compareMonthKeys.length === 2, 'all period compare failed');
  console.assert(parseExcelArrayBuffer instanceof Function, 'excel parser exists');
  console.assert(TABS[1].id === 'income', 'income tab should be second');
  console.assert(getCurrentMonthKey().length === 7 && getCurrentMonthKey().includes('-'), 'current month key failed');
  console.assert(TABS.some((tab) => tab.id === 'insights' && tab.label === 'תובנות חכמות'), 'smart insights tab label failed');
  console.assert(normalizePreferences({ showTrendChart: false }).showTrendChart === false, 'preferences override failed');
  console.assert(normalizePreferences({}).householdProfileId === DEFAULT_SUPABASE_PROFILE_ID, 'household profile default failed');
  console.assert(getSafeTheme('Missing').accent === THEME_STYLES.Sage.accent, 'theme fallback failed');
  console.assert(getSafeTheme('Dark').page.includes('111111'), 'dark theme page exists');
  console.assert(noSingleWordLine('אחת שתיים שלוש').includes(String.fromCharCode(160)), 'no orphan text helper failed');
  console.assert(getInitialLearnedRules() && typeof getInitialLearnedRules() === 'object', 'initial learned rules failed');
}

if (typeof window !== 'undefined') runSmokeTests();

function StatCard({ title, value, note, tone = 'neutral', help = '' }) {
  const [showHelp, setShowHelp] = useState(false);
  const toneClass = {
    neutral: 'border-neutral-200 bg-white',
    good: 'border-neutral-200 bg-white',
    warn: 'border-amber-200 bg-amber-50',
    danger: 'border-red-200 bg-red-50',
  }[tone] || 'border-neutral-200 bg-white';
  const noteClass = {
    neutral: 'text-neutral-500',
    good: 'text-[#6F7D65]',
    warn: 'text-amber-700',
    danger: 'text-red-700',
  }[tone] || 'text-neutral-500';

  return (
    <div className={`min-h-[150px] rounded-[22px] border ${toneClass} p-4 shadow-sm transition hover:shadow-md sm:min-h-[180px] sm:p-5 lg:min-h-[140px]`}>
      <div className="flex items-center justify-center gap-2 text-center text-xs font-semibold uppercase tracking-widest text-neutral-400">
  <span>{title}</span>
  {help ? (
    <button
  type="button"
  onClick={() => setShowHelp((current) => !current)}
  className="inline-flex h-5 w-5 items-center justify-center rounded-full border border-neutral-300 text-[11px] text-neutral-500"
>
  ?
</button>
  ) : null}
</div>
      {showHelp && help ? (
  <div className="mt-3 rounded-xl border border-neutral-200 bg-neutral-50 p-3 text-center text-xs leading-6 text-neutral-600">
    {help}
  </div>
) : null}
      <div className="mt-4 text-center text-2xl font-semibold tracking-tight text-neutral-950 sm:mt-6 sm:text-3xl">{value}</div>
      <div className={`mt-4 px-1 text-center text-xs font-medium leading-6 sm:mt-6 sm:px-2 sm:text-sm sm:leading-7 ${noteClass} no-single-word-lines`}>{noSingleWordLine(note)}</div>
    </div>
  );
}

function Section({ children, className = '' }) {
  return <section className={`rounded-[24px] border border-neutral-200 bg-white p-5 shadow-sm sm:p-8 ${className}`}>{children}</section>;
}

function EmptyState({ title, text, action }) {
  return (
    <div className="rounded-[24px] border border-dashed border-neutral-300 bg-neutral-50 p-10 text-center">
      <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-2xl bg-white text-lg font-semibold text-neutral-500">＋</div>
      <h3 className="mt-4 text-lg font-semibold text-neutral-900">{title}</h3>
      <p className="mx-auto mt-2 max-w-md text-sm leading-7 text-neutral-500 no-orphans">{noSingleWordLine(text)}</p>
      {action ? <div className="mt-5">{action}</div> : null}
    </div>
  );
}

function PrimaryButton({ children, className = '', theme = THEME_STYLES.Sage, ...props }) {
  const accent = theme?.accent || THEME_STYLES.Sage.accent;
  const accentHover = theme?.accentHover || THEME_STYLES.Sage.accentHover;
  return (
    <button
      {...props}
      className={`rounded-lg px-4 py-2 text-sm font-semibold text-white transition disabled:opacity-50 ${className}`}
      style={{ backgroundColor: accent }}
      onMouseEnter={(event) => { event.currentTarget.style.backgroundColor = accentHover; }}
      onMouseLeave={(event) => { event.currentTarget.style.backgroundColor = accent; }}
    >
      {children}
    </button>
  );
}

function GhostButton({ children, className = '', ...props }) {
  return <button {...props} className={`rounded-lg border border-neutral-200 bg-white px-3 py-2 text-sm font-semibold text-neutral-700 transition hover:border-neutral-400 hover:bg-neutral-50 ${className}`}>{children}</button>;
}

function Field({ className = '', ...props }) {
  return <input {...props} className={`rounded-lg border border-neutral-300 bg-white px-3 py-2 text-sm outline-none transition focus:border-neutral-900 focus:ring-2 focus:ring-neutral-100 ${className}`} />;
}

function SelectField({ children, className = '', ...props }) {
  return <select {...props} className={`rounded-lg border border-neutral-300 bg-white px-3 py-2 text-sm outline-none transition focus:border-neutral-900 focus:ring-2 focus:ring-neutral-100 ${className}`}>{children}</select>;
}

function LabeledField({ label, children }) {
  return (
    <label className="grid gap-2 text-xs font-semibold text-neutral-500">
      <span>{label}</span>
      {children}
    </label>
  );
}

function InputRow({ children }) {
  return <div className="grid gap-3 rounded-[24px] border border-neutral-200 p-4 md:grid-cols-[1fr_160px_44px]">{children}</div>;
}

function TrendLineChart({ data, theme }) {
  const chartData = Array.isArray(data) ? data : [];
  const width = 720;
  const height = 220;
  const padding = { top: 28, right: 36, bottom: 44, left: 72 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const values = chartData.flatMap((item) => [toNumber(item.income), toNumber(item.expenses), toNumber(item.savings)]);
  const minValue = Math.min(0, ...values);
  const maxValue = Math.max(1, ...values);
  const range = maxValue - minValue || 1;
  const xFor = (index) => padding.left + (chartData.length <= 1 ? innerWidth / 2 : (index / (chartData.length - 1)) * innerWidth);
  const yFor = (value) => padding.top + innerHeight - ((toNumber(value) - minValue) / range) * innerHeight;
  const pathFor = (key) => chartData.map((item, index) => `${index === 0 ? 'M' : 'L'} ${xFor(index)} ${yFor(item[key])}`).join(' ');
  const zeroY = yFor(0);
  const ticks = [0, 0.25, 0.5, 0.75, 1].map((ratio) => minValue + range * ratio);

  if (!chartData.length) {
    return <EmptyState title="אין עדיין נתונים לגרף" text="כדי לראות מגמות, מלאי לפחות חודש אחד של הכנסות והוצאות." />;
  }

  return (
    <div className="w-full overflow-hidden rounded-[22px] border border-neutral-200 bg-neutral-50 p-3 sm:rounded-[24px] sm:p-4">
      <svg viewBox={`0 0 ${width} ${height}`} className="w-full" role="img" aria-label="גרף מגמות הכנסות הוצאות וחיסכון">
        <rect x="0" y="0" width={width} height={height} rx="24" fill="white" />
        {ticks.map((tick) => {
          const y = yFor(tick);
          return (
            <g key={tick}>
              <line x1={padding.left} x2={width - padding.right} y1={y} y2={y} stroke="#E5E5E5" strokeWidth="1" />
              <text x={padding.left - 12} y={y + 4} textAnchor="end" fontSize="12" fill="#737373">{SHEKEL.format(tick)}</text>
            </g>
          );
        })}
        <line x1={padding.left} x2={width - padding.right} y1={zeroY} y2={zeroY} stroke="#D4D4D4" strokeWidth="1.5" strokeDasharray="4 6" />

        <path d={pathFor('income')} fill="none" stroke={theme?.accent || '#66725E'} strokeWidth="4" strokeLinecap="round" strokeLinejoin="round" />
        <path d={pathFor('expenses')} fill="none" stroke="#D97706" strokeWidth="4" strokeLinecap="round" strokeLinejoin="round" />
        <path d={pathFor('savings')} fill="none" stroke="#2563EB" strokeWidth="4" strokeLinecap="round" strokeLinejoin="round" />

        {chartData.map((item, index) => (
          <g key={item.month}>
            <text x={xFor(index)} y={height - 18} textAnchor="middle" fontSize="12" fill="#737373">{monthLabel(item.month)}</text>
            <circle cx={xFor(index)} cy={yFor(item.income)} r="5" fill={theme?.accent || '#66725E'} />
            <circle cx={xFor(index)} cy={yFor(item.expenses)} r="5" fill="#D97706" />
            <circle cx={xFor(index)} cy={yFor(item.savings)} r="5" fill="#2563EB" />
          </g>
        ))}
      </svg>
      <div className="mt-4 flex flex-wrap gap-3 text-sm font-semibold text-neutral-600">
        <span className="rounded-full bg-white px-3 py-2">● הכנסות</span>
        <span className="rounded-full bg-white px-3 py-2 text-amber-700">● הוצאות</span>
        <span className="rounded-full bg-white px-3 py-2 text-blue-700">● חיסכון</span>
      </div>
    </div>
  );
}

function TransactionEditorTable({ rows, cardId, mode, onUpdate, onRemove }) {
  const isPending = mode === 'pending';

  function formatTransactionAmount(transaction) {
    if (transaction.currency && transaction.currency !== 'ILS') {
      return `${transaction.originalAmount || transaction.amount} ${transaction.currency}`;
    }
    return SHEKEL.format(transaction.amount);
  }

  return (
    <div className="mt-5 max-h-[900px] overflow-auto rounded-[24px] border border-neutral-200 bg-white">
      <div className="min-w-[860px]">
        <div className="md:sticky md:top-0 z-10 grid grid-cols-[110px_minmax(180px,1fr)_150px_130px_160px_52px] bg-neutral-100 px-5 py-4 text-sm font-semibold text-neutral-700">
          <div>תאריך</div>
          <div>{isPending ? 'בית עסק' : 'עסקה'}</div>
          <div>{isPending ? 'קטגוריה' : 'קטגוריה לומדת'}</div>
          <div>סכום</div>
          <div>סוג הוצאה</div>
          <div />
        </div>

        {rows.map((transaction) => (
          <div key={transaction.id} className="grid grid-cols-[110px_minmax(180px,1fr)_150px_130px_160px_52px] gap-4 border-t border-neutral-100 p-4">
            {isPending ? <Field value={transaction.date || ''} onChange={(event) => onUpdate(cardId, transaction.id, 'date', event.target.value)} /> : <div className="px-3 py-3 text-sm text-neutral-500">{transaction.date}</div>}
            {isPending ? <Field value={transaction.merchant} onChange={(event) => onUpdate(cardId, transaction.id, 'merchant', event.target.value)} /> : <Field value={transaction.merchant} readOnly className="bg-neutral-50" />}

            <SelectField value={transaction.category} onChange={(event) => isPending ? onUpdate(cardId, transaction.id, 'category', event.target.value) : onUpdate(transaction.id, event.target.value)}>
              {EXPENSE_CATEGORIES.map((category) => <option key={category}>{category}</option>)}
            </SelectField>

            {isPending ? (
              <Field type="number" value={transaction.amount} onChange={(event) => onUpdate(cardId, transaction.id, 'amount', event.target.value)} className="w-full text-left" />
            ) : (
              <div className="px-3 py-3 text-left text-sm font-semibold text-neutral-900">{formatTransactionAmount(transaction)}</div>
            )}

            <SelectField value={transaction.necessity || 'מותרות'} onChange={(event) => isPending ? onUpdate(cardId, transaction.id, 'necessity', event.target.value) : onUpdate(transaction.id, event.target.value)}>
              {['חיוני', 'חשוב', 'מותרות'].map((item) => <option key={item}>{item}</option>)}
            </SelectField>

            <GhostButton onClick={() => onRemove(cardId, transaction.id)} className="px-0">×</GhostButton>
          </div>
        ))}

        {rows.length === 0 ? <div className="p-10 text-center text-sm text-neutral-400">עדיין לא העלית פירוט אשראי. העלי CSV או Excel כדי להתחיל ניתוח חכם של ההוצאות.</div> : null}
      </div>
    </div>
  );
}

function CreditCardPanel(props) {
  const {
    card,
    cardTotal,
    onUpdateCard,
    onRemoveCard,
    onImportFile,
    onUpdatePending,
    onRemovePending,
    onApprovePending,
    onAddTransaction,
    onRemoveTransaction,
    onUpdateCategory,
    theme,
  } = props;

  const safeTheme = theme || THEME_STYLES.Sage;
  const pendingRows = card.pendingTransactions || [];
  const approvedRows = card.transactions || [];

  return (
    <div className="rounded-[28px] border border-neutral-200 bg-white p-6 shadow-sm">
      <div className="grid gap-3 md:grid-cols-[minmax(180px,1fr)_minmax(140px,1fr)_110px_40px]">
        <Field value={card.name} onChange={(event) => onUpdateCard(card.id, 'name', event.target.value)} placeholder="שם הכרטיס" />
        <Field value={card.owner} onChange={(event) => onUpdateCard(card.id, 'owner', event.target.value)} placeholder="בעל/ת הכרטיס" />
        <div className="rounded-xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-sm font-semibold text-neutral-800">{SHEKEL.format(cardTotal)}</div>
        <GhostButton onClick={() => onRemoveCard(card.id)} className="px-0">×</GhostButton>
      </div>

      <div className="mt-5 rounded-[24px] border border-dashed border-neutral-300 bg-neutral-50 p-4">
        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <div>
            <div className="text-sm font-semibold text-neutral-900">העלאת CSV / Excel של פירוט אשראי</div>
            <div className="mt-1 text-xs text-neutral-500">העלאה נמצאת כאן, בתוך הכרטיס הרלוונטי.</div>
          </div>
          <label className="cursor-pointer rounded-xl px-4 py-3 text-sm font-semibold text-white transition" style={{ backgroundColor: safeTheme.accent }}>
            העלאת קובץ
            <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) onImportFile(card.id, file); }} />
          </label>
        </div>
        {card.importedFile ? <div className="mt-3 rounded-xl bg-white px-4 py-3 text-sm text-neutral-600">נקלט קובץ: <strong>{card.importedFile}</strong></div> : null}
      </div>

      {pendingRows.length > 0 ? (
        <div className="mt-4 rounded-[24px] border border-neutral-200 bg-neutral-50 p-4">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <div className="text-sm font-semibold text-neutral-900">עסקאות שזוהו לאישור</div>
              <div className="mt-1 text-xs text-neutral-500">בדקו סכומים וקטגוריות לפני שהן נכנסות להוצאות.</div>
            </div>
            <PrimaryButton theme={safeTheme} onClick={() => onApprovePending(card.id)}>אשר והכנס להוצאות</PrimaryButton>
          </div>
          <TransactionEditorTable rows={pendingRows} cardId={card.id} mode="pending" onUpdate={onUpdatePending} onRemove={onRemovePending} />
        </div>
      ) : null}

      <TransactionEditorTable rows={approvedRows} cardId={card.id} mode="approved" onUpdate={onUpdateCategory} onRemove={onRemoveTransaction} />
      <PrimaryButton theme={safeTheme} onClick={() => onAddTransaction(card.id)} className="mt-4">+ הוספת עסקה</PrimaryButton>
    </div>
  );
}

export default function PersonalIsraeliFamilyFinanceDashboard() {
  // selectedMonth switches the active month, while globalPreferences remains shared across all months.
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonthKey());
  const [months, setMonths] = useState(getInitialMonths);
  const [learnedRules, setLearnedRules] = useState(getInitialLearnedRules);
  const [globalPreferences, setGlobalPreferences] = useState(getInitialGlobalPreferences);
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('הכול');
  const [minAmount, setMinAmount] = useState('');
  const [maxAmount, setMaxAmount] = useState('');
  const [cloudStatus, setCloudStatus] = useState('טוען מהענן…');
  const [hasLoadedCloud, setHasLoadedCloud] = useState(false);
  const [hasAttemptedCloudLoad, setHasAttemptedCloudLoad] = useState(false);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [comparePeriod, setComparePeriod] = useState('previous');
  const [authSession, setAuthSession] = useState(getInitialAuthSession);
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authStatus, setAuthStatus] = useState('');
  const [fxRates, setFxRates] = useState(getStoredFxRates);

  const monthData = normalizeMonthData(months[selectedMonth]);
  const preferences = normalizePreferences(globalPreferences);
  const activeTheme = getSafeTheme(preferences.themeMood);
  const isDark = preferences.themeMood === 'Dark';
  const modeConfig = getFinancialModeConfig(preferences.financialMode);
  const householdProfileId = authSession?.user_id || preferences.householdProfileId || DEFAULT_SUPABASE_PROFILE_ID;
  const supabaseConfig = {
    url: preferences.supabaseUrl || SUPABASE_URL,
    key: preferences.supabaseAnonKey || SUPABASE_ANON_KEY,
    accessToken: authSession?.access_token || '',
  };
  const setupHealth = {
    localStorage: typeof window !== 'undefined',
    supabaseEnv: Boolean(supabaseConfig.url && supabaseConfig.key),
    signedIn: Boolean(authSession?.access_token),
    householdProfileId,
    xlsxParser: Boolean(XLSX && XLSX.read),
  };
  useEffect(() => {
  async function loadFxRates() {
    const rates = await fetchFxRatesToIls();
    setFxRates(rates);
  }

  loadFxRates();
}, []);
  
  // Initial cloud load merges cloud state into local state without wiping missing months or settings.
  useEffect(() => {
    // Cloud load merges Supabase data into local state without deleting the currently selected month.
    async function loadCloudState() {
      if (hasAttemptedCloudLoad) return;
      setHasAttemptedCloudLoad(true);
      try {
        const data = await loadFinanceStateFromSupabase(householdProfileId, supabaseConfig);
        if (data?.months) {
          setMonths((current) => {
            const cloudMonths = data.months && typeof data.months === 'object' && !Array.isArray(data.months) ? data.months : {};
            return mergeMonthsKeepingRicher(current, cloudMonths);
          });
        }
        if (data?.learned_rules) setLearnedRules(data.learned_rules);
        if (data?.global_preferences) {
          setGlobalPreferences((current) => mergeCloudPreferences(current, data.global_preferences));
        }
        setCloudStatus(data?.months ? 'מסונכרן מהענן' : 'אין עדיין נתוני ענן, עובדים מקומית');
      } catch (error) {
        setCloudStatus(`ענן לא זמין: ${error?.message || 'שגיאת Supabase'}`);
      } finally {
        setHasLoadedCloud(true);
      }
    }
    loadCloudState();
  }, [householdProfileId, supabaseConfig.url, supabaseConfig.key, hasAttemptedCloudLoad]);

  // LocalStorage is always updated as a fallback, even when Cloud Sync is enabled.
  useEffect(() => {
    // Always keep a local copy so the dashboard still works if Supabase is unavailable.
    setStorageItem(STORAGE_KEY, JSON.stringify(months));
    setStorageItem(`${STORAGE_KEY}-rules`, JSON.stringify(learnedRules));
  }, [months, learnedRules]);

  useEffect(() => {
    // Global preferences are persisted separately from months to prevent resets when creating a new month.
    setStorageItem(SETTINGS_STORAGE_KEY, JSON.stringify(globalPreferences));
  }, [globalPreferences]);

  // Cloud saving is debounced so typing in a field does not fire a network request on every keystroke.
  useEffect(() => {
    // Debounced cloud save prevents a Supabase write on every keystroke while still autosaving changes.
    if (!hasLoadedCloud) return;
    const saveTimeout = setTimeout(async () => {
      try {
        if (preferences.syncMode === 'Cloud Sync' || preferences.syncMode === 'Auto Backup') {
          await saveFinanceStateToSupabase(months, learnedRules, preferences, householdProfileId, supabaseConfig);
          setCloudStatus(supabaseConfig.url && supabaseConfig.key ? 'נשמר בענן' : 'לא הוגדר Supabase, נשמר מקומית');
        } else {
          setCloudStatus('Local Only: נשמר רק בדפדפן');
        }
      } catch (error) {
        setCloudStatus(`לא נשמר בענן: ${error?.message || 'שגיאת Supabase'}`);
      }
    }, 900);
    return () => clearTimeout(saveTimeout);
  }, [months, learnedRules, globalPreferences, hasLoadedCloud, preferences.syncMode, householdProfileId, supabaseConfig.url, supabaseConfig.key]);

  function setSelectedMonthData(nextData) {
    setMonths((current) => ({ ...current, [selectedMonth]: nextData }));
  }

  function ensureMonth(monthKey) {
    setMonths((current) => {
      const previousMonthKey = getPreviousMonthKey(monthKey);
      const previousMonthData = current[previousMonthKey] || Object.entries(current).filter(([key]) => key < monthKey).sort(([a], [b]) => b.localeCompare(a))[0]?.[1];
      if (current[monthKey]) {
        return { ...current, [monthKey]: carryForwardRollingFields(current[monthKey], previousMonthData) };
      }
      return { ...current, [monthKey]: createRollingMonthFromPrevious(previousMonthData) };
    });
    setSelectedMonth(monthKey);
  }

  function updateMonthField(field, value) {
    setSelectedMonthData({ ...monthData, [field]: value });
  }

  function updateRow(section, id, field, value) {
    const numericFields = ['amount', 'monthlyDeposit', 'currentBalance', 'targetAmount', 'currentAmount', 'openingBalance', 'closingBalance'];
    setSelectedMonthData({
      ...monthData,
      [section]: monthData[section].map((row) => (row.id === id ? { ...row, [field]: numericFields.includes(field) ? toNumber(value) : value } : row)),
    });
  }

  function removeRow(section, id) {
    setSelectedMonthData({ ...monthData, [section]: monthData[section].filter((row) => row.id !== id) });
  }

  function addBankAccount() {
    setSelectedMonthData({ ...monthData, bankAccounts: [...monthData.bankAccounts, { id: makeId('bank'), name: 'חשבון חדש', owner: 'משפחה', openingBalance: 0, closingBalance: 0, importedFile: '', transactions: [] }] });
  }

  async function importBankFile(accountId, file) {
    try {
      const lower = file.name.toLowerCase();
      let importedTransactions = [];
      if (lower.endsWith('.csv')) importedTransactions = parseBankCsvText(await file.text());
      else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) importedTransactions = parseBankExcelArrayBuffer(await file.arrayBuffer());
      else {
        alert('לעו״ש אפשר להעלות CSV או Excel מהבנק.');
        return;
      }

      if (!importedTransactions.length) {
        setCloudStatus('קובץ העו״ש נקלט, אבל לא זוהו תנועות. בדקי שיש עמודות תאריך, פירוט, חובה/זכות או סכום.');
        alert('קובץ העו״ש נקלט, אבל לא זוהו תנועות. אם זה פורמט אחר של הבנק, נצטרך להתאים את העמודות.');
        return;
      }

      setSelectedMonthData({
        ...monthData,
        bankAccounts: monthData.bankAccounts.map((account) => {
          if (account.id !== accountId) return account;
          const balances = getBankBalancesFromTransactions(importedTransactions, account.openingBalance, account.closingBalance);
          return {
            ...account,
            importedFile: file.name,
            transactions: importedTransactions,
            openingBalance: balances.openingBalance,
            closingBalance: balances.closingBalance,
          };
        }),
      });
      setCloudStatus(`יובאו ${importedTransactions.length} תנועות עו״ש.`);
    } catch (error) {
      setCloudStatus(`שגיאה בייבוא עו״ש: ${error?.message || 'לא ידוע'}`);
      alert(`שגיאה בייבוא עו״ש: ${error?.message || 'לא ידוע'}`);
    }
  }

  function addIncome() {
    setSelectedMonthData({ ...monthData, incomes: [...monthData.incomes, { id: makeId('income'), name: 'הכנסה חדשה', amount: 0 }] });
  }

  function addManualExpense() {
    setSelectedMonthData({ ...monthData, manualExpenses: [...monthData.manualExpenses, { id: makeId('expense'), category: 'הוצאה חדשה', amount: 0, type: 'משתנה' }] });
  }

  function addSavingsProduct() {
    setSelectedMonthData({ ...monthData, savingsProducts: [...monthData.savingsProducts, { id: makeId('saving'), name: 'חיסכון חדש', type: 'חיסכון', owner: 'משפחה', monthlyDeposit: 0, currentBalance: 0 }] });
  }

  function addSavingGoal() {
    setSelectedMonthData({ ...monthData, savingGoals: [...monthData.savingGoals, { id: makeId('goal'), name: 'יעד חדש', targetAmount: 0, currentAmount: 0, monthlyDeposit: 0 }] });
  }

  function updateSelfEmployedField(field, value) {
    const numericFields = ['salaryTransferToHousehold', 'grossRevenue', 'vatCollected', 'vatPaidOnExpenses', 'incomeTaxAdvance', 'nationalInsurance', 'businessExpenses'];
    setSelectedMonthData({ ...monthData, selfEmployed: { ...monthData.selfEmployed, [field]: numericFields.includes(field) ? toNumber(value) : value } });
  }

  // Preferences are global, not monthly, so changing months cannot reset connection details or dashboard settings.
  // Updates global settings, not the current month. This keeps Supabase/theme/widgets consistent across months.
  function updatePreference(field, value) {
    const numericFields = ['monthlyBudgetTarget', 'savingsRateTarget'];
    setGlobalPreferences((current) => ({
      ...normalizePreferences(current),
      [field]: numericFields.includes(field) ? toNumber(value) : value,
    }));
  }

  // Salary PDFs are attached for record keeping only. Net salary is still entered manually because there is no OCR parser here.
  function attachSalarySlipFile(file) {
    const nextDocument = {
      id: makeId('doc'),
      name: file.name,
      type: file.type || 'application/pdf',
      size: file.size || 0,
      addedAt: new Date().toISOString(),
    };
    setSelectedMonthData({
      ...monthData,
      lastSalaryImport: file.name,
      attachedDocuments: [...(monthData.attachedDocuments || []), nextDocument],
    });
    setCloudStatus('התלוש צורף לתיעוד. סכום הנטו לא מתפענח אוטומטית, הזיני אותו בשורת ההכנסה הרלוונטית.');
  }

  function removeAttachedDocument(documentId) {
    setSelectedMonthData({ ...monthData, attachedDocuments: (monthData.attachedDocuments || []).filter((document) => document.id !== documentId) });
  }

  async function importIncomeFile(file) {
    try {
      const lower = file.name.toLowerCase();
      let importedIncomes = [];
      if (lower.endsWith('.csv')) importedIncomes = parseIncomeCsvText(await file.text());
      else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) importedIncomes = parseIncomeExcelArrayBuffer(await file.arrayBuffer());
      else if (lower.endsWith('.pdf')) importedIncomes = await parseIncomePdfFile(file);
      else {
        alert('להכנסות אפשר להעלות PDF תלוש, CSV או Excel.');
        return;
      }

      if (!importedIncomes.length) {
        setCloudStatus('הקובץ נקלט, אבל לא זוהתה הכנסה. אם זה PDF סרוק או תלוש דחוס, צריך להזין נטו ידנית.');
        alert('הקובץ נקלט, אבל לא זוהתה הכנסה. אם זה PDF סרוק או תלוש דחוס, צריך להזין נטו ידנית.');
        return;
      }

      setSelectedMonthData({
        ...monthData,
        incomes: [...monthData.incomes, ...importedIncomes],
        lastSalaryImport: file.name,
      });
      setCloudStatus(`יובאו ${importedIncomes.length} שורות הכנסה.`);
    } catch (error) {
      setCloudStatus(`שגיאה בייבוא הכנסות: ${error?.message || 'לא ידוע'}`);
      alert(`שגיאה בייבוא הכנסות: ${error?.message || 'לא ידוע'}`);
    }
  }

  function exportBackup() {
    if (typeof document === 'undefined' || typeof URL === 'undefined') return;
    const backup = { version: 1, exportedAt: new Date().toISOString(), months, learnedRules };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `finance-backup-${selectedMonth}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  async function importBackupFile(file) {
    try {
      const backup = JSON.parse(await file.text());
      if (!backup || !backup.months) throw new Error('Invalid backup');
      setMonths(backup.months);
      setLearnedRules(backup.learnedRules || {});
      setCloudStatus('שוחזר מגיבוי JSON');
    } catch {
      alert('קובץ הגיבוי לא תקין. נא להעלות JSON שיוצא מהמערכת.');
    }
  }

  function resetCurrentMonth() {
    const confirmed = typeof window === 'undefined' ? false : window.confirm('לאפס את החודש הנוכחי? הפעולה תמחק נתונים של החודש בלבד.');
    if (!confirmed) return;
    setSelectedMonthData(createDefaultMonth());
  }

  function addCreditCard() {
    setSelectedMonthData({ ...monthData, creditCards: [...monthData.creditCards, { id: makeId('card'), name: 'כרטיס אשראי חדש', owner: '', importedFile: '', transactions: [], pendingTransactions: [] }] });
  }

  function updateCreditCard(cardId, field, value) {
    setSelectedMonthData({ ...monthData, creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, [field]: value } : card)) });
  }

  function removeCreditCard(cardId) {
    setSelectedMonthData({ ...monthData, creditCards: monthData.creditCards.filter((card) => card.id !== cardId) });
  }

  // Imports CSV/XLS/XLSX into pending transactions first, so users can review before adding them to expenses.
  async function importCreditFile(cardId, file) {
    try {
      const lower = file.name.toLowerCase();
      let importedTransactions = [];
      if (lower.endsWith('.csv')) importedTransactions = parseCsvText(await file.text(), learnedRules, fxRates);
else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) importedTransactions = parseExcelArrayBuffer(await file.arrayBuffer(), learnedRules, fxRates);
      else {
        alert('נא להעלות CSV או Excel');
        return;
      }

      if (!importedTransactions.length) {
        setCloudStatus('הקובץ נקלט, אבל לא זוהו עסקאות. בדקי שיש עמודות תאריך, בית עסק וסכום.');
        alert('הקובץ נקלט, אבל לא זוהו עסקאות. אם זה פירוט בנק/אשראי בפורמט אחר, נצטרך להתאים את מבנה העמודות.');
      } else {
        setCloudStatus(`זוהו ${importedTransactions.length} עסקאות לאישור בכרטיס האשראי.`);
      }

      setSelectedMonthData({
        ...monthData,
        creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, importedFile: file.name, pendingTransactions: importedTransactions } : card)),
      });
    } catch (error) {
      setCloudStatus(`שגיאה בייבוא הקובץ: ${error?.message || 'לא ידוע'}`);
      alert(`שגיאה בייבוא הקובץ: ${error?.message || 'לא ידוע'}`);
    }
  }

  function updatePendingTransaction(cardId, transactionId, field, value) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) => (
        card.id === cardId
          ? { ...card, pendingTransactions: (card.pendingTransactions || []).map((transaction) => (transaction.id === transactionId ? { ...transaction, [field]: field === 'amount' ? toNumber(value) : value } : transaction)) }
          : card
      )),
    });
  }

  function updateTransactionCategory(transactionId, newCategory) {
    let changedTransaction = null;
    const updatedCards = monthData.creditCards.map((card) => ({
      ...card,
      transactions: (card.transactions || []).map((transaction) => {
        if (transaction.id === transactionId) {
          changedTransaction = transaction;
          return { ...transaction, category: newCategory };
        }
        return transaction;
      }),
    }));
    if (!changedTransaction) return;
    const normalizedMerchant = normalizeMerchantName(changedTransaction.merchant);
    setLearnedRules((current) => ({ ...current, [normalizedMerchant]: newCategory }));
    setSelectedMonthData({
      ...monthData,
      creditCards: updatedCards.map((card) => ({
        ...card,
        transactions: (card.transactions || []).map((transaction) => (normalizeMerchantName(transaction.merchant) === normalizedMerchant ? { ...transaction, category: newCategory } : transaction)),
      })),
    });
  }

  function removePendingTransaction(cardId, transactionId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, pendingTransactions: (card.pendingTransactions || []).filter((transaction) => transaction.id !== transactionId) } : card)),
    });
  }

  function approvePendingTransactions(cardId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, transactions: [...(card.transactions || []), ...(card.pendingTransactions || [])], pendingTransactions: [] } : card)),
    });
  }

  function addTransaction(cardId) {
  setSelectedMonthData({
    ...monthData,
    creditCards: monthData.creditCards.map((card) =>
      card.id === cardId
        ? {
            ...card,
            pendingTransactions: [
              ...(card.pendingTransactions || []),
              {
                id: makeId('tx'),
                date: '',
                merchant: 'עסקה חדשה',
                category: 'אחר',
                amount: 0,
                originalAmount: 0,
                currency: 'ILS',
necessity: 'מותרות',
              },
            ],
          }
        : card
    ),
  });
}

  function removeTransaction(cardId, transactionId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, transactions: (card.transactions || []).filter((transaction) => transaction.id !== transactionId) } : card)),
    });
  }

  // Derived totals below power the dashboard cards, insights, charts, and warnings.
  const totalIncome = monthData.incomes.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalBankOpening = monthData.bankAccounts.reduce((sum, account) => sum + toNumber(account.openingBalance), 0);
  const totalBankClosing = monthData.bankAccounts.reduce((sum, account) => sum + toNumber(account.closingBalance), 0);
  const bankBalanceChange = totalBankClosing - totalBankOpening;
  const allBankTransactions = monthData.bankAccounts.flatMap((account) => account.transactions || []);
  const totalBankDeposits = allBankTransactions.filter((transaction) => toNumber(transaction.amount) > 0).reduce((sum, transaction) => sum + toNumber(transaction.amount), 0);
  const totalBankWithdrawals = allBankTransactions.filter((transaction) => toNumber(transaction.amount) < 0).reduce((sum, transaction) => sum + Math.abs(toNumber(transaction.amount)), 0);
  const allCreditTransactions = useMemo(() => monthData.creditCards.flatMap((card) => card.transactions || []), [monthData.creditCards]);
  const realCreditTransactions = allCreditTransactions.filter((transaction) => !isInternalTransferTransaction(transaction));
const internalCreditTransfers = allCreditTransactions.filter(isInternalTransferTransaction);

const totalCreditCards = realCreditTransactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
const totalInternalCreditTransfers = internalCreditTransfers.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalManualExpenses = monthData.manualExpenses.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalSavingsProducts = monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const totalSavingGoals = monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const totalPlannedSavings = totalSavingsProducts + totalSavingGoals;
  const includeSelfEmployed = Boolean(preferences.includeSelfEmployed);
  const selfEmployedVatDue = Math.max(0, toNumber(monthData.selfEmployed.vatCollected) - toNumber(monthData.selfEmployed.vatPaidOnExpenses));
  const rawSelfEmployedPayments = selfEmployedVatDue + toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance) + toNumber(monthData.selfEmployed.businessExpenses);
  const totalSelfEmployedPayments = includeSelfEmployed ? rawSelfEmployedPayments : 0;
  const totalExpenses = totalCreditCards + totalManualExpenses + totalSelfEmployedPayments;
const remainingCashFlow = totalIncome - totalExpenses - totalPlannedSavings;
const savingsRate = totalIncome ? (totalPlannedSavings / totalIncome) * 100 : 0;
const totalAssets =
  totalBankClosing +
  toNumber(monthData.emergencyFund) +
  monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.currentBalance), 0) +
  monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.currentAmount), 0);

const netWorth = totalAssets;
const emergencyMonths = toNumber(monthData.emergencyFund) / (totalExpenses || 1);
  const bankVsCalculatedCashFlow = bankBalanceChange - remainingCashFlow;
  const categoryTotals = useMemo(() => getCategoryTotals(realCreditTransactions), [realCreditTransactions]);
  const recurringTransactions = useMemo(() => detectRecurringTransactions(realCreditTransactions, months, selectedMonth), [realCreditTransactions, months, selectedMonth]);
  const monthlyCompare = useMemo(() => getMonthlyCompare(months, selectedMonth, comparePeriod), [months, selectedMonth, comparePeriod]);
  const trend = useMemo(() => getMonthlyTrend(months), [months]);
  const trendSixMonths = useMemo(() => trend.slice(-6), [trend]);
  const maxTrend = Math.max(1, ...trend.map((item) => item.total));
  const burnRate = trend.length ? trend.reduce((sum, item) => sum + item.total, 0) / trend.length : 0;
  const cashFlow = totalPlannedSavings;
  const topCategories = useMemo(() => Object.entries(categoryTotals).sort((a, b) => toNumber(b[1]) - toNumber(a[1])).slice(0, 6), [categoryTotals]);
  const filteredTransactions = useMemo(() => {
    const normalizedSearch = normalizeMerchantName(searchTerm);
    return realCreditTransactions.filter((transaction) => {
      const merchantMatch = normalizeMerchantName(transaction.merchant).includes(normalizedSearch);
      const categoryMatch = categoryFilter === 'הכול' || transaction.category === categoryFilter;
      const amount = toNumber(transaction.amount);
      return merchantMatch && categoryMatch && (minAmount === '' || amount >= toNumber(minAmount)) && (maxAmount === '' || amount <= toNumber(maxAmount));
    });
  }, [allCreditTransactions, searchTerm, categoryFilter, minAmount, maxAmount]);

  const financialHealthScore = calculateFinancialHealthScore(allCreditTransactions, modeConfig) || 0;
  const monthlyBudgetTarget = toNumber(preferences.monthlyBudgetTarget);
  const effectiveBudgetTarget = monthlyBudgetTarget ? monthlyBudgetTarget / modeConfig.strictness : 0;
  const budgetUsageRate = effectiveBudgetTarget ? (totalExpenses / effectiveBudgetTarget) * 100 : 0;
  const targetSavingsRate = toNumber(preferences.savingsRateTarget) || modeConfig.savingsTarget;
  const realInsights = useMemo(
    () => buildRealInsights(allCreditTransactions, recurringTransactions, totalIncome, preferences.financialMode, { savingsRate, burnRate, cashFlow, totalAssets }),
    [allCreditTransactions, recurringTransactions, totalIncome, preferences.financialMode, savingsRate, burnRate, cashFlow, totalAssets]
  );

  const operatingModeMessages = {
    Survival: 'המערכת מתמקדת כרגע בצמצום הוצאות ושמירה על יציבות.',
    Stable: 'המערכת מתמקדת באיזון פיננסי וחיסכון יציב.',
    Growth: 'המערכת מתמקדת בצמיחה, הגדלת הכנסות והשקעות.',
    'Wealth Building': 'המערכת מתמקדת באופטימיזציה ובניית הון ארוך טווח.',
  };

  const modeInsight = {
    Survival: 'המיקוד כרגע הוא הורדת burn rate וצמצום הוצאות לא חיוניות.',
    Stable: 'המיקוד כרגע הוא איזון בין איכות חיים לחיסכון יציב.',
    Growth: 'המיקוד כרגע הוא הגדלת הכנסות והשקעה בצמיחה.',
    'Wealth Building': 'המיקוד כרגע הוא בניית הון ואופטימיזציה פיננסית ארוכת טווח.',
  };

  // Notifications are derived from current totals and global notification preferences.
  const activeNotifications = [
    preferences.notifications?.budget80 && budgetUsageRate >= modeConfig.budgetWarningAt ? `הגעתם ל־${modeConfig.budgetWarningAt}% מהתקציב לפי מצב ${modeConfig.label}.` : null,
    preferences.notifications?.woltSpike && (categoryTotals['מסעדות ובתי קפה'] || 0) > (CATEGORY_BUDGETS['מסעדות ובתי קפה'] || 0) ? 'מסעדות ובתי קפה חרגו מהתקציב שהוגדר.' : null,
    preferences.notifications?.savingsDrop && savingsRate < targetSavingsRate ? 'שיעור החיסכון נמוך מהיעד שהוגדר.' : null,
  ].filter(Boolean);

  const monthlyStory = totalIncome
  ? `במצב ${modeConfig.label}, ההוצאות האמיתיות החודש הן ${SHEKEL.format(totalExpenses)}, והועברו לחיסכון ${SHEKEL.format(totalPlannedSavings)}. העו״ש הנוכחי הוא נתון אמיתי שהוזן ידנית או הגיע מהבנק. אחרי הוצאות והעברות לחיסכון נשאר תזרים חודשי של ${SHEKEL.format(remainingCashFlow)}. שיעור החיסכון הוא ${formatPercent(savingsRate)} מתוך יעד של ${formatPercent(targetSavingsRate)}.`
  : `מצב ${modeConfig.label} פעיל. התחילו להזין הכנסות, הוצאות, עו״ש וחסכונות כדי להבין את המצב הפיננסי האמיתי.`;
  
  const monthlyCompareStory = monthlyCompare.hasPrevious
    ? `לעומת ממוצע ${monthlyCompare.period.label} (${monthlyCompare.compareMonthKeys.length} חודשים), ההוצאות ${monthlyCompare.current.expenses >= monthlyCompare.previous.expenses ? 'עלו' : 'ירדו'} ב־${SHEKEL.format(Math.abs(monthlyCompare.current.expenses - monthlyCompare.previous.expenses))}, והיתרה ${monthlyCompare.current.net >= monthlyCompare.previous.net ? 'השתפרה' : 'נחלשה'} ב־${SHEKEL.format(Math.abs(monthlyCompare.current.net - monthlyCompare.previous.net))}.`
    : `אין עדיין מספיק חודשים להשוואת ${monthlyCompare.period.label}. הוסיפי עוד חודשים כדי לקבל Monthly Compare אמיתי ורחב יותר.`;

  function getBudgetHeatColor(category, amount) {
    const budget = CATEGORY_BUDGETS[category];
    if (!budget) return 'bg-white border-neutral-200 text-neutral-900';
    if (amount > budget) return 'bg-red-50 border-red-200 text-red-900';
    if (amount >= budget * 0.8) return 'bg-amber-50 border-amber-200 text-amber-900';
    return 'bg-[#F4F6F1] border-[#D6DDCF] text-[#66725E]';
  }

  // Conic gradient keeps the chart dependency-free while still showing category proportions.
  const pieChart = topCategories.length
    ? `conic-gradient(${topCategories.map(([, amount], index) => {
        const start = topCategories.slice(0, index).reduce((sum, [, value]) => sum + value, 0) / (totalCreditCards || 1);
        const end = topCategories.slice(0, index + 1).reduce((sum, [, value]) => sum + value, 0) / (totalCreditCards || 1);
        const colors = [activeTheme.accent, '#111111', '#737373', '#a3a3a3', '#d4d4d4', '#f5f5f5'];
        return `${colors[index % colors.length]} ${start * 100}% ${end * 100}%`;
      }).join(', ')})`
    : 'conic-gradient(#dddddd 0% 100%)';

  function applyAuthSession(session) {
  setAuthSession(session);
  setSelectedMonth(getCurrentMonthKey());
  setMonths({ [getCurrentMonthKey()]: createDefaultMonth() });
  setLearnedRules({});
  setHasLoadedCloud(false);
  setHasAttemptedCloudLoad(false);
}

async function handleSignIn(event) {
  event.preventDefault();
  try {
    setAuthStatus('מתחברת...');
    const session = await signInWithSupabasePassword(authEmail, authPassword, supabaseConfig);
    applyAuthSession(session);
    setAuthPassword('');
    setAuthStatus('מחוברת');
    setCloudStatus('מחוברת לחשבון Supabase');
  } catch (error) {
    setAuthStatus(error?.message || 'הכניסה נכשלה');
  }
}

  function handleSignOut() {
  clearAuthSession();
  setAuthSession(null);
  setSelectedMonth(getCurrentMonthKey());
  setMonths({ [getCurrentMonthKey()]: createDefaultMonth() });
  setLearnedRules({});
  setHasLoadedCloud(false);
  setHasAttemptedCloudLoad(false);
  setAuthStatus('התנתקת');
  setCloudStatus('התנתקת, נשמר מקומית עד כניסה מחדש');
}

  // Auth UI is intentionally disabled for now until global settings and cloud sync are fully stable.
  // Login UI is intentionally parked behind false until Supabase Auth is enabled as a separate step.
  if (!authSession && preferences.syncMode !== 'Local Only') {
    return (
      <div dir="rtl" className={`min-h-screen p-6 text-right transition-colors duration-300 ${activeTheme.page}`} style={{ fontFamily: 'Circular, Arial, Helvetica, sans-serif' }}>
        <div className="mx-auto flex min-h-[calc(100vh-48px)] max-w-5xl items-center justify-center">
          <div className="grid w-full gap-6 lg:grid-cols-[1fr_0.9fr]">
            <section className="rounded-[32px] border border-neutral-200 bg-white p-8 shadow-sm">
              <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">SECURE ACCESS</div>
              <h1 className="mt-4 text-4xl font-semibold tracking-tight text-neutral-950">כניסה למערכת</h1>
              <p className="mt-4 text-sm leading-7 text-neutral-500">
  התחברי עם המייל והסיסמה שלך. לכל משתמש נוצר בית פיננסי פרטי ונפרד.
</p>
              <form onSubmit={handleSignIn} className="mt-7 grid gap-4">
                <label className="text-sm font-semibold text-neutral-600">
                  אימייל
                  <Field type="email" value={authEmail} onChange={(event) => setAuthEmail(event.target.value)} className="mt-2 w-full" placeholder="name@example.com" />
                </label>
                <label className="text-sm font-semibold text-neutral-600">
                  סיסמה
                  <Field type="password" value={authPassword} onChange={(event) => setAuthPassword(event.target.value)} className="mt-2 w-full" placeholder="••••••••" />
                </label>
                <PrimaryButton theme={activeTheme} type="submit" className="mt-2 w-full">
  כניסה
</PrimaryButton>

<div className="text-center text-xs font-medium text-neutral-400">
  אין לך חשבון?
</div>
                <GhostButton
  type="button"
  onClick={async () => {
    try {
      setAuthStatus('יוצרת חשבון...');
      const session = await signUpWithSupabasePassword(authEmail, authPassword, supabaseConfig);
      applyAuthSession(session);
      setAuthPassword('');
      setAuthStatus('החשבון נוצר והתחברת');
      setCloudStatus('נוצר בית פיננסי חדש');
    } catch (error) {
      setAuthStatus(error?.message || 'יצירת החשבון נכשלה');
    }
  }}
  className="w-full"
>
  יצירת חשבון חדש
</GhostButton>
              </form>
              {authStatus ? <div className="mt-4 rounded-2xl bg-neutral-50 p-4 text-sm leading-7 text-neutral-600">{authStatus}</div> : null}
            </section>

            <section className="rounded-[32px] border border-neutral-200 bg-neutral-50 p-8 shadow-sm">
              <h2 className="text-2xl font-semibold tracking-tight text-neutral-950">לפני הכניסה</h2>
              <div className="mt-5 grid gap-4 text-sm leading-7 text-neutral-600">
                <p>צריך ליצור משתמש ב־Supabase דרך Authentication → Users.</p>
                <p>אם עוד לא יצרת משתמש, הכניסה תיכשל עד שיוגדר אימייל וסיסמה.</p>
                <p>מצב Local Only עדיין אפשרי דרך ההגדרות אחרי כניסה, אבל בשביל ענן צריך חשבון.</p>
              </div>
              <div className="mt-6 rounded-2xl border border-neutral-200 bg-white p-4 text-sm text-neutral-600">
                <strong>Supabase</strong>
                <div className="mt-2">{setupHealth.supabaseEnv ? 'מחובר להגדרות Supabase' : 'חסרים URL או Publishable Key בהגדרות שנשמרו בדפדפן'}</div>
              </div>
            </section>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div dir="rtl" className={`min-h-screen p-4 text-right transition-colors duration-300 sm:p-6 ${activeTheme.page} ${isDark ? 'theme-dark' : ''}`} style={{ fontFamily: 'Circular, Arial, Helvetica, sans-serif' }}>
      <style>{`
        .theme-dark .bg-white,
        .theme-dark section.bg-white,
        .theme-dark input.bg-white,
        .theme-dark select.bg-white,
        .theme-dark button.bg-white { background-color: #181818 !important; }
        .theme-dark .dark-surface,
        .theme-dark .dark-nav,
        .theme-dark .bg-white\/95 { background-color: rgba(18, 18, 18, 0.96) !important; border-color: #333333 !important; color: #F5F5F5 !important; }
        .theme-dark .bg-neutral-50,
        .theme-dark .bg-neutral-100 { background-color: #202020 !important; }
        .theme-dark .border-neutral-100,
        .theme-dark .border-neutral-200,
        .theme-dark .border-neutral-300 { border-color: #333333 !important; }
        .theme-dark .text-neutral-950,
        .theme-dark .text-neutral-900,
        .theme-dark .text-neutral-800,
        .theme-dark .text-neutral-700 { color: #F5F5F5 !important; }
        .theme-dark .text-neutral-600,
        .theme-dark .text-neutral-500,
        .theme-dark .text-neutral-400 { color: #A3A3A3 !important; }
        .theme-dark input,
        .theme-dark select { background-color: #181818 !important; color: #F5F5F5 !important; border-color: #3A3A3A !important; }
        .theme-dark input::placeholder { color: #737373 !important; }
        .theme-dark .shadow-sm,
        .theme-dark .shadow-md { box-shadow: 0 1px 2px rgba(0, 0, 0, 0.35) !important; }
        .theme-dark .from-amber-50 { --tw-gradient-from: #2A2418 !important; --tw-gradient-to: rgba(42, 36, 24, 0) !important; }
        .theme-dark .to-white { --tw-gradient-to: #181818 !important; }
        .theme-dark .hero-banner { background: #151515 !important; border-color: #333333 !important; }
        .no-orphans { text-wrap: balance; overflow-wrap: normal; word-break: keep-all; }
        .no-single-word-lines { white-space: normal; word-break: keep-all; overflow-wrap: normal; hyphens: none; }
        .nowrap-chip { white-space: nowrap; }
      `}</style>

      <div className="mx-auto w-full max-w-7xl overflow-x-hidden space-y-5 sm:space-y-7">
        <div className="dark-nav md:sticky md:top-0 z-40 rounded-2xl border border-neutral-200 bg-white/95 p-2 shadow-sm backdrop-blur-xl" style={isDark ? { backgroundColor: 'rgba(18, 18, 18, 0.96)', borderColor: '#333333' } : undefined}>
          <div className="flex items-center gap-1 overflow-x-auto whitespace-nowrap pb-1 [-ms-overflow-style:none] [scrollbar-width:none] [&::-webkit-scrollbar]:hidden">
            {TABS.map((tab) => (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveTab(tab.id)}
                className={`shrink-0 rounded-xl px-3 py-2 text-sm font-semibold transition sm:px-4 sm:py-2.5 ${activeTab === tab.id ? 'text-white shadow-sm' : 'text-neutral-600 hover:bg-neutral-100 hover:text-neutral-950'}`}
                style={activeTab === tab.id ? { backgroundColor: activeTheme.accent } : undefined}
              >
                {tab.label}
              </button>
            ))}
            <div className="ms-2 flex shrink-0 items-center gap-2 rounded-xl border border-neutral-200 bg-neutral-50 px-3 py-2 text-xs font-semibold text-neutral-600">
              <span>{authSession?.email || 'Local'}</span>
              {authSession ? <button type="button" onClick={handleSignOut} className="text-neutral-900 underline">התנתקות</button> : null}
            </div>
          </div>
        </div>

        <section className="hero-banner overflow-hidden rounded-[24px] border border-neutral-200 bg-white shadow-sm sm:rounded-3xl" style={isDark ? { backgroundColor: '#151515', borderColor: '#333333' } : undefined}>
          <div className={`p-5 sm:p-8 ${isDark ? 'text-white' : 'text-neutral-950'}`} style={isDark ? { backgroundColor: '#151515' } : undefined}>
            <div className="flex flex-col gap-5 lg:flex-row lg:items-center lg:justify-between">
              <div>
                <input value={monthData.dashboardTitle} onChange={(event) => updateMonthField('dashboardTitle', event.target.value)} className="w-full max-w-3xl rounded-xl border border-transparent bg-transparent px-0 py-1 text-3xl font-semibold leading-tight tracking-tight text-neutral-950 outline-none transition placeholder:text-neutral-400 sm:py-2 sm:text-4xl md:text-5xl" placeholder="שם הדשבורד המשפחתי" />
                <p className="mt-3 max-w-4xl text-sm leading-7 text-neutral-500 no-orphans no-single-word-lines sm:mt-4 sm:text-base sm:leading-8">{noSingleWordLine('ממלאים הכנסות, הוצאות, אשראי, עצמאי, קרנות ויעדים. המערכת מחשבת תזרים, חיסכון ותובנות אמיתיות.')}</p>
                <div className="mt-4 inline-flex max-w-full rounded-full px-3 py-2 text-xs font-semibold leading-6 no-orphans sm:px-4 sm:text-sm" style={{ backgroundColor: activeTheme.soft, color: activeTheme.text }}>{noSingleWordLine(`${modeInsight[preferences.financialMode] || modeInsight.Stable} יעד חיסכון: ${formatPercent(targetSavingsRate)} | התראה ב־${modeConfig.budgetWarningAt}%`)}</div>
                <div className="mt-3 inline-flex max-w-full rounded-full border border-neutral-200 bg-neutral-50 px-3 py-2 text-xs font-medium leading-6 text-neutral-600 no-orphans sm:mt-5 sm:px-4 sm:text-sm">{noSingleWordLine(cloudStatus)}</div>
              </div>
              <div className="rounded-2xl border border-neutral-200 bg-neutral-50 p-5">
                <label className="text-xs font-semibold uppercase tracking-widest text-neutral-400">חודש</label>
                <input type="month" value={selectedMonth} onChange={(event) => ensureMonth(event.target.value)} className="mt-2 w-full rounded-xl border border-neutral-300 bg-white px-4 py-3 text-lg font-semibold text-neutral-900 outline-none focus:border-neutral-900 focus:ring-2 focus:ring-neutral-100" />
              </div>
            </div>
          </div>

          <div className="dark-surface grid grid-cols-1 gap-3 border-t border-neutral-100 bg-white p-4 sm:gap-4 sm:p-6 md:grid-cols-2 xl:grid-cols-4" style={isDark ? { backgroundColor: '#151515', borderColor: '#333333' } : undefined}>
            <StatCard title="סה״כ הכנסות" value={SHEKEL.format(totalIncome)} note="כל מקורות ההכנסה" tone="good" />
            <StatCard title="עו״ש נוכחי" value={SHEKEL.format(totalBankClosing)} note={`שינוי החודש: ${SHEKEL.format(bankBalanceChange)}`} tone={totalBankClosing >= 0 ? 'good' : 'danger'} />
            <StatCard title="סה״כ הוצאות" value={SHEKEL.format(totalExpenses)} note={effectiveBudgetTarget ? `${formatPercent(budgetUsageRate)} מתוך יעד ${modeConfig.label}` : `${totalIncome ? formatPercent((totalExpenses / totalIncome) * 100) : '0%'} מההכנסה`} tone={(effectiveBudgetTarget && totalExpenses > effectiveBudgetTarget) || (totalIncome && totalExpenses > totalIncome) ? 'danger' : budgetUsageRate >= modeConfig.budgetWarningAt ? 'warn' : 'neutral'} />
            <StatCard title="סה״כ אשראי" value={SHEKEL.format(totalCreditCards)} note="מכרטיסי האשראי" />
            <StatCard title="עצמאי" value={SHEKEL.format(totalSelfEmployedPayments)} note={includeSelfEmployed ? 'כלול בתזרים המשפחתי' : 'לא כלול בתזרים'} />
            <StatCard title="העברות לחיסכון" value={SHEKEL.format(totalPlannedSavings)} note="לא הוצאה, אלא העברה פנימית" tone="good" />
            <StatCard
  title="עודף / גירעון חודשי"
  value={SHEKEL.format(remainingCashFlow)}
  note="אחרי הוצאות אמיתיות והעברות לחיסכון"
  tone={remainingCashFlow >= 0 ? 'good' : 'danger'}
  help="הכנסות פחות הוצאות אמיתיות פחות העברות לחיסכון. אם המספר שלילי, החודש בגירעון תזרימי."
/>
            <StatCard title="שווי נטו שהוזן" value={SHEKEL.format(netWorth)} note={`עו״ש + חסכונות + יעדים + קרן חירום`} tone="good" />
          </div>
        </section>

        {activeTab === 'dashboard' ? (
          <>
            {(preferences.showMonthlyStory || preferences.showFinancialHealth || activeNotifications.length > 0) ? (
              <Section>
                <div className="grid gap-4 xl:grid-cols-[1.4fr_0.6fr]">
                  {activeNotifications.length > 0 ? (
                    <div className="grid gap-3 md:grid-cols-3 lg:col-span-2">
                      {activeNotifications.map((notification) => (
                        <div key={notification} className="rounded-[24px] border border-amber-200 bg-gradient-to-br from-amber-50 to-white px-5 py-5 shadow-sm">
                          <div className="flex items-center gap-3">
                            <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-amber-100 text-amber-700">⚠</div>
                            <div>
                              <div className="text-xs font-semibold uppercase tracking-widest text-amber-500">Smart Notification</div>
                              <div className="mt-1 text-sm font-semibold leading-6 text-amber-900 no-orphans">{noSingleWordLine(notification)}</div>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : null}

                  {preferences.showMonthlyStory ? (
                    <div>
                      <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">MONTHLY STORY</div>
                      <h2 className="mt-3 text-3xl font-semibold leading-tight tracking-tight text-neutral-950 sm:mt-4 sm:text-4xl">הסיפור של החודש שלכם</h2>
                      <p className="mt-4 max-h-40 max-w-3xl overflow-auto text-sm leading-7 text-neutral-600 no-orphans sm:max-h-none sm:text-lg sm:leading-9">
  {noSingleWordLine(monthlyStory)}
</p>
                      <div className="mt-8 grid gap-3 md:grid-cols-3">
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">Burn Rate</div><div className="mt-2 text-xl font-semibold text-neutral-950">{SHEKEL.format(burnRate)}</div></div>
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">Cash Flow לחיסכון</div><div className="mt-2 text-xl font-semibold text-neutral-950">{SHEKEL.format(cashFlow)}</div></div>
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">שיעור חיסכון</div><div className="mt-2 text-xl font-semibold text-neutral-950">{formatPercent(savingsRate)}</div></div>
                      </div>
                    </div>
                  ) : null}

                  {preferences.showFinancialHealth ? (
                    <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-6">
                      <div className="text-sm font-semibold text-neutral-500">Financial Health</div>
                      <div className="mt-4 text-6xl font-semibold text-neutral-950">{financialHealthScore}</div>
                      <div className="mt-5 h-3 overflow-hidden rounded-full bg-neutral-200"><div className="h-full rounded-full" style={{ width: `${financialHealthScore}%`, backgroundColor: activeTheme.accent }} /></div>
                      <div className="mt-3 text-sm leading-7 text-neutral-500">ציון לפי מצב {modeConfig.label}: קשיחות תקציב, חריגות, פיזור הוצאות ועסקאות גדולות.</div>
                    </div>
                  ) : null}
                </div>
              </Section>
            ) : null}

            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                <div>
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">ACCOUNTS</div>
                  <h2 className="mt-3 text-3xl font-semibold tracking-tight text-neutral-950">חשבונות ועו״ש</h2>
                  <p className="mt-3 max-w-3xl text-sm leading-7 text-neutral-500 no-orphans">{noSingleWordLine('מעלים פירוט עו״ש CSV/Excel מהבנק, והמערכת מחשבת יתרת פתיחה, יתרה נוכחית ותנועות. היתרה אחרי הכול היא תזרים מחושב ולא נכנסת אוטומטית לעו״ש.')}</p>
                </div>
                <PrimaryButton theme={activeTheme} onClick={addBankAccount}>+ הוספת חשבון</PrimaryButton>
              </div>

              <div className="mt-6 grid gap-4 md:grid-cols-4">
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">יתרת פתיחה</div>
                  <div className="mt-3 text-2xl font-semibold text-neutral-950">{SHEKEL.format(totalBankOpening)}</div>
                </div>
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">יתרה נוכחית</div>
                  <div className="mt-3 text-2xl font-semibold text-neutral-950">{SHEKEL.format(totalBankClosing)}</div>
                </div>
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">נכנס לעו״ש</div>
                  <div className="mt-3 text-2xl font-semibold text-neutral-950">{SHEKEL.format(totalBankDeposits)}</div>
                </div>
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">יצא מהעו״ש</div>
                  <div className="mt-3 text-2xl font-semibold text-neutral-950">{SHEKEL.format(totalBankWithdrawals)}</div>
                </div>
              </div>

              <div className="mt-5 grid gap-3">
                {monthData.bankAccounts.map((account) => (
                  <div key={account.id} className="rounded-[22px] border border-neutral-200 bg-white p-4 shadow-sm">
  <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
    <div className="grid flex-1 gap-3 md:grid-cols-2 xl:grid-cols-4">
      <LabeledField label="חשבון">
        <Field value={account.name} onChange={(event) => updateRow('bankAccounts', account.id, 'name', event.target.value)} placeholder="עו״ש משותף" />
      </LabeledField>

      <LabeledField label="שייך ל">
        <Field value={account.owner} onChange={(event) => updateRow('bankAccounts', account.id, 'owner', event.target.value)} placeholder="משפחה" />
      </LabeledField>

      <LabeledField label="יתרת פתיחה">
        <Field type="number" value={account.openingBalance} onChange={(event) => updateRow('bankAccounts', account.id, 'openingBalance', event.target.value)} />
      </LabeledField>

      <LabeledField label="יתרה נוכחית">
        <Field type="number" value={account.closingBalance} onChange={(event) => updateRow('bankAccounts', account.id, 'closingBalance', event.target.value)} />
      </LabeledField>
    </div>

    <div className="flex gap-2 items-end">
      <label className="flex-1 cursor-pointer rounded-lg border border-neutral-200 bg-neutral-50 px-3 py-2 text-center text-xs font-semibold text-neutral-700 transition hover:border-neutral-400 hover:bg-white">
        ייבוא עו״ש
        <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) importBankFile(account.id, file); }} />
      </label>

      <GhostButton onClick={() => removeRow('bankAccounts', account.id)} className="w-10 px-0">
        ×
      </GhostButton>
    </div>
  </div>
                    {account.importedFile ? <div className="mt-3 rounded-xl bg-neutral-50 px-4 py-3 text-sm text-neutral-600">נקלט קובץ עו״ש: <strong>{account.importedFile}</strong> · {account.transactions?.length || 0} תנועות</div> : null}
                    {account.transactions?.length ? (
                      <div className="mt-3 max-h-64 overflow-auto rounded-2xl border border-neutral-200">
                        <div className="grid grid-cols-[110px_1fr_130px_130px] bg-neutral-100 px-4 py-3 text-xs font-semibold text-neutral-600">
                          <div>תאריך</div>
                          <div>פירוט</div>
                          <div>סכום</div>
                          <div>יתרה</div>
                        </div>
                        {account.transactions.slice(0, 80).map((transaction) => (
                          <div key={transaction.id} className="grid grid-cols-[110px_1fr_130px_130px] gap-3 border-t border-neutral-100 px-4 py-3 text-sm">
                            <div className="text-neutral-500">{transaction.date}</div>
                            <div>{transaction.description}</div>
                            <div className={toNumber(transaction.amount) >= 0 ? 'font-semibold text-[#66725E]' : 'font-semibold text-red-700'}>{SHEKEL.format(transaction.amount)}</div>
                            <div>{transaction.balance ? SHEKEL.format(transaction.balance) : '—'}</div>
                          </div>
                        ))}
                      </div>
                    ) : null}
                  </div>
                ))}
              </div>
            </Section>

            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
                <div>
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">MONTHLY COMPARE</div>
                  <h2 className="mt-3 text-3xl font-semibold tracking-tight text-neutral-950">השוואה לאורך זמן</h2>
                  <p className="mt-3 max-w-3xl text-sm leading-7 text-neutral-500 no-orphans">{noSingleWordLine(monthlyCompareStory)}</p>
                </div>
                <div className="flex flex-wrap gap-2">
                  {COMPARE_PERIODS.map((period) => (
                    <button
                      key={period.id}
                      type="button"
                      onClick={() => setComparePeriod(period.id)}
                      className="rounded-full border px-4 py-2 text-sm font-semibold transition"
                      style={comparePeriod === period.id ? { backgroundColor: activeTheme.soft, color: activeTheme.text, borderColor: activeTheme.accent } : undefined}
                    >
                      {period.label}
                    </button>
                  ))}
                </div>
              </div>
              {monthlyCompare.hasPrevious ? (
                <div className="mt-6 grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-4">
                  {monthlyCompare.rows.map((row) => {
                    const isGoodDirection = row.key === 'income' || row.key === 'net' || row.key === 'savingsRate' || row.key === 'savings';
                    const improved = isGoodDirection ? row.diff >= 0 : row.diff <= 0;
                    const value = row.type === 'percent' ? formatPercent(row.currentValue) : SHEKEL.format(row.currentValue);
                    const diff = row.type === 'percent' ? formatPercent(Math.abs(row.diff)) : SHEKEL.format(Math.abs(row.diff));
                    return (
                      <div key={row.key} className={`rounded-[24px] border p-5 ${improved ? 'border-[#D6DDCF] bg-[#F4F6F1] text-[#66725E]' : 'border-amber-200 bg-amber-50 text-amber-900'}`}>
                        <div className="text-xs font-semibold uppercase tracking-widest opacity-70">{row.label}</div>
                        <div className="mt-3 text-2xl font-semibold text-neutral-950">{value}</div>
                        <div className="mt-2 text-sm font-semibold">{improved ? 'שיפור' : 'דורש תשומת לב'}: {diff}</div>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div className="mt-6">
                  <EmptyState title="אין עדיין מספיק חודשים" text="צרי או מלאי נתונים בעוד חודשים כדי לראות השוואה אוטומטית מול 3 חודשים, 6 חודשים, שנה או כל התקופה." />
                </div>
              )}
            </Section>

            {(preferences.showCategoryChart || preferences.showTrendChart) ? (
              <section className="grid grid-cols-1 gap-5 xl:grid-cols-3">
                {preferences.showCategoryChart ? (
                  <Section>
                    <div className="flex items-center justify-between gap-4"><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">התפלגות הוצאות לפי קטגוריות</h2><span className="rounded-full border border-neutral-200 bg-white px-3 py-1 text-xs font-medium text-neutral-500">Heatmap</span></div>
                    <div className="mx-auto mt-5 h-40 w-40 rounded-full sm:h-56 sm:w-56" style={{ background: pieChart }} />
                    <div className="mt-6 space-y-2">
                      {topCategories.length ? topCategories.map(([category, amount]) => <div key={category} className={`flex justify-between rounded-2xl border px-4 py-3 text-sm ${getBudgetHeatColor(category, amount)}`}><span>{category}</span><strong>{SHEKEL.format(amount)}</strong></div>) : <EmptyState title="אין עדיין קטגוריות" text="העלי פירוט אשראי כדי לראות התפלגות צבעונית לפי קטגוריות." />}
                    </div>
                  </Section>
                ) : null}

                {preferences.showTrendChart ? (
                  <Section className="lg:col-span-2">
                    <div className="flex items-center justify-between gap-4"><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">מגמת 6 חודשים</h2><span className="rounded-full border border-neutral-200 bg-white px-3 py-1 text-xs font-medium text-neutral-500">Income · Expenses · Savings</span></div>
                    <p className="mt-2 text-sm leading-7 text-neutral-500">הכנסות, הוצאות וחיסכון נטו לפי חודשים. בלי ספריית גרפים חיצונית, כדי שה־build יישאר נקי.</p>
                    <div className="mt-6">
                      <TrendLineChart data={trendSixMonths} theme={activeTheme} />
                    </div>
                    <div className="mt-4 rounded-2xl bg-white p-4 text-sm text-neutral-600">Burn Rate ממוצע: <strong>{SHEKEL.format(burnRate)}</strong> | Cash Flow לחיסכון: <strong>{SHEKEL.format(cashFlow)}</strong></div>
                  </Section>
                ) : null}
              </section>
            ) : null}
          </>
        ) : null}

        {activeTab === 'credit' ? (
          <>
            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                <div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">סיכום כרטיסי אשראי</h2><p className="mt-2 text-sm text-neutral-500">כאן מעלים CSV/Excel לכל כרטיס, בודקים קטגוריות, ואז מאשרים הכנסה להוצאות.</p></div>
                <PrimaryButton theme={activeTheme} onClick={addCreditCard}>+ הוספת כרטיס</PrimaryButton>
              </div>
              <div className="mt-7 grid gap-8 xl:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2">
                {monthData.creditCards.map((card) => {
                  const cardTotal = (card.transactions || []).reduce((sum, item) => sum + toNumber(item.amount), 0);
                  return (
                    <CreditCardPanel
                      key={card.id}
                      card={card}
                      cardTotal={cardTotal}
                      onUpdateCard={updateCreditCard}
                      onRemoveCard={removeCreditCard}
                      onImportFile={importCreditFile}
                      onUpdatePending={updatePendingTransaction}
                      onRemovePending={removePendingTransaction}
                      onApprovePending={approvePendingTransactions}
                      onAddTransaction={addTransaction}
                      onRemoveTransaction={removeTransaction}
                      onUpdateCategory={updateTransactionCategory}
                      theme={activeTheme}
                    />
                  );
                })}
              </div>
            </Section>

            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                <div>
                  <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">הוצאות ידניות</h2>
                  <p className="mt-2 text-sm text-neutral-500">הוצאות שלא נכנסות מכרטיסי האשראי ונחשבות יחד עם ההוצאות החודשיות.</p>
                </div>
                <PrimaryButton theme={activeTheme} onClick={addManualExpense}>+ הוספת הוצאה</PrimaryButton>
              </div>
              <div className="mt-7 overflow-x-auto rounded-[24px] border border-neutral-200 bg-white">
                <div className="min-w-[720px]">
                  <div className="grid grid-cols-[minmax(320px,1fr)_180px_180px_60px] gap-3 bg-neutral-100 px-5 py-4 text-sm font-semibold text-neutral-700"><div>קטגוריה</div><div>סוג</div><div>סכום</div><div /></div>
                  {monthData.manualExpenses.map((expense) => (
                    <div key={expense.id} className="grid grid-cols-[minmax(320px,1fr)_180px_180px_60px] gap-3 border-t border-neutral-100 p-4">
                      <Field value={expense.category} onChange={(event) => updateRow('manualExpenses', expense.id, 'category', event.target.value)} className="w-full" />
                      <SelectField value={expense.type} onChange={(event) => updateRow('manualExpenses', expense.id, 'type', event.target.value)} className="w-full"><option>קבועה</option><option>משתנה</option><option>חיסכון</option><option>חד פעמית</option></SelectField>
                      <Field type="number" value={expense.amount} onChange={(event) => updateRow('manualExpenses', expense.id, 'amount', event.target.value)} className="w-full" />
                      <GhostButton onClick={() => removeRow('manualExpenses', expense.id)} className="px-0">×</GhostButton>
                    </div>
                  ))}
                </div>
              </div>
            </Section>

            <Section>
              <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">כל עסקאות האשראי המסוננות</h2>
              <div className="mt-6 overflow-x-auto rounded-[24px] border border-neutral-200 bg-white">
                <div className="min-w-[860px]">
                  <div className="grid grid-cols-[110px_1fr_170px_120px_90px] bg-neutral-100 px-6 py-4 text-sm font-semibold text-neutral-700"><div>תאריך</div><div>בית עסק</div><div>קטגוריה לומדת</div><div>סכום</div><div>זיהוי</div></div>
                  {filteredTransactions.map((transaction) => {
                    const isRecurring = recurringTransactions.some((item) => item.id === transaction.id);
                    return (
                      <div key={transaction.id} className="grid grid-cols-[110px_1fr_170px_120px_90px] gap-4 border-t border-neutral-100 px-6 py-4">
                        <div>{transaction.date}</div>
                        <div>{transaction.merchant}</div>
                        <SelectField value={transaction.category} onChange={(event) => updateTransactionCategory(transaction.id, event.target.value)}>{EXPENSE_CATEGORIES.map((category) => <option key={category}>{category}</option>)}</SelectField>
                        <div className="font-semibold">{SHEKEL.format(transaction.amount)}</div>
                        <div>{isRecurring ? 'חוזר קבוע' : '—'}</div>
                      </div>
                    );
                  })}
                  {filteredTransactions.length === 0 ? <div className="p-16 text-center text-neutral-400">לא נמצאו עסקאות לפי החיפוש והפילטרים שבחרתם.</div> : null}
                </div>
              </div>
            </Section>
          </>
        ) : null}

        {activeTab === 'savings' ? (
          <>
            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">קרנות, פנסיה וחסכונות</h2><p className="mt-2 text-sm text-neutral-500">הפרשות חודשיות לקרן השתלמות, פנסיה וחסכונות קבועים.</p></div><PrimaryButton theme={activeTheme} onClick={addSavingsProduct}>+ הוספת חיסכון</PrimaryButton></div>
              <div className="mt-7 grid gap-3">
                {monthData.savingsProducts.map((product) => (
                  <div key={product.id} className="grid gap-3 rounded-[24px] border border-neutral-200 p-4 md:grid-cols-[1.4fr_150px_130px_160px_160px_44px]">
                    <LabeledField label="שם החיסכון"><Field value={product.name} onChange={(event) => updateRow('savingsProducts', product.id, 'name', event.target.value)} placeholder="למשל פנסיה נועה" /></LabeledField>
                    <LabeledField label="סוג"><SelectField value={product.type} onChange={(event) => updateRow('savingsProducts', product.id, 'type', event.target.value)}><option>קרן השתלמות</option><option>פנסיה</option><option>קופת גמל</option><option>חיסכון</option><option>השקעות</option></SelectField></LabeledField>
                    <LabeledField label="שייך ל"><Field value={product.owner} onChange={(event) => updateRow('savingsProducts', product.id, 'owner', event.target.value)} placeholder="נועה / אורן" /></LabeledField>
                    <LabeledField label="הפקדה חודשית"><Field type="number" value={product.monthlyDeposit} onChange={(event) => updateRow('savingsProducts', product.id, 'monthlyDeposit', event.target.value)} placeholder="₪ לחודש" /></LabeledField>
                    <LabeledField label="יתרה נוכחית"><Field type="number" value={product.currentBalance} onChange={(event) => updateRow('savingsProducts', product.id, 'currentBalance', event.target.value)} placeholder="כמה נצבר" /></LabeledField>
                    <div className="flex items-end"><GhostButton onClick={() => removeRow('savingsProducts', product.id)} className="w-full px-0">×</GhostButton></div>
                  </div>
                ))}
              </div>
            </Section>

            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">יעדי חיסכון</h2><p className="mt-2 text-sm text-neutral-500">טיסה ליפן, חתונה, קרן חירום וכל יעד אחר.</p></div><PrimaryButton theme={activeTheme} onClick={addSavingGoal}>+ הוספת יעד</PrimaryButton></div>
              <div className="mt-6 grid gap-5 md:grid-cols-3">
                {monthData.savingGoals.map((goal) => {
                  const progress = toNumber(goal.targetAmount) ? Math.min(100, Math.round((toNumber(goal.currentAmount) / toNumber(goal.targetAmount)) * 100)) : 0;
                  const remaining = Math.max(0, toNumber(goal.targetAmount) - toNumber(goal.currentAmount));
                  const monthlyDeposit = Math.max(1, toNumber(goal.monthlyDeposit));
                  const etaMonths = Math.ceil(remaining / monthlyDeposit);
                  const boostedEta = Math.ceil(remaining / (monthlyDeposit + 500));
                  return (
                    <div key={goal.id} className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                      <LabeledField label="שם היעד"><Field value={goal.name} onChange={(event) => updateRow('savingGoals', goal.id, 'name', event.target.value)} className="w-full font-semibold" placeholder="למשל טיסה ליפן" /></LabeledField>
                      <div className="mt-3 grid gap-3">
                        <LabeledField label="סכום יעד"><Field type="number" value={goal.targetAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'targetAmount', event.target.value)} placeholder="כמה צריך להגיע" /></LabeledField>
                        <LabeledField label="נצבר עד עכשיו"><Field type="number" value={goal.currentAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'currentAmount', event.target.value)} placeholder="כמה כבר יש" /></LabeledField>
                        <LabeledField label="הפקדה חודשית"><Field type="number" value={goal.monthlyDeposit} onChange={(event) => updateRow('savingGoals', goal.id, 'monthlyDeposit', event.target.value)} placeholder="כמה מוסיפים כל חודש" /></LabeledField>
                      </div>
                      <div className="mt-4 flex justify-between text-sm font-semibold"><span>{progress}%</span><button onClick={() => removeRow('savingGoals', goal.id)} className="text-neutral-700">מחיקה</button></div>
                      <div className="mt-3 rounded-2xl border border-neutral-200 bg-white p-3 text-sm leading-7 text-neutral-600"><div>ETA ליעד: <strong>{Number.isFinite(etaMonths) ? `${etaMonths} חודשים` : 'לא מוגדר'}</strong></div><div className="mt-1">אם תגדילו ב־₪500 בחודש תגיעו בערך תוך <strong>{Number.isFinite(boostedEta) ? `${boostedEta} חודשים` : '—'}</strong>.</div></div>
                      <div className="mt-2 h-3 overflow-hidden rounded-full bg-white"><div className="h-full rounded-full" style={{ width: `${progress}%`, backgroundColor: activeTheme.accent }} /></div>
                    </div>
                  );
                })}
              </div>
            </Section>

            <Section><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">קרן חירום</h2><p className="mt-2 text-sm text-neutral-500">מלאו סכום חיסכון נזיל נוכחי.</p><Field type="number" value={monthData.emergencyFund} onChange={(event) => updateMonthField('emergencyFund', toNumber(event.target.value))} className="mt-6 w-full text-xl font-semibold" /></Section>
          </>
        ) : null}

        {activeTab === 'income' ? (
          <section className="grid gap-6 lg:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2">
            <Section>
              <div className="flex items-center justify-between gap-4"><div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">הכנסות</h2><p className="mt-2 text-sm text-neutral-500">אפשר לייבא הכנסות מ־PDF תלוש, CSV או Excel. אם ה־PDF טקסטואלי, המערכת תנסה לזהות נטו לתשלום אוטומטית.</p></div><div className="flex flex-wrap gap-3"><label className="cursor-pointer rounded-xl bg-neutral-950 px-4 py-3 text-sm font-semibold text-white transition hover:bg-neutral-800">ייבוא הכנסות PDF/CSV/Excel<input type="file" accept="application/pdf,.csv,.xlsx,.xls" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) importIncomeFile(file); }} /></label><label className="cursor-pointer rounded-xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-sm font-semibold text-neutral-700 transition hover:border-neutral-400 hover:bg-white">צירוף תלוש PDF<input type="file" accept="application/pdf" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) attachSalarySlipFile(file); }} /></label><PrimaryButton theme={activeTheme} onClick={addIncome}>+ הוספה</PrimaryButton></div></div>
              {(monthData.attachedDocuments || []).length ? <div className="mt-4 space-y-2 rounded-2xl bg-neutral-50 p-4 text-sm text-neutral-600">{monthData.attachedDocuments.map((document) => <div key={document.id} className="flex items-center justify-between gap-3"><span>תלוש מצורף: <strong>{document.name}</strong></span><button type="button" onClick={() => removeAttachedDocument(document.id)} className="font-semibold text-neutral-700">הסרה</button></div>)}</div> : null}
              <div className="mt-6 space-y-3">{monthData.incomes.map((income) => <InputRow key={income.id}><Field value={income.name} onChange={(event) => updateRow('incomes', income.id, 'name', event.target.value)} /><Field type="number" value={income.amount} onChange={(event) => updateRow('incomes', income.id, 'amount', event.target.value)} /><GhostButton onClick={() => removeRow('incomes', income.id)} className="px-0">×</GhostButton></InputRow>)}</div>
            </Section>

            <Section>
              <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">הכנסות אורן / עצמאי</h2><p className="mt-2 text-sm text-neutral-500">אם העסק נפרד, מזינים כאן רק את הסכום שאורן מעביר לחשבון המשותף כהכנסה. אפשר להדליק מצב עצמאי רק אם רוצים לכלול מע״מ, מס וביטוח לאומי בתזרים הבית.</p>
              <div className="mt-5 rounded-[24px] border border-neutral-200 bg-neutral-50 p-4">
                <label className="flex items-center justify-between gap-4 text-sm font-semibold text-neutral-700">
                  <span>לכלול עצמאי בתזרים המשפחתי</span>
                  <input type="checkbox" checked={includeSelfEmployed} onChange={(event) => updatePreference('includeSelfEmployed', event.target.checked)} className="h-5 w-5" style={{ accentColor: activeTheme.accent }} />
                </label>
                <p className="mt-3 text-xs leading-6 text-neutral-500">כבוי: העסק נשאר מחוץ לדשבורד, ורק העברה/משכורת לחשבון המשותף נספרת כהכנסה. דולק: תשלומי עצמאי נספרים כהוצאות בית.</p>
              </div>
              <div className="mt-6 grid gap-4 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2">{[['owner', 'בעל העסק', 'text'], ['salaryTransferToHousehold', 'העברה / משכורת לחשבון המשותף', 'number'], ['grossRevenue', 'הכנסה עסקית ברוטו', 'number'], ['vatCollected', 'מע״מ שנגבה מלקוחות', 'number'], ['vatPaidOnExpenses', 'מע״מ על הוצאות מוכרות', 'number'], ['incomeTaxAdvance', 'מקדמת מס הכנסה', 'number'], ['nationalInsurance', 'ביטוח לאומי', 'number'], ['businessExpenses', 'הוצאות עסקיות ששולמו החודש', 'number']].map(([field, label, type]) => <label key={field} className="text-sm font-semibold text-neutral-600">{label}<Field type={type} value={monthData.selfEmployed[field]} onChange={(event) => updateSelfEmployedField(field, event.target.value)} className="mt-2 w-full" /></label>)}</div>
              <div className="mt-6 grid gap-4 md:grid-cols-3"><StatCard title="מע״מ צפוי" value={SHEKEL.format(selfEmployedVatDue)} note="נגבה פחות מוכר" /><StatCard title="מס + ביטוח" value={SHEKEL.format(toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance))} note="תשלומי חובה" /><StatCard title="סה״כ עצמאי" value={SHEKEL.format(totalSelfEmployedPayments)} note={includeSelfEmployed ? 'כלול בבית' : 'מחוץ לבית'} /></div>
            </Section>
          </section>
        ) : null}

        {activeTab === 'insights' ? (
          <section className="grid gap-6 lg:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2">
            {preferences.showSmartInsightCards ? <Section><div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between"><div><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">תובנות חכמות</h2><p className="mt-2 text-sm text-neutral-500">תובנות מחושבות ישירות מהנתונים: חריגות, תקציבים, בתי עסק מובילים, חיובים חוזרים ודפוסים חודשיים.</p></div><div className="rounded-full px-4 py-2 text-sm font-semibold" style={{ backgroundColor: activeTheme.soft, color: activeTheme.text }}>מתעדכן אוטומטית</div></div><div className="mt-5 grid gap-4">{realInsights.map((insight, index) => <div key={insight} className="flex items-start gap-4 rounded-[24px] border border-neutral-200 bg-white p-5 shadow-sm"><div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-neutral-100 text-lg font-semibold text-neutral-500">{index % 3 === 0 ? '◔' : index % 3 === 1 ? '▲' : '✦'}</div><div className="flex-1 text-sm leading-7 text-neutral-700 no-orphans">{noSingleWordLine(insight)}</div></div>)}</div></Section> : null}
            {preferences.showRecurringDetection ? <Section><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">זיהוי חיובים קבועים</h2><p className="mt-2 text-sm text-neutral-500">זיהוי מנויים, ביטוחים, סלולר ושכירות לפי מילות מפתח וחזרה בין חודשים.</p><div className="mt-5 space-y-3">{recurringTransactions.length ? recurringTransactions.map((item) => <div key={item.id} className="flex justify-between rounded-2xl bg-neutral-50 p-4 text-sm"><span>{item.merchant}</span><strong>{SHEKEL.format(item.amount)}</strong></div>) : <EmptyState title="אין עדיין חיובים קבועים" text="העלי פירוטים של כמה חודשים כדי שנוכל לזהות מנויים ותשלומים חוזרים בצורה חכמה." />}</div></Section> : null}
          </section>
        ) : null}

        {activeTab === 'settings' ? (
          <>
            <Section>
              <div className="flex flex-col gap-2">
                <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">התאמה אישית</h2>
                <p className="text-sm leading-7 text-neutral-500">
                  כאן מגדירים איך הטופס והדשבורד יתנהגו: שמות, יעדים ומה יוצג במסך הראשי.
                </p>
              </div>

              <div className="mt-6 grid gap-6 lg:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2">
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">פרטי הבית</h3>
                  <div className="mt-4 grid gap-3">
                    <label className="text-sm font-semibold text-neutral-600">
                      שם הדשבורד
                      <Field
                        value={monthData.dashboardTitle}
                        onChange={(event) => updateMonthField('dashboardTitle', event.target.value)}
                        className="mt-2 w-full"
                      />
                    </label>
                    <label className="text-sm font-semibold text-neutral-600">
                      מזהה בית / Household ID
                      <Field
                        value={householdProfileId}
                        onChange={(event) => updatePreference('householdProfileId', event.target.value || DEFAULT_SUPABASE_PROFILE_ID)}
                        className="mt-2 w-full"
                      />
                    </label>
                    <div className="grid gap-3 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2">
                      <label className="text-sm font-semibold text-neutral-600">
                        משתמש/ת ראשון/ה
                        <Field
                          value={preferences.primaryPerson}
                          onChange={(event) => updatePreference('primaryPerson', event.target.value)}
                          className="mt-2 w-full"
                        />
                      </label>
                      <label className="text-sm font-semibold text-neutral-600">
                        משתמש/ת שני/ה
                        <Field
                          value={preferences.secondaryPerson}
                          onChange={(event) => updatePreference('secondaryPerson', event.target.value)}
                          className="mt-2 w-full"
                        />
                      </label>
                    </div>
                    <div className="mt-3 flex flex-wrap gap-3 border-t border-neutral-200 pt-5">
                      <PrimaryButton
                        theme={activeTheme}
                        onClick={async () => {
                          try {
                            setCloudStatus('שומר...');
                            await saveFinanceStateToSupabase(months, learnedRules, preferences, householdProfileId, supabaseConfig);
                            setCloudStatus('נשמר בענן');
                          } catch (error) {
                            setCloudStatus(`שגיאה בשמירה: ${error?.message || 'לא ידוע'}`);
                          }
                        }}
                      >
                        שמור הגדרות בענן
                      </PrimaryButton>
                      <GhostButton onClick={() => window.location.reload()}>
                        רענן חיבור
                      </GhostButton>
                    </div>
                  </div>
                </div>

                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">יעדים חודשיים</h3>
                  <div className="mt-4 grid gap-3 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2">
                    <label className="text-sm font-semibold text-neutral-600">
                      יעד הוצאות חודשי
                      <Field
                        type="number"
                        value={preferences.monthlyBudgetTarget}
                        onChange={(event) => updatePreference('monthlyBudgetTarget', event.target.value)}
                        className="mt-2 w-full"
                      />
                    </label>
                    <label className="text-sm font-semibold text-neutral-600">
                      יעד שיעור חיסכון באחוזים
                      <Field
                        type="number"
                        value={preferences.savingsRateTarget}
                        onChange={(event) => updatePreference('savingsRateTarget', event.target.value)}
                        className="mt-2 w-full"
                      />
                    </label>
                  </div>
                </div>
              </div>

              <div className="mt-6 rounded-[24px] border border-neutral-200 bg-white p-5">
                <h3 className="text-lg font-semibold text-neutral-950">Home Widgets</h3>
                <div className="mt-4 grid gap-3 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2 lg:grid-cols-3">
                  {[
                    ['showMonthlyStory', 'Monthly Story'],
                    ['showFinancialHealth', 'Financial Health'],
                    ['showCategoryChart', 'גרף קטגוריות'],
                    ['showTrendChart', 'גרף מגמה'],
                    ['showSmartInsightCards', 'כרטיסי תובנות'],
                    ['showRecurringDetection', 'זיהוי חיובים קבועים'],
                  ].map(([field, label]) => (
                    <label key={field} className="flex cursor-pointer items-center justify-between gap-4 rounded-2xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-sm font-semibold text-neutral-700">
                      <span>{label}</span>
                      <input
                        type="checkbox"
                        checked={Boolean(preferences[field])}
                        onChange={(event) => updatePreference(field, event.target.checked)}
                        className="h-5 w-5"
                        style={{ accentColor: activeTheme.accent }}
                      />
                    </label>
                  ))}
                </div>
              </div>

              <div className="mt-6 grid gap-6 lg:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 grid-cols-1 md:grid-cols-2">
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">Theme Mood</h3>
                  <div className="mt-4 grid grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 md:grid-cols-2 gap-3">
                    {Object.keys(THEME_STYLES).map((themeName) => {
                      const themeStyle = getSafeTheme(themeName);
                      const isSelected = preferences.themeMood === themeName;
                      return (
                        <button
                          key={themeName}
                          type="button"
                          onClick={() => updatePreference('themeMood', themeName)}
                          className="rounded-2xl border px-4 py-4 text-sm font-semibold transition"
                          style={isSelected ? { borderColor: themeStyle.accent, backgroundColor: themeStyle.soft, color: themeStyle.text } : undefined}
                        >
                          {themeName}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">Financial Operating Mode</h3>
                  <div className="mt-4 space-y-3">
                    {['Survival', 'Stable', 'Growth', 'Wealth Building'].map((mode) => {
                      const config = getFinancialModeConfig(mode);
                      const isSelected = preferences.financialMode === mode;
                      return (
                        <button
                          key={mode}
                          type="button"
                          onClick={() => updatePreference('financialMode', mode)}
                          className="w-full rounded-2xl border px-4 py-4 text-right transition"
                          style={isSelected ? { borderColor: activeTheme.accent, backgroundColor: activeTheme.soft } : undefined}
                        >
                          <div className="font-semibold text-neutral-900">{mode}</div>
                          <div className="mt-1 text-sm text-neutral-500">{operatingModeMessages[mode]}</div>
                          <div className="mt-3 grid gap-2 text-xs text-neutral-500 md:grid-cols-3">
                            <span>יעד חיסכון {formatPercent(config.savingsTarget)}</span>
                            <span>התראה ב־{config.budgetWarningAt}%</span>
                            <span>{config.notificationTone}</span>
                          </div>
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>

              <div className="mt-6 rounded-[24px] border border-neutral-200 bg-white p-5">
                <h3 className="text-lg font-semibold text-neutral-950">Production Setup Health</h3>
                <div className="mt-4 grid gap-3 md:grid-cols-4">
                  <div className="rounded-2xl bg-neutral-50 p-4 text-sm">
                    <strong>LocalStorage</strong>
                    <div className="mt-1 text-neutral-500">{setupHealth.localStorage ? 'פעיל' : 'לא זמין'}</div>
                  </div>
                  <div className="rounded-2xl bg-neutral-50 p-4 text-sm">
                    <strong>Supabase ENV</strong>
                    <div className="mt-1 text-neutral-500">{setupHealth.supabaseEnv ? 'מוגדר' : 'לא מוגדר'}</div>
                  </div>
                  <div className="rounded-2xl bg-neutral-50 p-4 text-sm">
                    <strong>Excel Parser</strong>
                    <div className="mt-1 text-neutral-500">{setupHealth.xlsxParser ? 'פעיל' : 'חסר xlsx'}</div>
                  </div>
                  <div className="rounded-2xl bg-neutral-50 p-4 text-sm">
                    <strong>Household</strong>
                    <div className="mt-1 text-neutral-500">{setupHealth.householdProfileId}</div>
                  </div>
                </div>
                <p className="mt-4 text-xs leading-6 text-neutral-500">
                  זה לא דמו: כל סטטוס כאן משקף חיבור אמיתי בקוד. אם Supabase ENV לא מוגדר, המערכת עובדת במצב LocalStorage + JSON Backup.
                </p>
              </div>

              <div className="mt-6 grid gap-6 lg:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 md:grid-cols-2">
                <div className="rounded-[24px] border border-neutral-200 bg-white p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">Smart Notifications</h3>
                  <div className="mt-4 space-y-3">
                    {[
                      ['budget80', 'התראה לפי מצב פיננסי'],
                      ['woltSpike', 'התראה כשוולט עולה משמעותית'],
                      ['savingsDrop', 'התראה כששיעור החיסכון יורד'],
                    ].map(([field, label]) => (
                      <label key={field} className="flex items-center justify-between rounded-2xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-sm font-semibold text-neutral-700">
                        <span>{label}</span>
                        <input
                          type="checkbox"
                          checked={Boolean(preferences.notifications?.[field])}
                          onChange={(event) => updatePreference('notifications', { ...preferences.notifications, [field]: event.target.checked })}
                          className="h-5 w-5"
                          style={{ accentColor: activeTheme.accent }}
                        />
                      </label>
                    ))}
                  </div>
                </div>

                <div className="rounded-[24px] border border-neutral-200 bg-white p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">Privacy & Sync</h3>
                  <div className="mt-4 grid gap-3">
                    {['Cloud Sync', 'Local Only', 'Auto Backup'].map((mode) => {
                      const isSelected = preferences.syncMode === mode;
                      return (
                        <button
                          key={mode}
                          type="button"
                          onClick={() => updatePreference('syncMode', mode)}
                          className="rounded-2xl border px-4 py-4 text-right text-sm font-semibold transition"
                          style={isSelected ? { borderColor: activeTheme.accent, backgroundColor: activeTheme.soft, color: activeTheme.text } : undefined}
                        >
                          {mode}
                        </button>
                      );
                    })}
                  </div>
                  <div className="mt-5 grid gap-3 md:grid-cols-3">
                    <PrimaryButton theme={activeTheme} onClick={exportBackup}>ייצוא גיבוי JSON</PrimaryButton>
                    <label className="cursor-pointer rounded-xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-center text-sm font-semibold text-neutral-700">
                      ייבוא גיבוי
                      <input
                        type="file"
                        accept="application/json"
                        className="hidden"
                        onChange={(event) => {
                          const file = event.target.files?.[0];
                          if (file) importBackupFile(file);
                        }}
                      />
                    </label>
                    <GhostButton onClick={resetCurrentMonth}>איפוס חודש נוכחי</GhostButton>
                  </div>
                  <p className="mt-4 text-xs leading-6 text-neutral-500">
                    Cloud Sync עובד רק אם Supabase מוגדר. Local Only שומר בדפדפן. ייצוא/ייבוא JSON עובד תמיד.
                  </p>
                </div>
              </div>
            </Section>

            <Section>
              <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">חיפוש ופילטרים</h2>
              <div className="mt-5 grid gap-3">
                <Field value={searchTerm} onChange={(event) => setSearchTerm(event.target.value)} placeholder="חיפוש בית עסק, למשל וולט" />
                <SelectField value={categoryFilter} onChange={(event) => setCategoryFilter(event.target.value)}>
                  <option>הכול</option>
                  {EXPENSE_CATEGORIES.map((category) => <option key={category}>{category}</option>)}
                </SelectField>
                <div className="grid gap-3 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 md:grid-cols-1 grid-cols-1 md:grid-cols-2">
                  <Field value={minAmount} onChange={(event) => setMinAmount(event.target.value)} type="number" placeholder="סכום מינימום" />
                  <Field value={maxAmount} onChange={(event) => setMaxAmount(event.target.value)} type="number" placeholder="סכום מקסימום" />
                </div>
              </div>
            </Section>
          </>
        ) : null}
      </div>
    </div>
  );
}
