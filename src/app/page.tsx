// @ts-nocheck

'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

// LocalStorage keys are versioned so structural changes do not collide with older saved data.
// Monthly data is stored separately from global settings so switching months does not reset Supabase, theme, users, or dashboard preferences.
const STORAGE_KEY = 'family-finance-os-stable-v14';
const SETTINGS_STORAGE_KEY = 'family-finance-os-global-settings-v1';
const AUTH_STORAGE_KEY = 'family-finance-os-auth-session-v1';
const DEFAULT_SUPABASE_PROFILE_ID = 'default-household';
const APP_BUILD_MARKER = 'finance-dashboard-build-v14';

// Monthly compare periods determine how many previous months are averaged against the current month.
// Controls the Monthly Compare ranges. Each option compares the current month against an average of earlier months.
const COMPARE_PERIODS = [
  { id: 'previous', label: '◊ó◊ï◊ì◊© ◊ß◊ï◊ì◊ù', months: 1 },
  { id: 'quarter', label: '3 ◊ó◊ï◊ì◊©◊ô◊ù', months: 3 },
  { id: 'halfYear', label: '6 ◊ó◊ï◊ì◊©◊ô◊ù', months: 6 },
  { id: 'year', label: '◊©◊†◊î', months: 12 },
  { id: 'all', label: '◊õ◊ú ◊î◊™◊ß◊ï◊§◊î', months: Infinity },
];

const TABS = [
  { id: 'dashboard', label: 'Dashboard' },
  { id: 'income', label: '◊î◊õ◊†◊°◊ï◊™' },
  { id: 'credit', label: '◊î◊ï◊¶◊ê◊ï◊™' },
  { id: 'savings', label: '◊ó◊ô◊°◊õ◊ï◊ü' },
  { id: 'insights', label: '◊™◊ï◊ë◊†◊ï◊™ ◊ó◊õ◊û◊ï◊™' },
  { id: 'settings', label: '◊î◊í◊ì◊®◊ï◊™' },
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
    focus: '◊ß◊ô◊¶◊ï◊• ◊î◊ï◊¶◊ê◊ï◊™ ◊ï◊©◊û◊ô◊®◊î ◊¢◊ú ◊™◊ñ◊®◊ô◊ù ◊ó◊ô◊ï◊ë◊ô',
    priorityMetric: 'burnRate',
    notificationTone: '◊ê◊í◊®◊°◊ô◊ë◊ô',
  },
  Stable: {
    label: 'Stable',
    savingsTarget: 20,
    budgetWarningAt: 80,
    strictness: 1,
    focus: '◊ê◊ô◊ñ◊ï◊ü ◊ë◊ô◊ü ◊ê◊ô◊õ◊ï◊™ ◊ó◊ô◊ô◊ù ◊ú◊ó◊ô◊°◊õ◊ï◊ü ◊ô◊¶◊ô◊ë',
    priorityMetric: 'savingsRate',
    notificationTone: '◊û◊ê◊ï◊ñ◊ü',
  },
  Growth: {
    label: 'Growth',
    savingsTarget: 25,
    budgetWarningAt: 90,
    strictness: 0.85,
    focus: '◊î◊í◊ì◊ú◊™ ◊î◊õ◊†◊°◊ï◊™, ◊î◊©◊ß◊¢◊î ◊ë◊¶◊û◊ô◊ó◊î ◊ï◊©◊ô◊§◊ï◊® Cash Flow',
    priorityMetric: 'cashFlow',
    notificationTone: '◊¶◊û◊ô◊ó◊î',
  },
  'Wealth Building': {
    label: 'Wealth Building',
    savingsTarget: 35,
    budgetWarningAt: 95,
    strictness: 0.75,
    focus: '◊ë◊†◊ô◊ô◊™ ◊î◊ï◊ü, ◊î◊í◊ì◊ú◊™ ◊†◊õ◊°◊ô◊ù ◊ï◊ê◊ï◊§◊ò◊ô◊û◊ô◊ñ◊¶◊ô◊î ◊§◊ô◊†◊†◊°◊ô◊™',
    priorityMetric: 'netWorth',
    notificationTone: '◊ê◊ï◊§◊ò◊ô◊û◊ô◊ñ◊¶◊ô◊î',
  },
};

const EXPENSE_CATEGORIES = [
  '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î',
  '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î',
  '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò',
  '◊™◊ó◊ë◊ï◊®◊î ◊ï◊®◊õ◊ë◊ô◊ù',
  '◊ò◊ô◊°◊ï◊™ ◊ï◊™◊ô◊ô◊®◊ï◊™',
  '◊ë◊®◊ô◊ê◊ï◊™ ◊ï◊®◊§◊ï◊ê◊î',
  '◊§◊ê◊®◊ù ◊ï◊ß◊ï◊°◊û◊ò◊ô◊ß◊î',
  '◊ê◊ï◊§◊†◊î ◊ï◊î◊ú◊ë◊©◊î',
  '◊î◊¢◊ë◊®◊™ ◊õ◊°◊§◊ô◊ù',
  '◊û◊©◊ô◊õ◊™ ◊û◊ñ◊ï◊û◊ü',
  '◊°◊§◊®◊ô◊ù ◊ï◊ì◊§◊ï◊°',
  '◊ë◊ô◊ò◊ï◊ó◊ô◊ù',
  '◊û◊ô◊°◊ô◊ù ◊ï◊™◊©◊ú◊ï◊û◊ô◊ù',
  '◊ì◊ô◊ï◊® ◊ï◊ó◊©◊ë◊ï◊†◊ï◊™',
  '◊ó◊ô◊°◊õ◊ï◊ü ◊ï◊î◊©◊ß◊¢◊ï◊™',
  '◊î◊ï◊¶◊ê◊ï◊™ ◊¢◊°◊ß◊ô◊ï◊™',
  '◊©◊ï◊†◊ï◊™',
  '◊ê◊ó◊®',
];

const MAX_CATEGORY_MAP = {
  '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î': '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î',
  '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î': '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î',
  '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò': '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò',
  '◊™◊ó◊ë◊ï◊®◊î ◊ï◊®◊õ◊ë◊ô◊ù': '◊™◊ó◊ë◊ï◊®◊î ◊ï◊®◊õ◊ë◊ô◊ù',
  '◊ò◊ô◊°◊ï◊™ ◊ï◊™◊ô◊ô◊®◊ï◊™': '◊ò◊ô◊°◊ï◊™ ◊ï◊™◊ô◊ô◊®◊ï◊™',
  '◊ë◊®◊ô◊ê◊ï◊™ ◊ï◊®◊§◊ï◊ê◊î': '◊ë◊®◊ô◊ê◊ï◊™ ◊ï◊®◊§◊ï◊ê◊î',
  '◊§◊ê◊®◊ù ◊ï◊ß◊ï◊°◊û◊ò◊ô◊ß◊î': '◊§◊ê◊®◊ù ◊ï◊ß◊ï◊°◊û◊ò◊ô◊ß◊î',
  '◊ê◊ï◊§◊†◊î ◊ï◊î◊ú◊ë◊©◊î': '◊ê◊ï◊§◊†◊î ◊ï◊î◊ú◊ë◊©◊î',
  '◊î◊¢◊ë◊®◊™ ◊õ◊°◊§◊ô◊ù': '◊î◊¢◊ë◊®◊™ ◊õ◊°◊§◊ô◊ù',
  '◊û◊©◊ô◊õ◊™ ◊û◊ñ◊ï◊û◊ü': '◊û◊©◊ô◊õ◊™ ◊û◊ñ◊ï◊û◊ü',
  '◊°◊§◊®◊ô◊ù ◊ï◊ì◊§◊ï◊°': '◊°◊§◊®◊ô◊ù ◊ï◊ì◊§◊ï◊°',
  '◊©◊ï◊†◊ï◊™': '◊©◊ï◊†◊ï◊™',
};

const CATEGORY_BUDGETS = {
  '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î': 4000,
  '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î': 800,
  '◊™◊ó◊ë◊ï◊®◊î ◊ï◊®◊õ◊ë◊ô◊ù': 1800,
  '◊ê◊ï◊§◊†◊î ◊ï◊î◊ú◊ë◊©◊î': 1200,
  '◊ë◊®◊ô◊ê◊ï◊™ ◊ï◊®◊§◊ï◊ê◊î': 800,
  '◊§◊ê◊®◊ù ◊ï◊ß◊ï◊°◊û◊ò◊ô◊ß◊î': 700,
  '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò': 600,
  '◊ò◊ô◊°◊ï◊™ ◊ï◊™◊ô◊ô◊®◊ï◊™': 1500,
  '◊î◊¢◊ë◊®◊™ ◊õ◊°◊§◊ô◊ù': 1000,
  '◊©◊ï◊†◊ï◊™': 1000,
  ◊ê◊ó◊®: 1000,
};

// Merchant keywords are intentionally simple and editable: user corrections are saved in learnedRules.
// First-pass categorization rules for imported card transactions. User edits later become learnedRules.
const MERCHANT_CATEGORY_MAP = {
  wolt: '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î',
  tenbis: '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î',
  shufersal: '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î',
  ◊©◊ï◊§◊®◊°◊ú: '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î',
  ◊®◊û◊ô: '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î',
  victory: '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î',
  ◊ï◊ô◊ß◊ò◊ï◊®◊ô: '◊û◊ñ◊ï◊ü ◊ï◊¶◊®◊ô◊õ◊î',
  yellow: '◊™◊ó◊ë◊ï◊®◊î ◊ï◊®◊õ◊ë◊ô◊ù',
  ◊ì◊ï◊®: '◊™◊ó◊ë◊ï◊®◊î ◊ï◊®◊õ◊ë◊ô◊ù',
  ◊§◊ñ: '◊™◊ó◊ë◊ï◊®◊î ◊ï◊®◊õ◊ë◊ô◊ù',
  fox: '◊ê◊ï◊§◊†◊î ◊ï◊î◊ú◊ë◊©◊î',
  zara: '◊ê◊ï◊§◊†◊î ◊ï◊î◊ú◊ë◊©◊î',
  superpharm: '◊§◊ê◊®◊ù ◊ï◊ß◊ï◊°◊û◊ò◊ô◊ß◊î',
  ◊°◊ï◊§◊®◊§◊ê◊®◊ù: '◊§◊ê◊®◊ù ◊ï◊ß◊ï◊°◊û◊ò◊ô◊ß◊î',
  ◊õ◊ú◊ú◊ô◊™: '◊ë◊®◊ô◊ê◊ï◊™ ◊ï◊®◊§◊ï◊ê◊î',
  netflix: '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò',
  spotify: '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò',
  icloud: '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò',
  google: '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò',
  apple: '◊§◊†◊ê◊ô, ◊ë◊ô◊ì◊ï◊® ◊ï◊°◊§◊ï◊®◊ò',
};

const RECURRING_KEYWORDS = [
  'netflix', 'spotify', 'icloud', 'google', 'apple', 'cellcom', 'partner', 'pelephone', 'hot', 'yes',
  '◊ë◊ô◊ò◊ï◊ó', '◊î◊®◊ê◊ú', '◊û◊í◊ì◊ú', '◊õ◊ú◊ú', '◊°◊ú◊ß◊ï◊ù', '◊§◊®◊ò◊†◊®', '◊§◊ú◊ê◊§◊ï◊ü', '◊©◊õ◊ô◊®◊ï◊™',
];

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
    .split('‚Ç™').join('')
    .split('(').join('')
    .split(')').join('')
    .split(' ').join('')
    .split(String.fromCharCode(160)).join('')
    .split('‚àí').join('-')
    .split('‚Äì').join('-')
    .split('‚Äî').join('-')
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
  return '◊ê◊ó◊®';
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
  const headerIndex = findHeaderIndex(headers, ['amount', '◊°◊õ◊ï◊ù', '◊ó◊ô◊ï◊ë', '◊ó◊ï◊ë◊î', '◊ñ◊õ◊ï◊™', '◊¢◊°◊ß◊î', 'debit', 'credit', 'charge', 'total', '◊†◊ò◊ï', '◊ú◊™◊©◊ú◊ï◊ù'], -1);
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
function normalizeImportedRows(rows, learnedRules = {}) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const cleanedRows = rows.map((row) => (Array.isArray(row) ? row.map((cell) => String(cell || '').trim()) : [])).filter((row) => row.some(Boolean));
  if (!cleanedRows.length) return [];

  const headerCandidates = cleanedRows.slice(0, 25);
  const headerRowIndex = headerCandidates.findIndex((row) => {
    const joined = row.join(' ').toLowerCase();
    return ['date', '◊™◊ê◊®◊ô◊ö', 'amount', '◊°◊õ◊ï◊ù', 'merchant', '◊ë◊ô◊™ ◊¢◊°◊ß', '◊©◊ù ◊ë◊ô◊™ ◊î◊¢◊°◊ß', '◊™◊ô◊ê◊ï◊®', '◊§◊ô◊®◊ï◊ò', '◊ó◊ô◊ï◊ë', '◊ñ◊õ◊ï◊™', '◊ó◊ï◊ë◊î'].some((word) => joined.includes(word));
  });
  const hasHeader = headerRowIndex >= 0;
  const headers = hasHeader ? cleanedRows[headerRowIndex] : cleanedRows[0] || [];
  const dataRows = hasHeader ? cleanedRows.slice(headerRowIndex + 1) : cleanedRows;
  const sampleRows = dataRows.slice(0, 30);

  const dateIndex = hasHeader ? findHeaderIndex(headers, ['date', '◊™◊ê◊®◊ô◊ö', '◊™◊ê◊®◊ô◊ö ◊¢◊°◊ß◊î', '◊™◊ê◊®◊ô◊ö ◊®◊õ◊ô◊©◊î', '◊™◊ê◊®◊ô◊ö ◊ó◊ô◊ï◊ë'], 0) : 0;
  const merchantIndex = hasHeader ? findHeaderIndex(headers, ['merchant', '◊ë◊ô◊™ ◊¢◊°◊ß', '◊©◊ù ◊ë◊ô◊™ ◊î◊¢◊°◊ß', '◊©◊ù ◊ë◊ô◊™ ◊¢◊°◊ß', '◊°◊§◊ß', '◊™◊ô◊ê◊ï◊®', '◊§◊ô◊®◊ï◊ò', '◊©◊ù', '◊§◊®◊ò◊ô◊ù'], 1) : 1;
  const importedCategoryIndex = hasHeader ? findHeaderIndex(headers, ['◊ß◊ò◊í◊ï◊®◊ô◊î', 'category'], -1) : -1;
  const amountIndex = hasHeader ? findHeaderIndex(headers, ['◊°◊õ◊ï◊ù ◊ó◊ô◊ï◊ë', 'amount charged', '◊ó◊ô◊ï◊ë', '◊°◊õ◊ï◊ù', '◊ó◊ï◊ë◊î', '◊ñ◊õ◊ï◊™', 'amount', 'charge', 'total'], -1) : findAmountIndex(headers, sampleRows);
  const finalAmountIndex = amountIndex >= 0 ? amountIndex : findAmountIndex(headers, sampleRows);

  return dataRows
    .map((row) => {
      const amountCell = finalAmountIndex >= 0 ? row[finalAmountIndex] : [...row].reverse().find((cell) => Math.abs(toNumber(cell)) > 0);
      const amount = Math.abs(toNumber(amountCell));
      const date = row[dateIndex] || row.find((cell) => String(cell || '').includes('/')) || row.find((cell) => String(cell || '').includes('-')) || '';
      const merchant = row[merchantIndex] || row.find((cell, index) => index !== dateIndex && index !== finalAmountIndex && String(cell || '').trim() && Math.abs(toNumber(cell)) === 0) || '◊¢◊°◊ß◊î';
      const importedCategory = importedCategoryIndex >= 0 ? row[importedCategoryIndex] : '';
      const normalizedMerchant = normalizeMerchantName(merchant);
      const isSummaryRow = normalizedMerchant.includes('◊°◊ö ◊î◊õ◊ú') || normalizedMerchant.includes('total') || normalizedMerchant.includes('◊°◊î◊õ');
      return {
        id: makeId('tx'),
        date,
        merchant,
        amount,
        category: detectCategory(merchant, learnedRules, importedCategory),
      };
    })
    .filter((transaction) => transaction.amount > 0 && normalizeMerchantName(transaction.merchant) !== normalizeMerchantName('◊¢◊°◊ß◊î') && !normalizeMerchantName(transaction.merchant).includes('◊°◊ö ◊î◊õ◊ú'));
}

function normalizeBankRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const cleanedRows = rows.map((row) => (Array.isArray(row) ? row.map((cell) => String(cell || '').trim()) : [])).filter((row) => row.some(Boolean));
  if (!cleanedRows.length) return [];

  const headerCandidates = cleanedRows.slice(0, 25);
  const headerRowIndex = headerCandidates.findIndex((row) => {
    const joined = row.join(' ').toLowerCase();
    return ['◊™◊ê◊®◊ô◊ö', 'date', '◊™◊ô◊ê◊ï◊®', '◊§◊ô◊®◊ï◊ò', '◊ê◊°◊û◊õ◊™◊ê', '◊ó◊ï◊ë◊î', '◊ñ◊õ◊ï◊™', '◊ô◊™◊®◊î', 'balance', 'debit', 'credit'].some((word) => joined.includes(word));
  });
  const hasHeader = headerRowIndex >= 0;
  const headers = hasHeader ? cleanedRows[headerRowIndex] : cleanedRows[0] || [];
  const dataRows = hasHeader ? cleanedRows.slice(headerRowIndex + 1) : cleanedRows;
  const sampleRows = dataRows.slice(0, 40);

  const dateIndex = hasHeader ? findHeaderIndex(headers, ['◊™◊ê◊®◊ô◊ö', 'date', '◊™◊ê◊®◊ô◊ö ◊§◊¢◊ï◊ú◊î', '◊™◊ê◊®◊ô◊ö ◊¢◊®◊ö'], 0) : 0;
  const descriptionIndex = hasHeader ? findHeaderIndex(headers, ['◊™◊ô◊ê◊ï◊®', '◊§◊ô◊®◊ï◊ò', '◊§◊®◊ò◊ô◊ù', '◊©◊ù', '◊§◊¢◊ï◊ú◊î', 'description'], 1) : 1;
  const debitIndex = hasHeader ? findHeaderIndex(headers, ['◊ó◊ï◊ë◊î', 'debit', '◊ó◊ô◊ï◊ë', '◊û◊©◊ô◊õ◊î'], -1) : -1;
  const creditIndex = hasHeader ? findHeaderIndex(headers, ['◊ñ◊õ◊ï◊™', 'credit', '◊î◊§◊ß◊ì◊î'], -1) : -1;
  const balanceIndex = hasHeader ? findHeaderIndex(headers, ['◊ô◊™◊®◊î', 'balance', '◊ô◊™◊®◊î ◊ë◊©◊ó', '◊ô◊™◊®◊î ◊†◊ï◊õ◊ó◊ô◊™'], -1) : -1;
  const amountIndex = debitIndex < 0 && creditIndex < 0 ? findAmountIndex(headers, sampleRows) : -1;

  return dataRows
    .map((row) => {
      const debit = debitIndex >= 0 ? Math.abs(toNumber(row[debitIndex])) : 0;
      const credit = creditIndex >= 0 ? Math.abs(toNumber(row[creditIndex])) : 0;
      const fallbackAmount = amountIndex >= 0 ? toNumber(row[amountIndex]) : 0;
      const amount = credit || debit ? credit - debit : fallbackAmount;
      const balance = balanceIndex >= 0 ? toNumber(row[balanceIndex]) : 0;
      const description = row[descriptionIndex] || row.find((cell, index) => index !== dateIndex && index !== debitIndex && index !== creditIndex && index !== balanceIndex && String(cell || '').trim() && Math.abs(toNumber(cell)) === 0) || '◊™◊†◊ï◊¢◊î ◊ë◊ë◊†◊ß';
      const date = row[dateIndex] || row.find((cell) => String(cell || '').includes('/') || String(cell || '').includes('-')) || '';
      const normalizedDescription = normalizeMerchantName(description);
      const isSummaryRow = normalizedDescription.includes('◊°◊ö ◊î◊õ◊ú') || normalizedDescription.includes('◊°◊î◊õ') || normalizedDescription.includes('total');
      return { id: makeId('banktx'), date, description, amount, debit, credit, balance };
    })
    .filter((transaction) => Math.abs(toNumber(transaction.amount)) > 0 && !normalizeMerchantName(transaction.description).includes('◊°◊ö ◊î◊õ◊ú'));
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

function parseCsvText(text, learnedRules = {}) {
  const carriageReturn = String.fromCharCode(13);
  const lineFeed = String.fromCharCode(10);
  const normalizedText = String(text || '').split(carriageReturn).join('');
  const rawLines = normalizedText.split(lineFeed);
  const lines = rawLines.map((line) => line.trim()).filter(Boolean);
  return normalizeImportedRows(lines.map((line) => splitCsvLine(line)), learnedRules);
}

// Reads the first sheet from an uploaded Excel file and sends it through the same transaction normalizer as CSV.
function parseExcelArrayBuffer(buffer, learnedRules = {}) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const allTransactions = workbook.SheetNames.flatMap((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: '' });
    return normalizeImportedRows(rows, learnedRules).map((transaction) => ({
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
    return ['income', 'salary', 'net', 'amount', '◊î◊õ◊†◊°◊î', '◊©◊õ◊®', '◊†◊ò◊ï', '◊°◊õ◊ï◊ù', '◊©◊ù', '◊ú◊™◊©◊ú◊ï◊ù', '◊ñ◊õ◊ï◊™'].some((word) => joined.includes(word));
  });
  const hasHeader = headerRowIndex >= 0;
  const headers = hasHeader ? cleanedRows[headerRowIndex] : cleanedRows[0] || [];
  const dataRows = hasHeader ? cleanedRows.slice(headerRowIndex + 1) : cleanedRows;
  const sampleRows = dataRows.slice(0, 25);

  const nameIndex = hasHeader ? findHeaderIndex(headers, ['name', '◊©◊ù', '◊¢◊ï◊ë◊ì', '◊û◊ß◊ï◊®', '◊™◊ô◊ê◊ï◊®', '◊§◊ô◊®◊ï◊ò', 'income', 'salary', '◊î◊õ◊†◊°◊î', '◊©◊õ◊®', '◊û◊¢◊°◊ô◊ß'], 0) : 0;
  const amountIndex = findAmountIndex(headers, sampleRows);

  return dataRows
    .map((row) => {
      const amountCell = amountIndex >= 0 ? row[amountIndex] : [...row].reverse().find((cell) => Math.abs(toNumber(cell)) > 0);
      const amount = Math.abs(toNumber(amountCell));
      const name = row[nameIndex] || row.find((cell, index) => index !== amountIndex && String(cell || '').trim() && Math.abs(toNumber(cell)) === 0) || '◊î◊õ◊†◊°◊î ◊û◊ô◊ï◊ë◊ê◊™';
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
  const keywords = ['◊†◊ò◊ï ◊ú◊™◊©◊ú◊ï◊ù', '◊ú◊™◊©◊ú◊ï◊ù ◊ë◊ë◊†◊ß', '◊°◊î◊õ ◊ú◊™◊©◊ú◊ï◊ù', '◊°◊î◊¥◊õ ◊ú◊™◊©◊ú◊ï◊ù', '◊©◊õ◊® ◊†◊ò◊ï', '◊†◊ò◊ï'];
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
  return [{ id: makeId('income'), name: `◊†◊ò◊ï ◊û◊™◊ú◊ï◊© ${file.name}`, amount: netSalary }];
}

function getCategoryTotals(transactions) {
  return transactions.reduce((acc, transaction) => {
    const category = transaction.category || '◊ê◊ó◊®';
    acc[category] = (acc[category] || 0) + toNumber(transaction.amount);
    return acc;
  }, {});
}

function getMerchantTotals(transactions) {
  return transactions.reduce((acc, transaction) => {
    const merchant = transaction.merchant || '◊ú◊ú◊ê ◊©◊ù';
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
  const uncategorizedAmount = categoryTotals['◊ê◊ó◊®'] || 0;
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
  const savings = savingsProducts + savingGoals;
  const expenses = credit + manual + savings + selfEmployed;
  const net = income - expenses;
  const savingsRate = income ? (net / income) * 100 : 0;
  return { income, credit, manual, savings, selfEmployed, selfEmployedRaw, expenses, net, savingsRate };
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
        // Bank balance is a real bank value only. It does not include calculated monthly net / "◊ô◊™◊®◊î ◊ê◊ó◊®◊ô ◊î◊õ◊ï◊ú".
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
      buildRow('◊î◊õ◊†◊°◊ï◊™', 'income'),
      buildRow('◊î◊ï◊¶◊ê◊ï◊™', 'expenses'),
      buildRow('◊ê◊©◊®◊ê◊ô', 'credit'),
      buildRow('◊î◊ï◊¶◊ê◊ï◊™ ◊ô◊ì◊†◊ô◊ï◊™', 'manual'),
      buildRow('◊ó◊ô◊°◊õ◊ï◊ü', 'savings'),
      buildRow('◊¢◊¶◊û◊ê◊ô', 'selfEmployed'),
      buildRow('◊ô◊™◊®◊î', 'net'),
      buildRow('◊©◊ô◊¢◊ï◊® ◊ó◊ô◊°◊õ◊ï◊ü', 'savingsRate', 'percent'),
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
  if (!transactions.length) return [`◊û◊¶◊ë ${modeConfig.label}: ${modeConfig.focus}. ◊î◊¢◊ú◊ï CSV ◊ê◊ï Excel ◊õ◊ì◊ô ◊ú◊ß◊ë◊ú ◊™◊ï◊ë◊†◊ï◊™.`];

  const total = transactions.reduce((sum, item) => sum + toNumber(item.amount), 0) || 1;
  const categoryTotals = getCategoryTotals(transactions);
  const merchantTotals = getMerchantTotals(transactions);
  const sortedCategories = Object.entries(categoryTotals).sort((a, b) => toNumber(b[1]) - toNumber(a[1]));
  const sortedMerchants = Object.entries(merchantTotals).sort((a, b) => toNumber(b[1]) - toNumber(a[1]));
  const healthScore = calculateFinancialHealthScore(transactions, modeConfig);
  const insights = [
    `◊û◊¶◊ë ${modeConfig.label}: ${modeConfig.focus}. ◊®◊û◊™ ◊î◊™◊®◊ê◊ï◊™: ${modeConfig.notificationTone}.`,
    `◊¶◊ô◊ï◊ü ◊ë◊®◊ô◊ê◊ï◊™ ◊î◊ï◊¶◊ê◊ï◊™ ◊ê◊©◊®◊ê◊ô ◊ú◊§◊ô ◊û◊¶◊ë ${modeConfig.label}: ${healthScore}/100.`,
  ];

  if (typeof context.savingsRate === 'number') {
    const gap = modeConfig.savingsTarget - context.savingsRate;
    if (gap > 0) insights.push(`◊©◊ô◊¢◊ï◊® ◊î◊ó◊ô◊°◊õ◊ï◊ü ◊†◊û◊ï◊ö ◊û◊î◊ô◊¢◊ì ◊©◊ú ◊û◊¶◊ë ${modeConfig.label} ◊ë÷æ${formatPercent(gap)}.`);
    else insights.push(`◊©◊ô◊¢◊ï◊® ◊î◊ó◊ô◊°◊õ◊ï◊ü ◊¢◊ï◊û◊ì ◊ë◊ô◊¢◊ì ◊©◊ú ◊û◊¶◊ë ${modeConfig.label} ◊ï◊ê◊£ ◊í◊ë◊ï◊î ◊û◊û◊†◊ï ◊ë÷æ${formatPercent(Math.abs(gap))}.`);
  }

  const [topCategory, topCategoryAmount] = sortedCategories[0] || [];
  if (topCategory) insights.push(`◊î◊ß◊ò◊í◊ï◊®◊ô◊î ◊î◊í◊ì◊ï◊ú◊î ◊ë◊ô◊ï◊™◊® ◊ë◊ê◊©◊®◊ê◊ô ◊î◊ô◊ê ${topCategory}: ${SHEKEL.format(topCategoryAmount)}, ◊©◊î◊ù ${Math.round((topCategoryAmount / total) * 100)}% ◊û◊î◊ó◊ô◊ï◊ë◊ô◊ù.`);

  const [topMerchant, topMerchantAmount] = sortedMerchants[0] || [];
  if (topMerchant) insights.push(`◊ë◊ô◊™ ◊î◊¢◊°◊ß ◊î◊ì◊ï◊û◊ô◊†◊†◊ò◊ô ◊ë◊ô◊ï◊™◊® ◊î◊ï◊ê ${topMerchant}: ${SHEKEL.format(topMerchantAmount)}.`);

  insights.push(`◊í◊ï◊ë◊î ◊¢◊°◊ß◊™ ◊ê◊©◊®◊ê◊ô ◊û◊û◊ï◊¶◊¢◊™: ${SHEKEL.format(total / transactions.length)}.`);
  if (totalIncome > 0) insights.push(`◊ó◊ô◊ï◊ë◊ô ◊î◊ê◊©◊®◊ê◊ô ◊î◊ù ${formatPercent((total / totalIncome) * 100)} ◊û◊î◊î◊õ◊†◊°◊î ◊©◊î◊ï◊ñ◊†◊î ◊î◊ó◊ï◊ì◊©.`);

  Object.entries(CATEGORY_BUDGETS).forEach(([category, budget]) => {
    const spent = categoryTotals[category] || 0;
    const adjustedBudget = budget / modeConfig.strictness;
    const warningPoint = adjustedBudget * (modeConfig.budgetWarningAt / 100);
    if (spent > adjustedBudget) insights.push(`${category} ◊ó◊®◊í◊î ◊û◊î◊™◊ß◊¶◊ô◊ë ◊î◊û◊ï◊™◊ê◊ù ◊ú◊û◊¶◊ë ${modeConfig.label} ◊ë÷æ${SHEKEL.format(spent - adjustedBudget)}.`);
    else if (spent >= warningPoint) insights.push(`${category} ◊û◊™◊ß◊®◊ë◊™ ◊ú◊™◊ß◊¶◊ô◊ë ◊ú◊§◊ô ◊û◊¶◊ë ${modeConfig.label}: ${SHEKEL.format(spent)} ◊û◊™◊ï◊ö ${SHEKEL.format(adjustedBudget)}.`);
  });

  const uncategorized = categoryTotals['◊ê◊ó◊®'] || 0;
  if (uncategorized > 0) insights.push(`${SHEKEL.format(uncategorized)} ◊¢◊ì◊ô◊ô◊ü ◊û◊°◊ï◊ï◊í◊ô◊ù ◊õ◊¥◊ê◊ó◊®◊¥. ◊©◊ô◊†◊ï◊ô ◊ô◊ì◊†◊ô ◊©◊ú ◊ß◊ò◊í◊ï◊®◊ô◊î ◊ô◊ú◊û◊ì ◊ê◊™ ◊î◊û◊¢◊®◊õ◊™ ◊ú◊§◊¢◊û◊ô◊ù ◊î◊ë◊ê◊ï◊™.`);

  const largeTransactions = transactions
    .filter((transaction) => toNumber(transaction.amount) >= Math.max(500, total * 0.08))
    .sort((a, b) => toNumber(b.amount) - toNumber(a.amount));
  if (largeTransactions.length > 0) insights.push(`◊ñ◊ï◊î◊ï ${largeTransactions.length} ◊¢◊°◊ß◊ê◊ï◊™ ◊í◊ì◊ï◊ú◊ï◊™ ◊ô◊ó◊°◊ô◊™. ◊î◊í◊ì◊ï◊ú◊î ◊ë◊ô◊ï◊™◊®: ${largeTransactions[0].merchant} ◊ë◊°◊ö ${SHEKEL.format(largeTransactions[0].amount)}.`);

  if (recurringTransactions.length > 0) {
    const recurringTotal = recurringTransactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
    insights.push(`◊ñ◊ï◊î◊ï ${recurringTransactions.length} ◊¢◊°◊ß◊ê◊ï◊™ ◊ó◊ï◊ñ◊®◊ï◊™/◊û◊†◊ï◊ô◊ô◊ù ◊ë◊°◊ö ◊õ◊ï◊ú◊ú ◊©◊ú ${SHEKEL.format(recurringTotal)}.`);
  }

  if (modeConfig.priorityMetric === 'burnRate' && context.burnRate) insights.push(`◊ë◊û◊¶◊ë Survival ◊õ◊ì◊ê◊ô ◊ú◊î◊ï◊®◊ô◊ì Burn Rate ◊û◊™◊ó◊™ ◊ú÷æ${SHEKEL.format(context.burnRate * 0.9)} ◊ë◊ó◊ï◊ì◊© ◊î◊ë◊ê.`);
  if (modeConfig.priorityMetric === 'cashFlow' && context.cashFlow) insights.push(`◊ë◊û◊¶◊ë Growth ◊î◊ì◊í◊© ◊î◊ï◊ê ◊ú◊î◊í◊ì◊ô◊ú Cash Flow ◊§◊†◊ï◊ô ◊û◊¢◊ú ${SHEKEL.format(context.cashFlow + 1000)}.`);
  if (modeConfig.priorityMetric === 'netWorth' && context.totalAssets) insights.push(`◊ë◊û◊¶◊ë Wealth Building ◊î◊ì◊í◊© ◊î◊ï◊ê ◊ú◊î◊í◊ì◊ô◊ú ◊©◊ï◊ï◊ô ◊©◊î◊ï◊ñ◊ü ◊û◊¢◊ë◊® ◊ú÷æ${SHEKEL.format(context.totalAssets * 1.05)}.`);

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
    primaryPerson: '◊†◊ï◊¢◊î',
    householdProfileId: DEFAULT_SUPABASE_PROFILE_ID,
    secondaryPerson: '◊ê◊ï◊®◊ü',
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
    dashboardTitle: '◊û◊¢◊®◊õ◊™ ◊§◊ô◊†◊†◊°◊ô◊™ ◊û◊©◊§◊ó◊™◊ô◊™',
    emergencyFund: 0,
    lastSalaryImport: '',
    attachedDocuments: [],
    bankAccounts: [
      { id: makeId('bank'), name: '◊¢◊ï◊¥◊© ◊û◊©◊ï◊™◊£', owner: '◊û◊©◊§◊ó◊î', openingBalance: 0, closingBalance: 0, importedFile: '', transactions: [] },
      { id: makeId('bank'), name: '◊¢◊ï◊¥◊© ◊†◊ï◊¢◊î', owner: '◊†◊ï◊¢◊î', openingBalance: 0, closingBalance: 0, importedFile: '', transactions: [] },
      { id: makeId('bank'), name: '◊¢◊ï◊¥◊© ◊ê◊ï◊®◊ü', owner: '◊ê◊ï◊®◊ü', openingBalance: 0, closingBalance: 0, importedFile: '', transactions: [] },
    ],
    incomes: [
      { id: makeId('income'), name: '◊û◊©◊õ◊ï◊®◊™ ◊†◊ï◊¢◊î', amount: 0 },
      { id: makeId('income'), name: '◊î◊õ◊†◊°◊î ◊ê◊ï◊®◊ü', amount: 0 },
      { id: makeId('income'), name: '◊î◊õ◊†◊°◊î ◊†◊ï◊°◊§◊™', amount: 0 },
    ],
    manualExpenses: [
      { id: makeId('expense'), category: '◊û◊©◊õ◊†◊™◊ê / ◊©◊õ◊ô◊®◊ï◊™', amount: 0, type: '◊ß◊ë◊ï◊¢◊î' },
      { id: makeId('expense'), category: '◊ê◊®◊†◊ï◊†◊î', amount: 0, type: '◊ß◊ë◊ï◊¢◊î' },
      { id: makeId('expense'), category: '◊ó◊©◊û◊ú', amount: 0, type: '◊ß◊ë◊ï◊¢◊î' },
      { id: makeId('expense'), category: '◊û◊ô◊ù', amount: 0, type: '◊ß◊ë◊ï◊¢◊î' },
      { id: makeId('expense'), category: '◊ê◊ô◊†◊ò◊®◊†◊ò + ◊°◊ú◊ï◊ú◊®', amount: 0, type: '◊ß◊ë◊ï◊¢◊î' },
      { id: makeId('expense'), category: '◊ë◊ô◊ò◊ï◊ó◊ô◊ù', amount: 0, type: '◊ß◊ë◊ï◊¢◊î' },
    ],
    savingsProducts: [
      { id: makeId('saving'), name: '◊ß◊®◊ü ◊î◊©◊™◊ú◊û◊ï◊™ ◊†◊ï◊¢◊î', type: '◊ß◊®◊ü ◊î◊©◊™◊ú◊û◊ï◊™', owner: '◊†◊ï◊¢◊î', monthlyDeposit: 0, currentBalance: 0 },
      { id: makeId('saving'), name: '◊ß◊®◊ü ◊î◊©◊™◊ú◊û◊ï◊™ ◊ê◊ï◊®◊ü', type: '◊ß◊®◊ü ◊î◊©◊™◊ú◊û◊ï◊™', owner: '◊ê◊ï◊®◊ü', monthlyDeposit: 0, currentBalance: 0 },
      { id: makeId('saving'), name: '◊§◊†◊°◊ô◊î ◊†◊ï◊¢◊î', type: '◊§◊†◊°◊ô◊î', owner: '◊†◊ï◊¢◊î', monthlyDeposit: 0, currentBalance: 0 },
      { id: makeId('saving'), name: '◊§◊†◊°◊ô◊î ◊ê◊ï◊®◊ü', type: '◊§◊†◊°◊ô◊î', owner: '◊ê◊ï◊®◊ü', monthlyDeposit: 0, currentBalance: 0 },
    ],
    savingGoals: [
      { id: makeId('goal'), name: '◊ò◊ô◊°◊î ◊ú◊ô◊§◊ü', targetAmount: 30000, currentAmount: 0, monthlyDeposit: 0 },
      { id: makeId('goal'), name: '◊ó◊™◊ï◊†◊î', targetAmount: 100000, currentAmount: 0, monthlyDeposit: 0 },
      { id: makeId('goal'), name: '◊ß◊®◊ü ◊ó◊ô◊®◊ï◊ù', targetAmount: 60000, currentAmount: 0, monthlyDeposit: 0 },
    ],
    creditCards: [
      { id: makeId('card'), name: '◊õ◊®◊ò◊ô◊° ◊ê◊©◊®◊ê◊ô ◊†◊ï◊¢◊î', owner: '◊†◊ï◊¢◊î', importedFile: '', transactions: [], pendingTransactions: [] },
      { id: makeId('card'), name: '◊õ◊®◊ò◊ô◊° ◊ê◊©◊®◊ê◊ô ◊ê◊ï◊®◊ü', owner: '◊ê◊ï◊®◊ü', importedFile: '', transactions: [], pendingTransactions: [] },
    ],
    selfEmployed: {
      owner: '◊ê◊ï◊®◊ü',
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

// Password login is prepared for Supabase Auth, but the login screen is currently disabled until global settings are stable.
async function signInWithSupabasePassword(email, password, config = {}) {
  const supabaseUrl = config.url || SUPABASE_URL;
  const supabaseKey = config.key || SUPABASE_ANON_KEY;
  if (!supabaseUrl || !supabaseKey) throw new Error('◊ó◊°◊®◊ô◊ù Supabase URL ◊ê◊ï Publishable Key');
  if (!email || !password) throw new Error('◊¶◊®◊ô◊ö ◊ú◊î◊ñ◊ô◊ü ◊ê◊ô◊û◊ô◊ô◊ú ◊ï◊°◊ô◊°◊û◊î');

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
    throw new Error(`◊î◊õ◊†◊ô◊°◊î ◊†◊õ◊©◊ú◊î: ${response.status} ${details}`);
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

// Lightweight runtime checks catch common parser/calculation regressions while developing in the browser canvas.
function runSmokeTests() {
  console.assert(APP_BUILD_MARKER === 'finance-dashboard-build-v14', 'build marker failed');
  console.assert(getPublicEnv('THIS_ENV_SHOULD_NOT_EXIST') === '', 'safe env fallback failed');
  console.assert(toNumber('‚Ç™1,250') === 1250, 'currency parsing failed');
  console.assert(detectCategory('Wolt TLV') === '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î', 'wolt category failed');
  console.assert(detectCategory('My Shop', { shop: '◊ß◊†◊ô◊ï◊™' }) === '◊ß◊†◊ô◊ï◊™', 'learned rule failed');
  console.assert(splitCsvLine('a,b,c').length === 3, 'csv split failed');
  console.assert(parseCsvText(['date,merchant,amount', '2026-01-01,Wolt,55'].join(String.fromCharCode(10))).length === 1, 'csv parse failed');
  console.assert(parseCsvText(['date,merchant,amount', '2026-01-01,Wolt,55'].join(String.fromCharCode(13) + String.fromCharCode(10))).length === 1, 'csv CRLF parse failed');
  console.assert(splitCsvLine('"a,b",c').length === 2, 'quoted csv parsing failed');
  console.assert(getCategoryTotals([{ category: '◊ß◊†◊ô◊ï◊™', amount: 10 }, { category: '◊ß◊†◊ô◊ï◊™', amount: 20 }]).◊ß◊†◊ô◊ï◊™ === 30, 'category totals failed');
  console.assert(buildRealInsights([{ merchant: 'Wolt', category: '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î', amount: 900 }], [], 0, 'Survival').some((insight) => insight.includes('Survival')), 'real budget insight failed');
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
  console.assert(TABS.some((tab) => tab.id === 'insights' && tab.label === '◊™◊ï◊ë◊†◊ï◊™ ◊ó◊õ◊û◊ï◊™'), 'smart insights tab label failed');
  console.assert(normalizePreferences({ showTrendChart: false }).showTrendChart === false, 'preferences override failed');
  console.assert(normalizePreferences({}).householdProfileId === DEFAULT_SUPABASE_PROFILE_ID, 'household profile default failed');
  console.assert(getSafeTheme('Missing').accent === THEME_STYLES.Sage.accent, 'theme fallback failed');
  console.assert(getSafeTheme('Dark').page.includes('111111'), 'dark theme page exists');
  console.assert(noSingleWordLine('◊ê◊ó◊™ ◊©◊™◊ô◊ô◊ù ◊©◊ú◊ï◊©').includes(String.fromCharCode(160)), 'no orphan text helper failed');
  console.assert(getInitialLearnedRules() && typeof getInitialLearnedRules() === 'object', 'initial learned rules failed');
}

if (typeof window !== 'undefined') runSmokeTests();

function StatCard({ title, value, note, tone = 'neutral' }) {
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
    <div className={`min-h-[150px] rounded-[22px] border ${toneClass} p-4 shadow-sm transition hover:shadow-md sm:min-h-[180px] sm:p-5 lg:min-h-[220px]`}>
      <div className="text-center text-xs font-semibold uppercase tracking-widest text-neutral-400">{title}</div>
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
    <div className="rounded-[22px] border border-dashed border-neutral-300 bg-neutral-50 p-6 text-center sm:p-10">
      <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-2xl bg-white text-lg font-semibold text-neutral-500">Ôºã</div>
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
      className={`rounded-xl px-5 py-3 text-sm font-semibold text-white transition disabled:opacity-50 ${className}`}
      style={{ backgroundColor: accent }}
      onMouseEnter={(event) => { event.currentTarget.style.backgroundColor = accentHover; }}
      onMouseLeave={(event) => { event.currentTarget.style.backgroundColor = accent; }}
    >
      {children}
    </button>
  );
}

function GhostButton({ children, className = '', ...props }) {
  return <button {...props} className={`rounded-xl border border-neutral-200 bg-white px-4 py-2.5 text-sm font-semibold text-neutral-700 transition hover:border-neutral-400 hover:bg-neutral-50 ${className}`}>{children}</button>;
}

function Field({ className = '', ...props }) {
  return <input {...props} className={`rounded-xl border border-neutral-300 bg-white px-4 py-3 text-sm outline-none transition focus:border-neutral-900 focus:ring-2 focus:ring-neutral-100 ${className}`} />;
}

function SelectField({ children, className = '', ...props }) {
  return <select {...props} className={`rounded-xl border border-neutral-300 bg-white px-4 py-3 text-sm outline-none transition focus:border-neutral-900 focus:ring-2 focus:ring-neutral-100 ${className}`}>{children}</select>;
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
  const width = 900;
  const height = 280;
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
    return <EmptyState title="◊ê◊ô◊ü ◊¢◊ì◊ô◊ô◊ü ◊†◊™◊ï◊†◊ô◊ù ◊ú◊í◊®◊£" text="◊õ◊ì◊ô ◊ú◊®◊ê◊ï◊™ ◊û◊í◊û◊ï◊™, ◊û◊ú◊ê◊ô ◊ú◊§◊ó◊ï◊™ ◊ó◊ï◊ì◊© ◊ê◊ó◊ì ◊©◊ú ◊î◊õ◊†◊°◊ï◊™ ◊ï◊î◊ï◊¶◊ê◊ï◊™." />;
  }

  return (
    <div className="overflow-x-auto rounded-[24px] border border-neutral-200 bg-neutral-50 p-4">
      <svg viewBox={`0 0 ${width} ${height}`} className="min-w-[760px] w-full" role="img" aria-label="◊í◊®◊£ ◊û◊í◊û◊ï◊™ ◊î◊õ◊†◊°◊ï◊™ ◊î◊ï◊¶◊ê◊ï◊™ ◊ï◊ó◊ô◊°◊õ◊ï◊ü">
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
        <span className="rounded-full bg-white px-3 py-2">‚óè ◊î◊õ◊†◊°◊ï◊™</span>
        <span className="rounded-full bg-white px-3 py-2 text-amber-700">‚óè ◊î◊ï◊¶◊ê◊ï◊™</span>
        <span className="rounded-full bg-white px-3 py-2 text-blue-700">‚óè ◊ó◊ô◊°◊õ◊ï◊ü</span>
      </div>
    </div>
  );
}

function TransactionEditorTable({ rows, cardId, mode, onUpdate, onRemove }) {
  const isPending = mode === 'pending';
  return (
    <div className="mt-5 max-h-[900px] overflow-auto rounded-[24px] border border-neutral-200 bg-white">
      <div className="min-w-[720px]">
        <div className="sticky top-0 z-10 grid grid-cols-[110px_minmax(180px,1fr)_170px_120px_44px] bg-neutral-100 px-5 py-4 text-sm font-semibold text-neutral-700">
          <div>◊™◊ê◊®◊ô◊ö</div>
          <div>{isPending ? '◊ë◊ô◊™ ◊¢◊°◊ß' : '◊¢◊°◊ß◊î'}</div>
          <div>{isPending ? '◊ß◊ò◊í◊ï◊®◊ô◊î' : '◊ß◊ò◊í◊ï◊®◊ô◊î ◊ú◊ï◊û◊ì◊™'}</div>
          <div>◊°◊õ◊ï◊ù</div>
          <div />
        </div>
        {rows.map((transaction) => (
          <div key={transaction.id} className="grid grid-cols-[110px_minmax(180px,1fr)_170px_120px_44px] gap-4 border-t border-neutral-100 p-4">
            {isPending ? <Field value={transaction.date || ''} onChange={(event) => onUpdate(cardId, transaction.id, 'date', event.target.value)} /> : <div className="px-3 py-3 text-sm text-neutral-500">{transaction.date}</div>}
            {isPending ? <Field value={transaction.merchant} onChange={(event) => onUpdate(cardId, transaction.id, 'merchant', event.target.value)} /> : <Field value={transaction.merchant} readOnly className="bg-neutral-50" />}
            <SelectField value={transaction.category} onChange={(event) => isPending ? onUpdate(cardId, transaction.id, 'category', event.target.value) : onUpdate(transaction.id, event.target.value)}>
              {EXPENSE_CATEGORIES.map((category) => <option key={category}>{category}</option>)}
            </SelectField>
            {isPending ? <Field type="number" value={transaction.amount} onChange={(event) => onUpdate(cardId, transaction.id, 'amount', event.target.value)} /> : <div className="px-3 py-3 text-sm font-semibold text-neutral-900">{SHEKEL.format(transaction.amount)}</div>}
            <GhostButton onClick={() => onRemove(cardId, transaction.id)} className="px-0">√ó</GhostButton>
          </div>
        ))}
        {rows.length === 0 ? <div className="p-10 text-center text-sm text-neutral-400">◊¢◊ì◊ô◊ô◊ü ◊ú◊ê ◊î◊¢◊ú◊ô◊™ ◊§◊ô◊®◊ï◊ò ◊ê◊©◊®◊ê◊ô. ◊î◊¢◊ú◊ô CSV ◊ê◊ï Excel ◊õ◊ì◊ô ◊ú◊î◊™◊ó◊ô◊ú ◊†◊ô◊™◊ï◊ó ◊ó◊õ◊ù ◊©◊ú ◊î◊î◊ï◊¶◊ê◊ï◊™.</div> : null}
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
        <Field value={card.name} onChange={(event) => onUpdateCard(card.id, 'name', event.target.value)} placeholder="◊©◊ù ◊î◊õ◊®◊ò◊ô◊°" />
        <Field value={card.owner} onChange={(event) => onUpdateCard(card.id, 'owner', event.target.value)} placeholder="◊ë◊¢◊ú/◊™ ◊î◊õ◊®◊ò◊ô◊°" />
        <div className="rounded-xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-sm font-semibold text-neutral-800">{SHEKEL.format(cardTotal)}</div>
        <GhostButton onClick={() => onRemoveCard(card.id)} className="px-0">√ó</GhostButton>
      </div>

      <div className="mt-5 rounded-[24px] border border-dashed border-neutral-300 bg-neutral-50 p-4">
        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <div>
            <div className="text-sm font-semibold text-neutral-900">◊î◊¢◊ú◊ê◊™ CSV / Excel ◊©◊ú ◊§◊ô◊®◊ï◊ò ◊ê◊©◊®◊ê◊ô</div>
            <div className="mt-1 text-xs text-neutral-500">◊î◊¢◊ú◊ê◊î ◊†◊û◊¶◊ê◊™ ◊õ◊ê◊ü, ◊ë◊™◊ï◊ö ◊î◊õ◊®◊ò◊ô◊° ◊î◊®◊ú◊ï◊ï◊†◊ò◊ô.</div>
          </div>
          <label className="cursor-pointer rounded-xl px-4 py-3 text-sm font-semibold text-white transition" style={{ backgroundColor: safeTheme.accent }}>
            ◊î◊¢◊ú◊ê◊™ ◊ß◊ï◊ë◊•
            <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) onImportFile(card.id, file); }} />
          </label>
        </div>
        {card.importedFile ? <div className="mt-3 rounded-xl bg-white px-4 py-3 text-sm text-neutral-600">◊†◊ß◊ú◊ò ◊ß◊ï◊ë◊•: <strong>{card.importedFile}</strong></div> : null}
      </div>

      {pendingRows.length > 0 ? (
        <div className="mt-4 rounded-[24px] border border-neutral-200 bg-neutral-50 p-4">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <div className="text-sm font-semibold text-neutral-900">◊¢◊°◊ß◊ê◊ï◊™ ◊©◊ñ◊ï◊î◊ï ◊ú◊ê◊ô◊©◊ï◊®</div>
              <div className="mt-1 text-xs text-neutral-500">◊ë◊ì◊ß◊ï ◊°◊õ◊ï◊û◊ô◊ù ◊ï◊ß◊ò◊í◊ï◊®◊ô◊ï◊™ ◊ú◊§◊†◊ô ◊©◊î◊ü ◊†◊õ◊†◊°◊ï◊™ ◊ú◊î◊ï◊¶◊ê◊ï◊™.</div>
            </div>
            <PrimaryButton theme={safeTheme} onClick={() => onApprovePending(card.id)}>◊ê◊©◊® ◊ï◊î◊õ◊†◊° ◊ú◊î◊ï◊¶◊ê◊ï◊™</PrimaryButton>
          </div>
          <TransactionEditorTable rows={pendingRows} cardId={card.id} mode="pending" onUpdate={onUpdatePending} onRemove={onRemovePending} />
        </div>
      ) : null}

      <TransactionEditorTable rows={approvedRows} cardId={card.id} mode="approved" onUpdate={onUpdateCategory} onRemove={onRemoveTransaction} />
      <PrimaryButton theme={safeTheme} onClick={() => onAddTransaction(card.id)} className="mt-4">+ ◊î◊ï◊°◊§◊™ ◊¢◊°◊ß◊î</PrimaryButton>
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
  const [categoryFilter, setCategoryFilter] = useState('◊î◊õ◊ï◊ú');
  const [minAmount, setMinAmount] = useState('');
  const [maxAmount, setMaxAmount] = useState('');
  const [cloudStatus, setCloudStatus] = useState('◊ò◊ï◊¢◊ü ◊û◊î◊¢◊†◊ü‚Ä¶');
  const [hasLoadedCloud, setHasLoadedCloud] = useState(false);
  const [hasAttemptedCloudLoad, setHasAttemptedCloudLoad] = useState(false);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [comparePeriod, setComparePeriod] = useState('previous');
  const [authSession, setAuthSession] = useState(getInitialAuthSession);
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authStatus, setAuthStatus] = useState('');

  const monthData = normalizeMonthData(months[selectedMonth]);
  const preferences = normalizePreferences(globalPreferences);
  const activeTheme = getSafeTheme(preferences.themeMood);
  const isDark = preferences.themeMood === 'Dark';
  const modeConfig = getFinancialModeConfig(preferences.financialMode);
  const householdProfileId = preferences.householdProfileId || DEFAULT_SUPABASE_PROFILE_ID;
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
        setCloudStatus(data?.months ? '◊û◊°◊ï◊†◊õ◊®◊ü ◊û◊î◊¢◊†◊ü' : '◊ê◊ô◊ü ◊¢◊ì◊ô◊ô◊ü ◊†◊™◊ï◊†◊ô ◊¢◊†◊ü, ◊¢◊ï◊ë◊ì◊ô◊ù ◊û◊ß◊ï◊û◊ô◊™');
      } catch (error) {
        setCloudStatus(`◊¢◊†◊ü ◊ú◊ê ◊ñ◊û◊ô◊ü: ${error?.message || '◊©◊í◊ô◊ê◊™ Supabase'}`);
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
          setCloudStatus(supabaseConfig.url && supabaseConfig.key ? '◊†◊©◊û◊® ◊ë◊¢◊†◊ü' : '◊ú◊ê ◊î◊ï◊í◊ì◊® Supabase, ◊†◊©◊û◊® ◊û◊ß◊ï◊û◊ô◊™');
        } else {
          setCloudStatus('Local Only: ◊†◊©◊û◊® ◊®◊ß ◊ë◊ì◊§◊ì◊§◊ü');
        }
      } catch (error) {
        setCloudStatus(`◊ú◊ê ◊†◊©◊û◊® ◊ë◊¢◊†◊ü: ${error?.message || '◊©◊í◊ô◊ê◊™ Supabase'}`);
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
    setSelectedMonthData({ ...monthData, bankAccounts: [...monthData.bankAccounts, { id: makeId('bank'), name: '◊ó◊©◊ë◊ï◊ü ◊ó◊ì◊©', owner: '◊û◊©◊§◊ó◊î', openingBalance: 0, closingBalance: 0, importedFile: '', transactions: [] }] });
  }

  async function importBankFile(accountId, file) {
    try {
      const lower = file.name.toLowerCase();
      let importedTransactions = [];
      if (lower.endsWith('.csv')) importedTransactions = parseBankCsvText(await file.text());
      else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) importedTransactions = parseBankExcelArrayBuffer(await file.arrayBuffer());
      else {
        alert('◊ú◊¢◊ï◊¥◊© ◊ê◊§◊©◊® ◊ú◊î◊¢◊ú◊ï◊™ CSV ◊ê◊ï Excel ◊û◊î◊ë◊†◊ß.');
        return;
      }

      if (!importedTransactions.length) {
        setCloudStatus('◊ß◊ï◊ë◊• ◊î◊¢◊ï◊¥◊© ◊†◊ß◊ú◊ò, ◊ê◊ë◊ú ◊ú◊ê ◊ñ◊ï◊î◊ï ◊™◊†◊ï◊¢◊ï◊™. ◊ë◊ì◊ß◊ô ◊©◊ô◊© ◊¢◊û◊ï◊ì◊ï◊™ ◊™◊ê◊®◊ô◊ö, ◊§◊ô◊®◊ï◊ò, ◊ó◊ï◊ë◊î/◊ñ◊õ◊ï◊™ ◊ê◊ï ◊°◊õ◊ï◊ù.');
        alert('◊ß◊ï◊ë◊• ◊î◊¢◊ï◊¥◊© ◊†◊ß◊ú◊ò, ◊ê◊ë◊ú ◊ú◊ê ◊ñ◊ï◊î◊ï ◊™◊†◊ï◊¢◊ï◊™. ◊ê◊ù ◊ñ◊î ◊§◊ï◊®◊û◊ò ◊ê◊ó◊® ◊©◊ú ◊î◊ë◊†◊ß, ◊†◊¶◊ò◊®◊ö ◊ú◊î◊™◊ê◊ô◊ù ◊ê◊™ ◊î◊¢◊û◊ï◊ì◊ï◊™.');
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
      setCloudStatus(`◊ô◊ï◊ë◊ê◊ï ${importedTransactions.length} ◊™◊†◊ï◊¢◊ï◊™ ◊¢◊ï◊¥◊©.`);
    } catch (error) {
      setCloudStatus(`◊©◊í◊ô◊ê◊î ◊ë◊ô◊ô◊ë◊ï◊ê ◊¢◊ï◊¥◊©: ${error?.message || '◊ú◊ê ◊ô◊ì◊ï◊¢'}`);
      alert(`◊©◊í◊ô◊ê◊î ◊ë◊ô◊ô◊ë◊ï◊ê ◊¢◊ï◊¥◊©: ${error?.message || '◊ú◊ê ◊ô◊ì◊ï◊¢'}`);
    }
  }

  function addIncome() {
    setSelectedMonthData({ ...monthData, incomes: [...monthData.incomes, { id: makeId('income'), name: '◊î◊õ◊†◊°◊î ◊ó◊ì◊©◊î', amount: 0 }] });
  }

  function addManualExpense() {
    setSelectedMonthData({ ...monthData, manualExpenses: [...monthData.manualExpenses, { id: makeId('expense'), category: '◊î◊ï◊¶◊ê◊î ◊ó◊ì◊©◊î', amount: 0, type: '◊û◊©◊™◊†◊î' }] });
  }

  function addSavingsProduct() {
    setSelectedMonthData({ ...monthData, savingsProducts: [...monthData.savingsProducts, { id: makeId('saving'), name: '◊ó◊ô◊°◊õ◊ï◊ü ◊ó◊ì◊©', type: '◊ó◊ô◊°◊õ◊ï◊ü', owner: '◊û◊©◊§◊ó◊î', monthlyDeposit: 0, currentBalance: 0 }] });
  }

  function addSavingGoal() {
    setSelectedMonthData({ ...monthData, savingGoals: [...monthData.savingGoals, { id: makeId('goal'), name: '◊ô◊¢◊ì ◊ó◊ì◊©', targetAmount: 0, currentAmount: 0, monthlyDeposit: 0 }] });
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
    setCloudStatus('◊î◊™◊ú◊ï◊© ◊¶◊ï◊®◊£ ◊ú◊™◊ô◊¢◊ï◊ì. ◊°◊õ◊ï◊ù ◊î◊†◊ò◊ï ◊ú◊ê ◊û◊™◊§◊¢◊†◊ó ◊ê◊ï◊ò◊ï◊û◊ò◊ô◊™, ◊î◊ñ◊ô◊†◊ô ◊ê◊ï◊™◊ï ◊ë◊©◊ï◊®◊™ ◊î◊î◊õ◊†◊°◊î ◊î◊®◊ú◊ï◊ï◊†◊ò◊ô◊™.');
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
        alert('◊ú◊î◊õ◊†◊°◊ï◊™ ◊ê◊§◊©◊® ◊ú◊î◊¢◊ú◊ï◊™ PDF ◊™◊ú◊ï◊©, CSV ◊ê◊ï Excel.');
        return;
      }

      if (!importedIncomes.length) {
        setCloudStatus('◊î◊ß◊ï◊ë◊• ◊†◊ß◊ú◊ò, ◊ê◊ë◊ú ◊ú◊ê ◊ñ◊ï◊î◊™◊î ◊î◊õ◊†◊°◊î. ◊ê◊ù ◊ñ◊î PDF ◊°◊®◊ï◊ß ◊ê◊ï ◊™◊ú◊ï◊© ◊ì◊ó◊ï◊°, ◊¶◊®◊ô◊ö ◊ú◊î◊ñ◊ô◊ü ◊†◊ò◊ï ◊ô◊ì◊†◊ô◊™.');
        alert('◊î◊ß◊ï◊ë◊• ◊†◊ß◊ú◊ò, ◊ê◊ë◊ú ◊ú◊ê ◊ñ◊ï◊î◊™◊î ◊î◊õ◊†◊°◊î. ◊ê◊ù ◊ñ◊î PDF ◊°◊®◊ï◊ß ◊ê◊ï ◊™◊ú◊ï◊© ◊ì◊ó◊ï◊°, ◊¶◊®◊ô◊ö ◊ú◊î◊ñ◊ô◊ü ◊†◊ò◊ï ◊ô◊ì◊†◊ô◊™.');
        return;
      }

      setSelectedMonthData({
        ...monthData,
        incomes: [...monthData.incomes, ...importedIncomes],
        lastSalaryImport: file.name,
      });
      setCloudStatus(`◊ô◊ï◊ë◊ê◊ï ${importedIncomes.length} ◊©◊ï◊®◊ï◊™ ◊î◊õ◊†◊°◊î.`);
    } catch (error) {
      setCloudStatus(`◊©◊í◊ô◊ê◊î ◊ë◊ô◊ô◊ë◊ï◊ê ◊î◊õ◊†◊°◊ï◊™: ${error?.message || '◊ú◊ê ◊ô◊ì◊ï◊¢'}`);
      alert(`◊©◊í◊ô◊ê◊î ◊ë◊ô◊ô◊ë◊ï◊ê ◊î◊õ◊†◊°◊ï◊™: ${error?.message || '◊ú◊ê ◊ô◊ì◊ï◊¢'}`);
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
      setCloudStatus('◊©◊ï◊ó◊ñ◊® ◊û◊í◊ô◊ë◊ï◊ô JSON');
    } catch {
      alert('◊ß◊ï◊ë◊• ◊î◊í◊ô◊ë◊ï◊ô ◊ú◊ê ◊™◊ß◊ô◊ü. ◊†◊ê ◊ú◊î◊¢◊ú◊ï◊™ JSON ◊©◊ô◊ï◊¶◊ê ◊û◊î◊û◊¢◊®◊õ◊™.');
    }
  }

  function resetCurrentMonth() {
    const confirmed = typeof window === 'undefined' ? false : window.confirm('◊ú◊ê◊§◊° ◊ê◊™ ◊î◊ó◊ï◊ì◊© ◊î◊†◊ï◊õ◊ó◊ô? ◊î◊§◊¢◊ï◊ú◊î ◊™◊û◊ó◊ß ◊†◊™◊ï◊†◊ô◊ù ◊©◊ú ◊î◊ó◊ï◊ì◊© ◊ë◊ú◊ë◊ì.');
    if (!confirmed) return;
    setSelectedMonthData(createDefaultMonth());
  }

  function addCreditCard() {
    setSelectedMonthData({ ...monthData, creditCards: [...monthData.creditCards, { id: makeId('card'), name: '◊õ◊®◊ò◊ô◊° ◊ê◊©◊®◊ê◊ô ◊ó◊ì◊©', owner: '', importedFile: '', transactions: [], pendingTransactions: [] }] });
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
      if (lower.endsWith('.csv')) importedTransactions = parseCsvText(await file.text(), learnedRules);
      else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) importedTransactions = parseExcelArrayBuffer(await file.arrayBuffer(), learnedRules);
      else {
        alert('◊†◊ê ◊ú◊î◊¢◊ú◊ï◊™ CSV ◊ê◊ï Excel');
        return;
      }

      if (!importedTransactions.length) {
        setCloudStatus('◊î◊ß◊ï◊ë◊• ◊†◊ß◊ú◊ò, ◊ê◊ë◊ú ◊ú◊ê ◊ñ◊ï◊î◊ï ◊¢◊°◊ß◊ê◊ï◊™. ◊ë◊ì◊ß◊ô ◊©◊ô◊© ◊¢◊û◊ï◊ì◊ï◊™ ◊™◊ê◊®◊ô◊ö, ◊ë◊ô◊™ ◊¢◊°◊ß ◊ï◊°◊õ◊ï◊ù.');
        alert('◊î◊ß◊ï◊ë◊• ◊†◊ß◊ú◊ò, ◊ê◊ë◊ú ◊ú◊ê ◊ñ◊ï◊î◊ï ◊¢◊°◊ß◊ê◊ï◊™. ◊ê◊ù ◊ñ◊î ◊§◊ô◊®◊ï◊ò ◊ë◊†◊ß/◊ê◊©◊®◊ê◊ô ◊ë◊§◊ï◊®◊û◊ò ◊ê◊ó◊®, ◊†◊¶◊ò◊®◊ö ◊ú◊î◊™◊ê◊ô◊ù ◊ê◊™ ◊û◊ë◊†◊î ◊î◊¢◊û◊ï◊ì◊ï◊™.');
      } else {
        setCloudStatus(`◊ñ◊ï◊î◊ï ${importedTransactions.length} ◊¢◊°◊ß◊ê◊ï◊™ ◊ú◊ê◊ô◊©◊ï◊® ◊ë◊õ◊®◊ò◊ô◊° ◊î◊ê◊©◊®◊ê◊ô.`);
      }

      setSelectedMonthData({
        ...monthData,
        creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, importedFile: file.name, pendingTransactions: importedTransactions } : card)),
      });
    } catch (error) {
      setCloudStatus(`◊©◊í◊ô◊ê◊î ◊ë◊ô◊ô◊ë◊ï◊ê ◊î◊ß◊ï◊ë◊•: ${error?.message || '◊ú◊ê ◊ô◊ì◊ï◊¢'}`);
      alert(`◊©◊í◊ô◊ê◊î ◊ë◊ô◊ô◊ë◊ï◊ê ◊î◊ß◊ï◊ë◊•: ${error?.message || '◊ú◊ê ◊ô◊ì◊ï◊¢'}`);
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
      creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, transactions: [...(card.transactions || []), { id: makeId('tx'), date: '', merchant: '◊¢◊°◊ß◊î ◊ó◊ì◊©◊î', category: '◊ê◊ó◊®', amount: 0 }] } : card)),
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
  const totalCreditCards = allCreditTransactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalManualExpenses = monthData.manualExpenses.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalSavingsProducts = monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const totalSavingGoals = monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const totalPlannedSavings = totalSavingsProducts + totalSavingGoals;
  const includeSelfEmployed = Boolean(preferences.includeSelfEmployed);
  const selfEmployedVatDue = Math.max(0, toNumber(monthData.selfEmployed.vatCollected) - toNumber(monthData.selfEmployed.vatPaidOnExpenses));
  const rawSelfEmployedPayments = selfEmployedVatDue + toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance) + toNumber(monthData.selfEmployed.businessExpenses);
  const totalSelfEmployedPayments = includeSelfEmployed ? rawSelfEmployedPayments : 0;
  const totalExpenses = totalCreditCards + totalManualExpenses + totalPlannedSavings + totalSelfEmployedPayments;
  const monthlySavings = totalIncome - totalExpenses;
  const savingsRate = totalIncome ? (monthlySavings / totalIncome) * 100 : 0;
  const emergencyMonths = toNumber(monthData.emergencyFund) / (totalExpenses || 1);
  const totalAssets = totalBankClosing + toNumber(monthData.emergencyFund) + monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.currentBalance), 0) + monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.currentAmount), 0);
  const bankVsCalculatedCashFlow = bankBalanceChange - monthlySavings;
  const categoryTotals = useMemo(() => getCategoryTotals(allCreditTransactions), [allCreditTransactions]);
  const recurringTransactions = useMemo(() => detectRecurringTransactions(allCreditTransactions, months, selectedMonth), [allCreditTransactions, months, selectedMonth]);
  const monthlyCompare = useMemo(() => getMonthlyCompare(months, selectedMonth, comparePeriod), [months, selectedMonth, comparePeriod]);
  const trend = useMemo(() => getMonthlyTrend(months), [months]);
  const trendSixMonths = useMemo(() => trend.slice(-6), [trend]);
  const maxTrend = Math.max(1, ...trend.map((item) => item.total));
  const burnRate = trend.length ? trend.reduce((sum, item) => sum + item.total, 0) / trend.length : 0;
  const cashFlow = totalPlannedSavings;
  const topCategories = useMemo(() => Object.entries(categoryTotals).sort((a, b) => toNumber(b[1]) - toNumber(a[1])).slice(0, 6), [categoryTotals]);
  const filteredTransactions = useMemo(() => {
    const normalizedSearch = normalizeMerchantName(searchTerm);
    return allCreditTransactions.filter((transaction) => {
      const merchantMatch = normalizeMerchantName(transaction.merchant).includes(normalizedSearch);
      const categoryMatch = categoryFilter === '◊î◊õ◊ï◊ú' || transaction.category === categoryFilter;
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
    Survival: '◊î◊û◊¢◊®◊õ◊™ ◊û◊™◊û◊ß◊ì◊™ ◊õ◊®◊í◊¢ ◊ë◊¶◊û◊¶◊ï◊ù ◊î◊ï◊¶◊ê◊ï◊™ ◊ï◊©◊û◊ô◊®◊î ◊¢◊ú ◊ô◊¶◊ô◊ë◊ï◊™.',
    Stable: '◊î◊û◊¢◊®◊õ◊™ ◊û◊™◊û◊ß◊ì◊™ ◊ë◊ê◊ô◊ñ◊ï◊ü ◊§◊ô◊†◊†◊°◊ô ◊ï◊ó◊ô◊°◊õ◊ï◊ü ◊ô◊¶◊ô◊ë.',
    Growth: '◊î◊û◊¢◊®◊õ◊™ ◊û◊™◊û◊ß◊ì◊™ ◊ë◊¶◊û◊ô◊ó◊î, ◊î◊í◊ì◊ú◊™ ◊î◊õ◊†◊°◊ï◊™ ◊ï◊î◊©◊ß◊¢◊ï◊™.',
    'Wealth Building': '◊î◊û◊¢◊®◊õ◊™ ◊û◊™◊û◊ß◊ì◊™ ◊ë◊ê◊ï◊§◊ò◊ô◊û◊ô◊ñ◊¶◊ô◊î ◊ï◊ë◊†◊ô◊ô◊™ ◊î◊ï◊ü ◊ê◊®◊ï◊ö ◊ò◊ï◊ï◊ó.',
  };

  const modeInsight = {
    Survival: '◊î◊û◊ô◊ß◊ï◊ì ◊õ◊®◊í◊¢ ◊î◊ï◊ê ◊î◊ï◊®◊ì◊™ burn rate ◊ï◊¶◊û◊¶◊ï◊ù ◊î◊ï◊¶◊ê◊ï◊™ ◊ú◊ê ◊ó◊ô◊ï◊†◊ô◊ï◊™.',
    Stable: '◊î◊û◊ô◊ß◊ï◊ì ◊õ◊®◊í◊¢ ◊î◊ï◊ê ◊ê◊ô◊ñ◊ï◊ü ◊ë◊ô◊ü ◊ê◊ô◊õ◊ï◊™ ◊ó◊ô◊ô◊ù ◊ú◊ó◊ô◊°◊õ◊ï◊ü ◊ô◊¶◊ô◊ë.',
    Growth: '◊î◊û◊ô◊ß◊ï◊ì ◊õ◊®◊í◊¢ ◊î◊ï◊ê ◊î◊í◊ì◊ú◊™ ◊î◊õ◊†◊°◊ï◊™ ◊ï◊î◊©◊ß◊¢◊î ◊ë◊¶◊û◊ô◊ó◊î.',
    'Wealth Building': '◊î◊û◊ô◊ß◊ï◊ì ◊õ◊®◊í◊¢ ◊î◊ï◊ê ◊ë◊†◊ô◊ô◊™ ◊î◊ï◊ü ◊ï◊ê◊ï◊§◊ò◊ô◊û◊ô◊ñ◊¶◊ô◊î ◊§◊ô◊†◊†◊°◊ô◊™ ◊ê◊®◊ï◊õ◊™ ◊ò◊ï◊ï◊ó.',
  };

  // Notifications are derived from current totals and global notification preferences.
  const activeNotifications = [
    preferences.notifications?.budget80 && budgetUsageRate >= modeConfig.budgetWarningAt ? `◊î◊í◊¢◊™◊ù ◊ú÷æ${modeConfig.budgetWarningAt}% ◊û◊î◊™◊ß◊¶◊ô◊ë ◊ú◊§◊ô ◊û◊¶◊ë ${modeConfig.label}.` : null,
    preferences.notifications?.woltSpike && (categoryTotals['◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î'] || 0) > (CATEGORY_BUDGETS['◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î'] || 0) ? '◊û◊°◊¢◊ì◊ï◊™ ◊ï◊ë◊™◊ô ◊ß◊§◊î ◊ó◊®◊í◊ï ◊û◊î◊™◊ß◊¶◊ô◊ë ◊©◊î◊ï◊í◊ì◊®.' : null,
    preferences.notifications?.savingsDrop && savingsRate < targetSavingsRate ? '◊©◊ô◊¢◊ï◊® ◊î◊ó◊ô◊°◊õ◊ï◊ü ◊†◊û◊ï◊ö ◊û◊î◊ô◊¢◊ì ◊©◊î◊ï◊í◊ì◊®.' : null,
  ].filter(Boolean);

  const monthlyStory = totalIncome
    ? `◊ë◊û◊¶◊ë ${modeConfig.label}, ◊î◊ó◊ï◊ì◊© ◊î◊ï◊¶◊ê◊™◊ù ${SHEKEL.format(totalExpenses)} ◊©◊î◊ù ${formatPercent((totalExpenses / totalIncome) * 100)} ◊û◊î◊î◊õ◊†◊°◊î. ◊î◊¢◊ï◊¥◊© ◊î◊†◊ï◊õ◊ó◊ô ◊î◊ï◊ê ◊ô◊™◊®◊î ◊ê◊û◊ô◊™◊ô◊™ ◊û◊î◊ë◊†◊ß/◊î◊ñ◊†◊î ◊ô◊ì◊†◊ô◊™, ◊ï◊î◊ô◊™◊®◊î ◊î◊û◊ó◊ï◊©◊ë◊™ ◊ê◊ó◊®◊ô ◊î◊õ◊ï◊ú ◊†◊©◊ê◊®◊™ ◊û◊ì◊ì ◊™◊ñ◊®◊ô◊ù ◊ë◊ú◊ë◊ì. ◊ô◊¢◊ì ◊î◊ó◊ô◊°◊õ◊ï◊ü ◊ú◊û◊¶◊ë ◊î◊ñ◊î ◊î◊ï◊ê ${formatPercent(targetSavingsRate)}, ◊ï◊î◊ô◊™◊®◊î ◊î◊û◊ó◊ï◊©◊ë◊™ ◊ê◊ó◊®◊ô ◊î◊õ◊ï◊ú ◊î◊ô◊ê ${SHEKEL.format(monthlySavings)}.`
    : `◊û◊¶◊ë ${modeConfig.label} ◊§◊¢◊ô◊ú. ◊î◊™◊ó◊ô◊ú◊ï ◊ú◊î◊ñ◊ô◊ü ◊î◊õ◊†◊°◊ï◊™ ◊ï◊î◊ï◊¶◊ê◊ï◊™ ◊õ◊ì◊ô ◊ú◊ß◊ë◊ú ◊°◊ô◊§◊ï◊® ◊§◊ô◊†◊†◊°◊ô ◊ó◊ï◊ì◊©◊ô ◊û◊ï◊™◊ê◊ù.`;

  const monthlyCompareStory = monthlyCompare.hasPrevious
    ? `◊ú◊¢◊ï◊û◊™ ◊û◊û◊ï◊¶◊¢ ${monthlyCompare.period.label} (${monthlyCompare.compareMonthKeys.length} ◊ó◊ï◊ì◊©◊ô◊ù), ◊î◊î◊ï◊¶◊ê◊ï◊™ ${monthlyCompare.current.expenses >= monthlyCompare.previous.expenses ? '◊¢◊ú◊ï' : '◊ô◊®◊ì◊ï'} ◊ë÷æ${SHEKEL.format(Math.abs(monthlyCompare.current.expenses - monthlyCompare.previous.expenses))}, ◊ï◊î◊ô◊™◊®◊î ${monthlyCompare.current.net >= monthlyCompare.previous.net ? '◊î◊©◊™◊§◊®◊î' : '◊†◊ó◊ú◊©◊î'} ◊ë÷æ${SHEKEL.format(Math.abs(monthlyCompare.current.net - monthlyCompare.previous.net))}.`
    : `◊ê◊ô◊ü ◊¢◊ì◊ô◊ô◊ü ◊û◊°◊§◊ô◊ß ◊ó◊ï◊ì◊©◊ô◊ù ◊ú◊î◊©◊ï◊ï◊ê◊™ ${monthlyCompare.period.label}. ◊î◊ï◊°◊ô◊§◊ô ◊¢◊ï◊ì ◊ó◊ï◊ì◊©◊ô◊ù ◊õ◊ì◊ô ◊ú◊ß◊ë◊ú Monthly Compare ◊ê◊û◊ô◊™◊ô ◊ï◊®◊ó◊ë ◊ô◊ï◊™◊®.`;

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

  async function handleSignIn(event) {
    event.preventDefault();
    try {
      setAuthStatus('◊û◊™◊ó◊ë◊®◊™...');
      const session = await signInWithSupabasePassword(authEmail, authPassword, supabaseConfig);
      setAuthSession(session);
      setAuthPassword('');
      setAuthStatus('◊û◊ó◊ï◊ë◊®◊™');
      setCloudStatus('◊û◊ó◊ï◊ë◊®◊™ ◊ú◊ó◊©◊ë◊ï◊ü Supabase');
    } catch (error) {
      setAuthStatus(error?.message || '◊î◊õ◊†◊ô◊°◊î ◊†◊õ◊©◊ú◊î');
    }
  }

  function handleSignOut() {
    clearAuthSession();
    setAuthSession(null);
    setAuthStatus('◊î◊™◊†◊™◊ß◊™');
    setCloudStatus('◊î◊™◊†◊™◊ß◊™, ◊†◊©◊û◊® ◊û◊ß◊ï◊û◊ô◊™ ◊¢◊ì ◊õ◊†◊ô◊°◊î ◊û◊ó◊ì◊©');
  }

  // Auth UI is intentionally disabled for now until global settings and cloud sync are fully stable.
  // Login UI is intentionally parked behind false until Supabase Auth is enabled as a separate step.
  if (false && !authSession && preferences.syncMode !== 'Local Only') {
    return (
      <div dir="rtl" className={`min-h-screen p-6 text-right transition-colors duration-300 ${activeTheme.page}`} style={{ fontFamily: 'Circular, Arial, Helvetica, sans-serif' }}>
        <div className="mx-auto flex min-h-[calc(100vh-48px)] max-w-5xl items-center justify-center">
          <div className="grid w-full gap-6 lg:grid-cols-[1fr_0.9fr]">
            <section className="rounded-[32px] border border-neutral-200 bg-white p-8 shadow-sm">
              <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">SECURE ACCESS</div>
              <h1 className="mt-4 text-4xl font-semibold tracking-tight text-neutral-950">◊õ◊†◊ô◊°◊î ◊ú◊ì◊©◊ë◊ï◊®◊ì ◊î◊§◊ô◊†◊†◊°◊ô</h1>
              <p className="mt-4 text-sm leading-7 text-neutral-500">◊î◊õ◊†◊ô◊°◊î ◊û◊©◊™◊û◊©◊™ ◊ë÷æSupabase Auth ◊¢◊ù ◊ê◊ô◊û◊ô◊ô◊ú ◊ï◊°◊ô◊°◊û◊î. ◊ê◊ó◊®◊ô ◊î◊õ◊†◊ô◊°◊î, ◊©◊û◊ô◊®◊î ◊ï◊ò◊¢◊ô◊†◊î ◊ë◊¢◊†◊ü ◊ô◊©◊™◊û◊©◊ï ◊ë÷æaccess token ◊©◊ú ◊î◊û◊©◊™◊û◊©.</p>
              <form onSubmit={handleSignIn} className="mt-7 grid gap-4">
                <label className="text-sm font-semibold text-neutral-600">
                  ◊ê◊ô◊û◊ô◊ô◊ú
                  <Field type="email" value={authEmail} onChange={(event) => setAuthEmail(event.target.value)} className="mt-2 w-full" placeholder="name@example.com" />
                </label>
                <label className="text-sm font-semibold text-neutral-600">
                  ◊°◊ô◊°◊û◊î
                  <Field type="password" value={authPassword} onChange={(event) => setAuthPassword(event.target.value)} className="mt-2 w-full" placeholder="‚Ä¢‚Ä¢‚Ä¢‚Ä¢‚Ä¢‚Ä¢‚Ä¢‚Ä¢" />
                </label>
                <PrimaryButton theme={activeTheme} type="submit" className="mt-2 w-full">◊õ◊†◊ô◊°◊î</PrimaryButton>
              </form>
              {authStatus ? <div className="mt-4 rounded-2xl bg-neutral-50 p-4 text-sm leading-7 text-neutral-600">{authStatus}</div> : null}
            </section>

            <section className="rounded-[32px] border border-neutral-200 bg-neutral-50 p-8 shadow-sm">
              <h2 className="text-2xl font-semibold tracking-tight text-neutral-950">◊ú◊§◊†◊ô ◊î◊õ◊†◊ô◊°◊î</h2>
              <div className="mt-5 grid gap-4 text-sm leading-7 text-neutral-600">
                <p>◊¶◊®◊ô◊ö ◊ú◊ô◊¶◊ï◊® ◊û◊©◊™◊û◊© ◊ë÷æSupabase ◊ì◊®◊ö Authentication ‚Üí Users.</p>
                <p>◊ê◊ù ◊¢◊ï◊ì ◊ú◊ê ◊ô◊¶◊®◊™ ◊û◊©◊™◊û◊©, ◊î◊õ◊†◊ô◊°◊î ◊™◊ô◊õ◊©◊ú ◊¢◊ì ◊©◊ô◊ï◊í◊ì◊® ◊ê◊ô◊û◊ô◊ô◊ú ◊ï◊°◊ô◊°◊û◊î.</p>
                <p>◊û◊¶◊ë Local Only ◊¢◊ì◊ô◊ô◊ü ◊ê◊§◊©◊®◊ô ◊ì◊®◊ö ◊î◊î◊í◊ì◊®◊ï◊™ ◊ê◊ó◊®◊ô ◊õ◊†◊ô◊°◊î, ◊ê◊ë◊ú ◊ë◊©◊ë◊ô◊ú ◊¢◊†◊ü ◊¶◊®◊ô◊ö ◊ó◊©◊ë◊ï◊ü.</p>
              </div>
              <div className="mt-6 rounded-2xl border border-neutral-200 bg-white p-4 text-sm text-neutral-600">
                <strong>Supabase</strong>
                <div className="mt-2">{setupHealth.supabaseEnv ? '◊û◊ó◊ï◊ë◊® ◊ú◊î◊í◊ì◊®◊ï◊™ Supabase' : '◊ó◊°◊®◊ô◊ù URL ◊ê◊ï Publishable Key ◊ë◊î◊í◊ì◊®◊ï◊™ ◊©◊†◊©◊û◊®◊ï ◊ë◊ì◊§◊ì◊§◊ü'}</div>
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

      <div className="mx-auto max-w-7xl space-y-5 sm:space-y-7">
        <div className="dark-nav sticky top-0 z-40 rounded-2xl border border-neutral-200 bg-white/95 p-2 shadow-sm backdrop-blur-xl" style={isDark ? { backgroundColor: 'rgba(18, 18, 18, 0.96)', borderColor: '#333333' } : undefined}>
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
              {authSession ? <button type="button" onClick={handleSignOut} className="text-neutral-900 underline">◊î◊™◊†◊™◊ß◊ï◊™</button> : null}
            </div>
          </div>
        </div>

        <section className="hero-banner overflow-hidden rounded-[24px] border border-neutral-200 bg-white shadow-sm sm:rounded-3xl" style={isDark ? { backgroundColor: '#151515', borderColor: '#333333' } : undefined}>
          <div className={`p-5 sm:p-8 ${isDark ? 'text-white' : 'text-neutral-950'}`} style={isDark ? { backgroundColor: '#151515' } : undefined}>
            <div className="flex flex-col gap-5 lg:flex-row lg:items-center lg:justify-between">
              <div>
                <input value={monthData.dashboardTitle} onChange={(event) => updateMonthField('dashboardTitle', event.target.value)} className="w-full max-w-3xl rounded-xl border border-transparent bg-transparent px-0 py-1 text-3xl font-semibold leading-tight tracking-tight text-neutral-950 outline-none transition placeholder:text-neutral-400 sm:py-2 sm:text-4xl md:text-5xl" placeholder="◊©◊ù ◊î◊ì◊©◊ë◊ï◊®◊ì ◊î◊û◊©◊§◊ó◊™◊ô" />
                <p className="mt-3 max-w-4xl text-sm leading-7 text-neutral-500 no-orphans no-single-word-lines sm:mt-4 sm:text-base sm:leading-8">{noSingleWordLine('◊û◊û◊ú◊ê◊ô◊ù ◊î◊õ◊†◊°◊ï◊™, ◊î◊ï◊¶◊ê◊ï◊™, ◊ê◊©◊®◊ê◊ô, ◊¢◊¶◊û◊ê◊ô, ◊ß◊®◊†◊ï◊™ ◊ï◊ô◊¢◊ì◊ô◊ù. ◊î◊û◊¢◊®◊õ◊™ ◊û◊ó◊©◊ë◊™ ◊™◊ñ◊®◊ô◊ù, ◊ó◊ô◊°◊õ◊ï◊ü ◊ï◊™◊ï◊ë◊†◊ï◊™ ◊ê◊û◊ô◊™◊ô◊ï◊™.')}</p>
                <div className="mt-4 inline-flex max-w-full rounded-full px-3 py-2 text-xs font-semibold leading-6 no-orphans sm:px-4 sm:text-sm" style={{ backgroundColor: activeTheme.soft, color: activeTheme.text }}>{`${modeInsight[preferences.financialMode] || modeInsight.Stable} ◊ô◊¢◊ì ◊ó◊ô◊°◊õ◊ï◊ü: ${formatPercent(targetSavingsRate)} | ◊î◊™◊®◊ê◊î ◊ë÷æ${modeConfig.budgetWarningAt}%`}</div>
                <div className="mt-3 inline-flex max-w-full rounded-full border border-neutral-200 bg-neutral-50 px-3 py-2 text-xs font-medium leading-6 text-neutral-600 no-orphans sm:mt-5 sm:px-4 sm:text-sm">{cloudStatus}</div>
              </div>
              <div className="rounded-2xl border border-neutral-200 bg-neutral-50 p-4 sm:p-5">
                <label className="text-xs font-semibold uppercase tracking-widest text-neutral-400">◊ó◊ï◊ì◊©</label>
                <input type="month" value={selectedMonth} onChange={(event) => ensureMonth(event.target.value)} className="mt-2 w-full rounded-xl border border-neutral-300 bg-white px-4 py-3 text-base font-semibold text-neutral-900 outline-none focus:border-neutral-900 focus:ring-2 focus:ring-neutral-100 sm:text-lg" />
              </div>
            </div>
          </div>

          <div className="dark-surface grid grid-cols-1 gap-3 border-t border-neutral-100 bg-white p-4 sm:gap-4 sm:p-6 md:grid-cols-2 xl:grid-cols-4" style={isDark ? { backgroundColor: '#151515', borderColor: '#333333' } : undefined}>
            <StatCard title="◊°◊î◊¥◊õ ◊î◊õ◊†◊°◊ï◊™" value={SHEKEL.format(totalIncome)} note="◊õ◊ú ◊û◊ß◊ï◊®◊ï◊™ ◊î◊î◊õ◊†◊°◊î" tone="good" />
            <StatCard title="◊¢◊ï◊¥◊© ◊†◊ï◊õ◊ó◊ô" value={SHEKEL.format(totalBankClosing)} note={`◊©◊ô◊†◊ï◊ô ◊î◊ó◊ï◊ì◊©: ${SHEKEL.format(bankBalanceChange)}`} tone={totalBankClosing >= 0 ? 'good' : 'danger'} />
            <StatCard title="◊°◊î◊¥◊õ ◊î◊ï◊¶◊ê◊ï◊™" value={SHEKEL.format(totalExpenses)} note={effectiveBudgetTarget ? `${formatPercent(budgetUsageRate)} ◊û◊™◊ï◊ö ◊ô◊¢◊ì ${modeConfig.label}` : `${totalIncome ? formatPercent((totalExpenses / totalIncome) * 100) : '0%'} ◊û◊î◊î◊õ◊†◊°◊î`} tone={(effectiveBudgetTarget && totalExpenses > effectiveBudgetTarget) || (totalIncome && totalExpenses > totalIncome) ? 'danger' : budgetUsageRate >= modeConfig.budgetWarningAt ? 'warn' : 'neutral'} />
            <StatCard title="◊°◊î◊¥◊õ ◊ê◊©◊®◊ê◊ô" value={SHEKEL.format(totalCreditCards)} note="◊û◊õ◊®◊ò◊ô◊°◊ô ◊î◊ê◊©◊®◊ê◊ô" />
            <StatCard title="◊¢◊¶◊û◊ê◊ô" value={SHEKEL.format(totalSelfEmployedPayments)} note={includeSelfEmployed ? '◊õ◊ú◊ï◊ú ◊ë◊™◊ñ◊®◊ô◊ù ◊î◊û◊©◊§◊ó◊™◊ô' : '◊ú◊ê ◊õ◊ú◊ï◊ú ◊ë◊™◊ñ◊®◊ô◊ù'} />
            <StatCard title="◊ó◊°◊õ◊ï◊†◊ï◊™" value={SHEKEL.format(totalPlannedSavings)} note="◊ß◊®◊†◊ï◊™, ◊§◊†◊°◊ô◊î ◊ï◊ô◊¢◊ì◊ô◊ù" tone="good" />
            <StatCard title="◊ô◊™◊®◊î ◊ê◊ó◊®◊ô ◊î◊õ◊ï◊ú" value={SHEKEL.format(monthlySavings)} note={`${formatPercent(savingsRate)} ◊ó◊ô◊°◊õ◊ï◊ü / ◊ô◊¢◊ì ${formatPercent(targetSavingsRate)}`} tone={monthlySavings >= 0 && savingsRate >= targetSavingsRate ? 'good' : monthlySavings < 0 ? 'danger' : 'neutral'} />
            <StatCard title="◊©◊ï◊ï◊ô ◊©◊î◊ï◊ñ◊ü" value={SHEKEL.format(totalAssets)} note={`${emergencyMonths.toFixed(1)} ◊ó◊ï◊ì◊©◊ô ◊ó◊ô◊®◊ï◊ù`} />
          </div>
        </section>

        {activeTab === 'dashboard' ? (
          <>
            {(preferences.showMonthlyStory || preferences.showFinancialHealth || activeNotifications.length > 0) ? (
              <Section>
                <div className="grid gap-6 lg:grid-cols-[1.4fr_0.6fr]">
                  {activeNotifications.length > 0 ? (
                    <div className="grid gap-3 md:grid-cols-3 lg:col-span-2">
                      {activeNotifications.map((notification) => (
                        <div key={notification} className="rounded-[24px] border border-amber-200 bg-gradient-to-br from-amber-50 to-white px-5 py-5 shadow-sm">
                          <div className="flex items-center gap-3">
                            <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-amber-100 text-amber-700">‚ö†</div>
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
                      <h2 className="mt-4 text-4xl font-semibold leading-tight tracking-tight text-neutral-950">◊î◊°◊ô◊§◊ï◊® ◊©◊ú ◊î◊ó◊ï◊ì◊© ◊©◊ú◊õ◊ù</h2>
                      <p className="mt-5 max-w-3xl text-lg leading-9 text-neutral-600 no-orphans">{noSingleWordLine(monthlyStory)}</p>
                      <div className="mt-8 grid gap-3 md:grid-cols-3">
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">Burn Rate</div><div className="mt-2 text-xl font-semibold text-neutral-950">{SHEKEL.format(burnRate)}</div></div>
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">Cash Flow ◊ú◊ó◊ô◊°◊õ◊ï◊ü</div><div className="mt-2 text-xl font-semibold text-neutral-950">{SHEKEL.format(cashFlow)}</div></div>
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">◊©◊ô◊¢◊ï◊® ◊ó◊ô◊°◊õ◊ï◊ü</div><div className="mt-2 text-xl font-semibold text-neutral-950">{formatPercent(savingsRate)}</div></div>
                      </div>
                    </div>
                  ) : null}

                  {preferences.showFinancialHealth ? (
                    <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-6">
                      <div className="text-sm font-semibold text-neutral-500">Financial Health</div>
                      <div className="mt-4 text-6xl font-semibold text-neutral-950">{financialHealthScore}</div>
                      <div className="mt-5 h-3 overflow-hidden rounded-full bg-neutral-200"><div className="h-full rounded-full" style={{ width: `${financialHealthScore}%`, backgroundColor: activeTheme.accent }} /></div>
                      <div className="mt-3 text-sm leading-7 text-neutral-500">◊¶◊ô◊ï◊ü ◊ú◊§◊ô ◊û◊¶◊ë {modeConfig.label}: ◊ß◊©◊ô◊ó◊ï◊™ ◊™◊ß◊¶◊ô◊ë, ◊ó◊®◊ô◊í◊ï◊™, ◊§◊ô◊ñ◊ï◊® ◊î◊ï◊¶◊ê◊ï◊™ ◊ï◊¢◊°◊ß◊ê◊ï◊™ ◊í◊ì◊ï◊ú◊ï◊™.</div>
                    </div>
                  ) : null}
                </div>
              </Section>
            ) : null}

            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                <div>
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">ACCOUNTS</div>
                  <h2 className="mt-3 text-3xl font-semibold tracking-tight text-neutral-950">◊ó◊©◊ë◊ï◊†◊ï◊™ ◊ï◊¢◊ï◊¥◊©</h2>
                  <p className="mt-3 max-w-3xl text-sm leading-7 text-neutral-500 no-orphans">{noSingleWordLine('◊û◊¢◊ú◊ô◊ù ◊§◊ô◊®◊ï◊ò ◊¢◊ï◊¥◊© CSV/Excel ◊û◊î◊ë◊†◊ß, ◊ï◊î◊û◊¢◊®◊õ◊™ ◊û◊ó◊©◊ë◊™ ◊ô◊™◊®◊™ ◊§◊™◊ô◊ó◊î, ◊ô◊™◊®◊î ◊†◊ï◊õ◊ó◊ô◊™ ◊ï◊™◊†◊ï◊¢◊ï◊™. ◊î◊ô◊™◊®◊î ◊ê◊ó◊®◊ô ◊î◊õ◊ï◊ú ◊î◊ô◊ê ◊™◊ñ◊®◊ô◊ù ◊û◊ó◊ï◊©◊ë ◊ï◊ú◊ê ◊†◊õ◊†◊°◊™ ◊ê◊ï◊ò◊ï◊û◊ò◊ô◊™ ◊ú◊¢◊ï◊¥◊©.')}</p>
                </div>
                <PrimaryButton theme={activeTheme} onClick={addBankAccount}>+ ◊î◊ï◊°◊§◊™ ◊ó◊©◊ë◊ï◊ü</PrimaryButton>
              </div>

              <div className="mt-6 grid gap-4 md:grid-cols-4">
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">◊ô◊™◊®◊™ ◊§◊™◊ô◊ó◊î</div>
                  <div className="mt-3 text-2xl font-semibold text-neutral-950">{SHEKEL.format(totalBankOpening)}</div>
                </div>
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">◊ô◊™◊®◊î ◊†◊ï◊õ◊ó◊ô◊™</div>
                  <div className="mt-3 text-2xl font-semibold text-neutral-950">{SHEKEL.format(totalBankClosing)}</div>
                </div>
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">◊†◊õ◊†◊° ◊ú◊¢◊ï◊¥◊©</div>
                  <div className="mt-3 text-2xl font-semibold text-neutral-950">{SHEKEL.format(totalBankDeposits)}</div>
                </div>
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">◊ô◊¶◊ê ◊û◊î◊¢◊ï◊¥◊©</div>
                  <div className="mt-3 text-2xl font-semibold text-neutral-950">{SHEKEL.format(totalBankWithdrawals)}</div>
                </div>
              </div>

              <div className="mt-5 grid gap-3">
                {monthData.bankAccounts.map((account) => (
                  <div key={account.id} className="rounded-[22px] border border-neutral-200 bg-white p-4">
                    <div className="grid gap-3 md:grid-cols-[1.1fr_110px_150px_150px_140px_44px]">
                      <LabeledField label="◊ó◊©◊ë◊ï◊ü"><Field value={account.name} onChange={(event) => updateRow('bankAccounts', account.id, 'name', event.target.value)} placeholder="◊¢◊ï◊¥◊© ◊û◊©◊ï◊™◊£" /></LabeledField>
                      <LabeledField label="◊©◊ô◊ô◊ö ◊ú"><Field value={account.owner} onChange={(event) => updateRow('bankAccounts', account.id, 'owner', event.target.value)} placeholder="◊û◊©◊§◊ó◊î" /></LabeledField>
                      <LabeledField label="◊ô◊™◊®◊™ ◊§◊™◊ô◊ó◊î"><Field type="number" value={account.openingBalance} onChange={(event) => updateRow('bankAccounts', account.id, 'openingBalance', event.target.value)} /></LabeledField>
                      <LabeledField label="◊ô◊™◊®◊î ◊†◊ï◊õ◊ó◊ô◊™"><Field type="number" value={account.closingBalance} onChange={(event) => updateRow('bankAccounts', account.id, 'closingBalance', event.target.value)} /></LabeledField>
                      <div className="flex items-end">
                        <label className="w-full cursor-pointer rounded-xl border border-neutral-200 bg-neutral-50 px-3 py-3 text-center text-xs font-semibold text-neutral-700 transition hover:border-neutral-400 hover:bg-white">
                          ◊ô◊ô◊ë◊ï◊ê ◊¢◊ï◊¥◊©
                          <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) importBankFile(account.id, file); }} />
                        </label>
                      </div>
                      <div className="flex items-end"><GhostButton onClick={() => removeRow('bankAccounts', account.id)} className="w-full px-0">√ó</GhostButton></div>
                    </div>
                    {account.importedFile ? <div className="mt-3 rounded-xl bg-neutral-50 px-4 py-3 text-sm text-neutral-600">◊†◊ß◊ú◊ò ◊ß◊ï◊ë◊• ◊¢◊ï◊¥◊©: <strong>{account.importedFile}</strong> ¬∑ {account.transactions?.length || 0} ◊™◊†◊ï◊¢◊ï◊™</div> : null}
                    {account.transactions?.length ? (
                      <div className="mt-3 max-h-64 overflow-auto rounded-2xl border border-neutral-200">
                        <div className="grid grid-cols-[110px_1fr_130px_130px] bg-neutral-100 px-4 py-3 text-xs font-semibold text-neutral-600">
                          <div>◊™◊ê◊®◊ô◊ö</div>
                          <div>◊§◊ô◊®◊ï◊ò</div>
                          <div>◊°◊õ◊ï◊ù</div>
                          <div>◊ô◊™◊®◊î</div>
                        </div>
                        {account.transactions.slice(0, 80).map((transaction) => (
                          <div key={transaction.id} className="grid grid-cols-[110px_1fr_130px_130px] gap-3 border-t border-neutral-100 px-4 py-3 text-sm">
                            <div className="text-neutral-500">{transaction.date}</div>
                            <div>{transaction.description}</div>
                            <div className={toNumber(transaction.amount) >= 0 ? 'font-semibold text-[#66725E]' : 'font-semibold text-red-700'}>{SHEKEL.format(transaction.amount)}</div>
                            <div>{transaction.balance ? SHEKEL.format(transaction.balance) : '‚Äî'}</div>
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
                  <h2 className="mt-3 text-3xl font-semibold tracking-tight text-neutral-950">◊î◊©◊ï◊ï◊ê◊î ◊ú◊ê◊ï◊®◊ö ◊ñ◊û◊ü</h2>
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
                <div className="mt-6 grid gap-3 md:grid-cols-2 xl:grid-cols-4">
                  {monthlyCompare.rows.map((row) => {
                    const isGoodDirection = row.key === 'income' || row.key === 'net' || row.key === 'savingsRate' || row.key === 'savings';
                    const improved = isGoodDirection ? row.diff >= 0 : row.diff <= 0;
                    const value = row.type === 'percent' ? formatPercent(row.currentValue) : SHEKEL.format(row.currentValue);
                    const diff = row.type === 'percent' ? formatPercent(Math.abs(row.diff)) : SHEKEL.format(Math.abs(row.diff));
                    return (
                      <div key={row.key} className={`rounded-[24px] border p-5 ${improved ? 'border-[#D6DDCF] bg-[#F4F6F1] text-[#66725E]' : 'border-amber-200 bg-amber-50 text-amber-900'}`}>
                        <div className="text-xs font-semibold uppercase tracking-widest opacity-70">{row.label}</div>
                        <div className="mt-3 text-2xl font-semibold text-neutral-950">{value}</div>
                        <div className="mt-2 text-sm font-semibold">{improved ? '◊©◊ô◊§◊ï◊®' : '◊ì◊ï◊®◊© ◊™◊©◊ï◊û◊™ ◊ú◊ë'}: {diff}</div>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div className="mt-6">
                  <EmptyState title="◊ê◊ô◊ü ◊¢◊ì◊ô◊ô◊ü ◊û◊°◊§◊ô◊ß ◊ó◊ï◊ì◊©◊ô◊ù" text="◊¶◊®◊ô ◊ê◊ï ◊û◊ú◊ê◊ô ◊†◊™◊ï◊†◊ô◊ù ◊ë◊¢◊ï◊ì ◊ó◊ï◊ì◊©◊ô◊ù ◊õ◊ì◊ô ◊ú◊®◊ê◊ï◊™ ◊î◊©◊ï◊ï◊ê◊î ◊ê◊ï◊ò◊ï◊û◊ò◊ô◊™ ◊û◊ï◊ú 3 ◊ó◊ï◊ì◊©◊ô◊ù, 6 ◊ó◊ï◊ì◊©◊ô◊ù, ◊©◊†◊î ◊ê◊ï ◊õ◊ú ◊î◊™◊ß◊ï◊§◊î." />
                </div>
              )}
            </Section>

            {(preferences.showCategoryChart || preferences.showTrendChart) ? (
              <section className="grid gap-6 lg:grid-cols-3">
                {preferences.showCategoryChart ? (
                  <Section>
                    <div className="flex items-center justify-between gap-4"><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">◊î◊™◊§◊ú◊í◊ï◊™ ◊î◊ï◊¶◊ê◊ï◊™ ◊ú◊§◊ô ◊ß◊ò◊í◊ï◊®◊ô◊ï◊™</h2><span className="rounded-full border border-neutral-200 bg-white px-3 py-1 text-xs font-medium text-neutral-500">Heatmap</span></div>
                    <div className="mx-auto mt-6 h-56 w-56 rounded-full" style={{ background: pieChart }} />
                    <div className="mt-6 space-y-2">
                      {topCategories.length ? topCategories.map(([category, amount]) => <div key={category} className={`flex justify-between rounded-2xl border px-4 py-3 text-sm ${getBudgetHeatColor(category, amount)}`}><span>{category}</span><strong>{SHEKEL.format(amount)}</strong></div>) : <EmptyState title="◊ê◊ô◊ü ◊¢◊ì◊ô◊ô◊ü ◊ß◊ò◊í◊ï◊®◊ô◊ï◊™" text="◊î◊¢◊ú◊ô ◊§◊ô◊®◊ï◊ò ◊ê◊©◊®◊ê◊ô ◊õ◊ì◊ô ◊ú◊®◊ê◊ï◊™ ◊î◊™◊§◊ú◊í◊ï◊™ ◊¶◊ë◊¢◊ï◊†◊ô◊™ ◊ú◊§◊ô ◊ß◊ò◊í◊ï◊®◊ô◊ï◊™." />}
                    </div>
                  </Section>
                ) : null}

                {preferences.showTrendChart ? (
                  <Section className="lg:col-span-2">
                    <div className="flex items-center justify-between gap-4"><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">◊û◊í◊û◊™ 6 ◊ó◊ï◊ì◊©◊ô◊ù</h2><span className="rounded-full border border-neutral-200 bg-white px-3 py-1 text-xs font-medium text-neutral-500">Income ¬∑ Expenses ¬∑ Savings</span></div>
                    <p className="mt-2 text-sm leading-7 text-neutral-500">◊î◊õ◊†◊°◊ï◊™, ◊î◊ï◊¶◊ê◊ï◊™ ◊ï◊ó◊ô◊°◊õ◊ï◊ü ◊†◊ò◊ï ◊ú◊§◊ô ◊ó◊ï◊ì◊©◊ô◊ù. ◊ë◊ú◊ô ◊°◊§◊®◊ô◊ô◊™ ◊í◊®◊§◊ô◊ù ◊ó◊ô◊¶◊ï◊†◊ô◊™, ◊õ◊ì◊ô ◊©◊î÷æbuild ◊ô◊ô◊©◊ê◊® ◊†◊ß◊ô.</p>
                    <div className="mt-6">
                      <TrendLineChart data={trendSixMonths} theme={activeTheme} />
                    </div>
                    <div className="mt-4 rounded-2xl bg-white p-4 text-sm text-neutral-600">Burn Rate ◊û◊û◊ï◊¶◊¢: <strong>{SHEKEL.format(burnRate)}</strong> | Cash Flow ◊ú◊ó◊ô◊°◊õ◊ï◊ü: <strong>{SHEKEL.format(cashFlow)}</strong></div>
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
                <div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊°◊ô◊õ◊ï◊ù ◊õ◊®◊ò◊ô◊°◊ô ◊ê◊©◊®◊ê◊ô</h2><p className="mt-2 text-sm text-neutral-500">◊õ◊ê◊ü ◊û◊¢◊ú◊ô◊ù CSV/Excel ◊ú◊õ◊ú ◊õ◊®◊ò◊ô◊°, ◊ë◊ï◊ì◊ß◊ô◊ù ◊ß◊ò◊í◊ï◊®◊ô◊ï◊™, ◊ï◊ê◊ñ ◊û◊ê◊©◊®◊ô◊ù ◊î◊õ◊†◊°◊î ◊ú◊î◊ï◊¶◊ê◊ï◊™.</p></div>
                <PrimaryButton theme={activeTheme} onClick={addCreditCard}>+ ◊î◊ï◊°◊§◊™ ◊õ◊®◊ò◊ô◊°</PrimaryButton>
              </div>
              <div className="mt-7 grid gap-8 xl:grid-cols-2">
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
                  <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊î◊ï◊¶◊ê◊ï◊™ ◊ô◊ì◊†◊ô◊ï◊™</h2>
                  <p className="mt-2 text-sm text-neutral-500">◊î◊ï◊¶◊ê◊ï◊™ ◊©◊ú◊ê ◊†◊õ◊†◊°◊ï◊™ ◊û◊õ◊®◊ò◊ô◊°◊ô ◊î◊ê◊©◊®◊ê◊ô ◊ï◊†◊ó◊©◊ë◊ï◊™ ◊ô◊ó◊ì ◊¢◊ù ◊î◊î◊ï◊¶◊ê◊ï◊™ ◊î◊ó◊ï◊ì◊©◊ô◊ï◊™.</p>
                </div>
                <PrimaryButton theme={activeTheme} onClick={addManualExpense}>+ ◊î◊ï◊°◊§◊™ ◊î◊ï◊¶◊ê◊î</PrimaryButton>
              </div>
              <div className="mt-7 overflow-x-auto rounded-[24px] border border-neutral-200 bg-white">
                <div className="min-w-[720px]">
                  <div className="grid grid-cols-[minmax(320px,1fr)_180px_180px_60px] gap-3 bg-neutral-100 px-5 py-4 text-sm font-semibold text-neutral-700"><div>◊ß◊ò◊í◊ï◊®◊ô◊î</div><div>◊°◊ï◊í</div><div>◊°◊õ◊ï◊ù</div><div /></div>
                  {monthData.manualExpenses.map((expense) => (
                    <div key={expense.id} className="grid grid-cols-[minmax(320px,1fr)_180px_180px_60px] gap-3 border-t border-neutral-100 p-4">
                      <Field value={expense.category} onChange={(event) => updateRow('manualExpenses', expense.id, 'category', event.target.value)} className="w-full" />
                      <SelectField value={expense.type} onChange={(event) => updateRow('manualExpenses', expense.id, 'type', event.target.value)} className="w-full"><option>◊ß◊ë◊ï◊¢◊î</option><option>◊û◊©◊™◊†◊î</option><option>◊ó◊ô◊°◊õ◊ï◊ü</option><option>◊ó◊ì ◊§◊¢◊û◊ô◊™</option></SelectField>
                      <Field type="number" value={expense.amount} onChange={(event) => updateRow('manualExpenses', expense.id, 'amount', event.target.value)} className="w-full" />
                      <GhostButton onClick={() => removeRow('manualExpenses', expense.id)} className="px-0">√ó</GhostButton>
                    </div>
                  ))}
                </div>
              </div>
            </Section>

            <Section>
              <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊õ◊ú ◊¢◊°◊ß◊ê◊ï◊™ ◊î◊ê◊©◊®◊ê◊ô ◊î◊û◊°◊ï◊†◊†◊ï◊™</h2>
              <div className="mt-6 overflow-x-auto rounded-[24px] border border-neutral-200 bg-white">
                <div className="min-w-[860px]">
                  <div className="grid grid-cols-[110px_1fr_170px_120px_90px] bg-neutral-100 px-6 py-4 text-sm font-semibold text-neutral-700"><div>◊™◊ê◊®◊ô◊ö</div><div>◊ë◊ô◊™ ◊¢◊°◊ß</div><div>◊ß◊ò◊í◊ï◊®◊ô◊î ◊ú◊ï◊û◊ì◊™</div><div>◊°◊õ◊ï◊ù</div><div>◊ñ◊ô◊î◊ï◊ô</div></div>
                  {filteredTransactions.map((transaction) => {
                    const isRecurring = recurringTransactions.some((item) => item.id === transaction.id);
                    return (
                      <div key={transaction.id} className="grid grid-cols-[110px_1fr_170px_120px_90px] gap-4 border-t border-neutral-100 px-6 py-4">
                        <div>{transaction.date}</div>
                        <div>{transaction.merchant}</div>
                        <SelectField value={transaction.category} onChange={(event) => updateTransactionCategory(transaction.id, event.target.value)}>{EXPENSE_CATEGORIES.map((category) => <option key={category}>{category}</option>)}</SelectField>
                        <div className="font-semibold">{SHEKEL.format(transaction.amount)}</div>
                        <div>{isRecurring ? '◊ó◊ï◊ñ◊® ◊ß◊ë◊ï◊¢' : '‚Äî'}</div>
                      </div>
                    );
                  })}
                  {filteredTransactions.length === 0 ? <div className="p-16 text-center text-neutral-400">◊ú◊ê ◊†◊û◊¶◊ê◊ï ◊¢◊°◊ß◊ê◊ï◊™ ◊ú◊§◊ô ◊î◊ó◊ô◊§◊ï◊© ◊ï◊î◊§◊ô◊ú◊ò◊®◊ô◊ù ◊©◊ë◊ó◊®◊™◊ù.</div> : null}
                </div>
              </div>
            </Section>
          </>
        ) : null}

        {activeTab === 'savings' ? (
          <>
            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊ß◊®◊†◊ï◊™, ◊§◊†◊°◊ô◊î ◊ï◊ó◊°◊õ◊ï◊†◊ï◊™</h2><p className="mt-2 text-sm text-neutral-500">◊î◊§◊®◊©◊ï◊™ ◊ó◊ï◊ì◊©◊ô◊ï◊™ ◊ú◊ß◊®◊ü ◊î◊©◊™◊ú◊û◊ï◊™, ◊§◊†◊°◊ô◊î ◊ï◊ó◊°◊õ◊ï◊†◊ï◊™ ◊ß◊ë◊ï◊¢◊ô◊ù.</p></div><PrimaryButton theme={activeTheme} onClick={addSavingsProduct}>+ ◊î◊ï◊°◊§◊™ ◊ó◊ô◊°◊õ◊ï◊ü</PrimaryButton></div>
              <div className="mt-7 grid gap-3">
                {monthData.savingsProducts.map((product) => (
                  <div key={product.id} className="grid gap-3 rounded-[24px] border border-neutral-200 p-4 md:grid-cols-[1.4fr_150px_130px_160px_160px_44px]">
                    <LabeledField label="◊©◊ù ◊î◊ó◊ô◊°◊õ◊ï◊ü"><Field value={product.name} onChange={(event) => updateRow('savingsProducts', product.id, 'name', event.target.value)} placeholder="◊ú◊û◊©◊ú ◊§◊†◊°◊ô◊î ◊†◊ï◊¢◊î" /></LabeledField>
                    <LabeledField label="◊°◊ï◊í"><SelectField value={product.type} onChange={(event) => updateRow('savingsProducts', product.id, 'type', event.target.value)}><option>◊ß◊®◊ü ◊î◊©◊™◊ú◊û◊ï◊™</option><option>◊§◊†◊°◊ô◊î</option><option>◊ß◊ï◊§◊™ ◊í◊û◊ú</option><option>◊ó◊ô◊°◊õ◊ï◊ü</option><option>◊î◊©◊ß◊¢◊ï◊™</option></SelectField></LabeledField>
                    <LabeledField label="◊©◊ô◊ô◊ö ◊ú"><Field value={product.owner} onChange={(event) => updateRow('savingsProducts', product.id, 'owner', event.target.value)} placeholder="◊†◊ï◊¢◊î / ◊ê◊ï◊®◊ü" /></LabeledField>
                    <LabeledField label="◊î◊§◊ß◊ì◊î ◊ó◊ï◊ì◊©◊ô◊™"><Field type="number" value={product.monthlyDeposit} onChange={(event) => updateRow('savingsProducts', product.id, 'monthlyDeposit', event.target.value)} placeholder="‚Ç™ ◊ú◊ó◊ï◊ì◊©" /></LabeledField>
                    <LabeledField label="◊ô◊™◊®◊î ◊†◊ï◊õ◊ó◊ô◊™"><Field type="number" value={product.currentBalance} onChange={(event) => updateRow('savingsProducts', product.id, 'currentBalance', event.target.value)} placeholder="◊õ◊û◊î ◊†◊¶◊ë◊®" /></LabeledField>
                    <div className="flex items-end"><GhostButton onClick={() => removeRow('savingsProducts', product.id)} className="w-full px-0">√ó</GhostButton></div>
                  </div>
                ))}
              </div>
            </Section>

            <Section>
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊ô◊¢◊ì◊ô ◊ó◊ô◊°◊õ◊ï◊ü</h2><p className="mt-2 text-sm text-neutral-500">◊ò◊ô◊°◊î ◊ú◊ô◊§◊ü, ◊ó◊™◊ï◊†◊î, ◊ß◊®◊ü ◊ó◊ô◊®◊ï◊ù ◊ï◊õ◊ú ◊ô◊¢◊ì ◊ê◊ó◊®.</p></div><PrimaryButton theme={activeTheme} onClick={addSavingGoal}>+ ◊î◊ï◊°◊§◊™ ◊ô◊¢◊ì</PrimaryButton></div>
              <div className="mt-6 grid gap-5 md:grid-cols-3">
                {monthData.savingGoals.map((goal) => {
                  const progress = toNumber(goal.targetAmount) ? Math.min(100, Math.round((toNumber(goal.currentAmount) / toNumber(goal.targetAmount)) * 100)) : 0;
                  const remaining = Math.max(0, toNumber(goal.targetAmount) - toNumber(goal.currentAmount));
                  const monthlyDeposit = Math.max(1, toNumber(goal.monthlyDeposit));
                  const etaMonths = Math.ceil(remaining / monthlyDeposit);
                  const boostedEta = Math.ceil(remaining / (monthlyDeposit + 500));
                  return (
                    <div key={goal.id} className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                      <LabeledField label="◊©◊ù ◊î◊ô◊¢◊ì"><Field value={goal.name} onChange={(event) => updateRow('savingGoals', goal.id, 'name', event.target.value)} className="w-full font-semibold" placeholder="◊ú◊û◊©◊ú ◊ò◊ô◊°◊î ◊ú◊ô◊§◊ü" /></LabeledField>
                      <div className="mt-3 grid gap-3">
                        <LabeledField label="◊°◊õ◊ï◊ù ◊ô◊¢◊ì"><Field type="number" value={goal.targetAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'targetAmount', event.target.value)} placeholder="◊õ◊û◊î ◊¶◊®◊ô◊ö ◊ú◊î◊í◊ô◊¢" /></LabeledField>
                        <LabeledField label="◊†◊¶◊ë◊® ◊¢◊ì ◊¢◊õ◊©◊ô◊ï"><Field type="number" value={goal.currentAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'currentAmount', event.target.value)} placeholder="◊õ◊û◊î ◊õ◊ë◊® ◊ô◊©" /></LabeledField>
                        <LabeledField label="◊î◊§◊ß◊ì◊î ◊ó◊ï◊ì◊©◊ô◊™"><Field type="number" value={goal.monthlyDeposit} onChange={(event) => updateRow('savingGoals', goal.id, 'monthlyDeposit', event.target.value)} placeholder="◊õ◊û◊î ◊û◊ï◊°◊ô◊§◊ô◊ù ◊õ◊ú ◊ó◊ï◊ì◊©" /></LabeledField>
                      </div>
                      <div className="mt-4 flex justify-between text-sm font-semibold"><span>{progress}%</span><button onClick={() => removeRow('savingGoals', goal.id)} className="text-neutral-700">◊û◊ó◊ô◊ß◊î</button></div>
                      <div className="mt-3 rounded-2xl border border-neutral-200 bg-white p-3 text-sm leading-7 text-neutral-600"><div>ETA ◊ú◊ô◊¢◊ì: <strong>{Number.isFinite(etaMonths) ? `${etaMonths} ◊ó◊ï◊ì◊©◊ô◊ù` : '◊ú◊ê ◊û◊ï◊í◊ì◊®'}</strong></div><div className="mt-1">◊ê◊ù ◊™◊í◊ì◊ô◊ú◊ï ◊ë÷æ‚Ç™500 ◊ë◊ó◊ï◊ì◊© ◊™◊í◊ô◊¢◊ï ◊ë◊¢◊®◊ö ◊™◊ï◊ö <strong>{Number.isFinite(boostedEta) ? `${boostedEta} ◊ó◊ï◊ì◊©◊ô◊ù` : '‚Äî'}</strong>.</div></div>
                      <div className="mt-2 h-3 overflow-hidden rounded-full bg-white"><div className="h-full rounded-full" style={{ width: `${progress}%`, backgroundColor: activeTheme.accent }} /></div>
                    </div>
                  );
                })}
              </div>
            </Section>

            <Section><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊ß◊®◊ü ◊ó◊ô◊®◊ï◊ù</h2><p className="mt-2 text-sm text-neutral-500">◊û◊ú◊ê◊ï ◊°◊õ◊ï◊ù ◊ó◊ô◊°◊õ◊ï◊ü ◊†◊ñ◊ô◊ú ◊†◊ï◊õ◊ó◊ô.</p><Field type="number" value={monthData.emergencyFund} onChange={(event) => updateMonthField('emergencyFund', toNumber(event.target.value))} className="mt-6 w-full text-xl font-semibold" /></Section>
          </>
        ) : null}

        {activeTab === 'income' ? (
          <section className="grid gap-6 lg:grid-cols-2">
            <Section>
              <div className="flex items-center justify-between gap-4"><div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊î◊õ◊†◊°◊ï◊™</h2><p className="mt-2 text-sm text-neutral-500">◊ê◊§◊©◊® ◊ú◊ô◊ô◊ë◊ê ◊î◊õ◊†◊°◊ï◊™ ◊û÷æPDF ◊™◊ú◊ï◊©, CSV ◊ê◊ï Excel. ◊ê◊ù ◊î÷æPDF ◊ò◊ß◊°◊ò◊ï◊ê◊ú◊ô, ◊î◊û◊¢◊®◊õ◊™ ◊™◊†◊°◊î ◊ú◊ñ◊î◊ï◊™ ◊†◊ò◊ï ◊ú◊™◊©◊ú◊ï◊ù ◊ê◊ï◊ò◊ï◊û◊ò◊ô◊™.</p></div><div className="flex flex-wrap gap-3"><label className="cursor-pointer rounded-xl bg-neutral-950 px-4 py-3 text-sm font-semibold text-white transition hover:bg-neutral-800">◊ô◊ô◊ë◊ï◊ê ◊î◊õ◊†◊°◊ï◊™ PDF/CSV/Excel<input type="file" accept="application/pdf,.csv,.xlsx,.xls" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) importIncomeFile(file); }} /></label><label className="cursor-pointer rounded-xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-sm font-semibold text-neutral-700 transition hover:border-neutral-400 hover:bg-white">◊¶◊ô◊®◊ï◊£ ◊™◊ú◊ï◊© PDF<input type="file" accept="application/pdf" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) attachSalarySlipFile(file); }} /></label><PrimaryButton theme={activeTheme} onClick={addIncome}>+ ◊î◊ï◊°◊§◊î</PrimaryButton></div></div>
              {(monthData.attachedDocuments || []).length ? <div className="mt-4 space-y-2 rounded-2xl bg-neutral-50 p-4 text-sm text-neutral-600">{monthData.attachedDocuments.map((document) => <div key={document.id} className="flex items-center justify-between gap-3"><span>◊™◊ú◊ï◊© ◊û◊¶◊ï◊®◊£: <strong>{document.name}</strong></span><button type="button" onClick={() => removeAttachedDocument(document.id)} className="font-semibold text-neutral-700">◊î◊°◊®◊î</button></div>)}</div> : null}
              <div className="mt-6 space-y-3">{monthData.incomes.map((income) => <InputRow key={income.id}><Field value={income.name} onChange={(event) => updateRow('incomes', income.id, 'name', event.target.value)} /><Field type="number" value={income.amount} onChange={(event) => updateRow('incomes', income.id, 'amount', event.target.value)} /><GhostButton onClick={() => removeRow('incomes', income.id)} className="px-0">√ó</GhostButton></InputRow>)}</div>
            </Section>

            <Section>
              <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊î◊õ◊†◊°◊ï◊™ ◊ê◊ï◊®◊ü / ◊¢◊¶◊û◊ê◊ô</h2><p className="mt-2 text-sm text-neutral-500">◊ê◊ù ◊î◊¢◊°◊ß ◊†◊§◊®◊ì, ◊û◊ñ◊ô◊†◊ô◊ù ◊õ◊ê◊ü ◊®◊ß ◊ê◊™ ◊î◊°◊õ◊ï◊ù ◊©◊ê◊ï◊®◊ü ◊û◊¢◊ë◊ô◊® ◊ú◊ó◊©◊ë◊ï◊ü ◊î◊û◊©◊ï◊™◊£ ◊õ◊î◊õ◊†◊°◊î. ◊ê◊§◊©◊® ◊ú◊î◊ì◊ú◊ô◊ß ◊û◊¶◊ë ◊¢◊¶◊û◊ê◊ô ◊®◊ß ◊ê◊ù ◊®◊ï◊¶◊ô◊ù ◊ú◊õ◊ú◊ï◊ú ◊û◊¢◊¥◊û, ◊û◊° ◊ï◊ë◊ô◊ò◊ï◊ó ◊ú◊ê◊ï◊û◊ô ◊ë◊™◊ñ◊®◊ô◊ù ◊î◊ë◊ô◊™.</p>
              <div className="mt-5 rounded-[24px] border border-neutral-200 bg-neutral-50 p-4">
                <label className="flex items-center justify-between gap-4 text-sm font-semibold text-neutral-700">
                  <span>◊ú◊õ◊ú◊ï◊ú ◊¢◊¶◊û◊ê◊ô ◊ë◊™◊ñ◊®◊ô◊ù ◊î◊û◊©◊§◊ó◊™◊ô</span>
                  <input type="checkbox" checked={includeSelfEmployed} onChange={(event) => updatePreference('includeSelfEmployed', event.target.checked)} className="h-5 w-5" style={{ accentColor: activeTheme.accent }} />
                </label>
                <p className="mt-3 text-xs leading-6 text-neutral-500">◊õ◊ë◊ï◊ô: ◊î◊¢◊°◊ß ◊†◊©◊ê◊® ◊û◊ó◊ï◊• ◊ú◊ì◊©◊ë◊ï◊®◊ì, ◊ï◊®◊ß ◊î◊¢◊ë◊®◊î/◊û◊©◊õ◊ï◊®◊™ ◊ú◊ó◊©◊ë◊ï◊ü ◊î◊û◊©◊ï◊™◊£ ◊†◊°◊§◊®◊™ ◊õ◊î◊õ◊†◊°◊î. ◊ì◊ï◊ú◊ß: ◊™◊©◊ú◊ï◊û◊ô ◊¢◊¶◊û◊ê◊ô ◊†◊°◊§◊®◊ô◊ù ◊õ◊î◊ï◊¶◊ê◊ï◊™ ◊ë◊ô◊™.</p>
              </div>
              <div className="mt-6 grid gap-4 md:grid-cols-2">{[['owner', '◊ë◊¢◊ú ◊î◊¢◊°◊ß', 'text'], ['salaryTransferToHousehold', '◊î◊¢◊ë◊®◊î / ◊û◊©◊õ◊ï◊®◊™ ◊ú◊ó◊©◊ë◊ï◊ü ◊î◊û◊©◊ï◊™◊£', 'number'], ['grossRevenue', '◊î◊õ◊†◊°◊î ◊¢◊°◊ß◊ô◊™ ◊ë◊®◊ï◊ò◊ï', 'number'], ['vatCollected', '◊û◊¢◊¥◊û ◊©◊†◊í◊ë◊î ◊û◊ú◊ß◊ï◊ó◊ï◊™', 'number'], ['vatPaidOnExpenses', '◊û◊¢◊¥◊û ◊¢◊ú ◊î◊ï◊¶◊ê◊ï◊™ ◊û◊ï◊õ◊®◊ï◊™', 'number'], ['incomeTaxAdvance', '◊û◊ß◊ì◊û◊™ ◊û◊° ◊î◊õ◊†◊°◊î', 'number'], ['nationalInsurance', '◊ë◊ô◊ò◊ï◊ó ◊ú◊ê◊ï◊û◊ô', 'number'], ['businessExpenses', '◊î◊ï◊¶◊ê◊ï◊™ ◊¢◊°◊ß◊ô◊ï◊™ ◊©◊©◊ï◊ú◊û◊ï ◊î◊ó◊ï◊ì◊©', 'number']].map(([field, label, type]) => <label key={field} className="text-sm font-semibold text-neutral-600">{label}<Field type={type} value={monthData.selfEmployed[field]} onChange={(event) => updateSelfEmployedField(field, event.target.value)} className="mt-2 w-full" /></label>)}</div>
              <div className="mt-6 grid gap-4 md:grid-cols-3"><StatCard title="◊û◊¢◊¥◊û ◊¶◊§◊ï◊ô" value={SHEKEL.format(selfEmployedVatDue)} note="◊†◊í◊ë◊î ◊§◊ó◊ï◊™ ◊û◊ï◊õ◊®" /><StatCard title="◊û◊° + ◊ë◊ô◊ò◊ï◊ó" value={SHEKEL.format(toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance))} note="◊™◊©◊ú◊ï◊û◊ô ◊ó◊ï◊ë◊î" /><StatCard title="◊°◊î◊¥◊õ ◊¢◊¶◊û◊ê◊ô" value={SHEKEL.format(totalSelfEmployedPayments)} note={includeSelfEmployed ? '◊õ◊ú◊ï◊ú ◊ë◊ë◊ô◊™' : '◊û◊ó◊ï◊• ◊ú◊ë◊ô◊™'} /></div>
            </Section>
          </section>
        ) : null}

        {activeTab === 'insights' ? (
          <section className="grid gap-6 lg:grid-cols-2">
            {preferences.showSmartInsightCards ? <Section><div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between"><div><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">◊™◊ï◊ë◊†◊ï◊™ ◊ó◊õ◊û◊ï◊™</h2><p className="mt-2 text-sm text-neutral-500">◊™◊ï◊ë◊†◊ï◊™ ◊û◊ó◊ï◊©◊ë◊ï◊™ ◊ô◊©◊ô◊®◊ï◊™ ◊û◊î◊†◊™◊ï◊†◊ô◊ù: ◊ó◊®◊ô◊í◊ï◊™, ◊™◊ß◊¶◊ô◊ë◊ô◊ù, ◊ë◊™◊ô ◊¢◊°◊ß ◊û◊ï◊ë◊ô◊ú◊ô◊ù, ◊ó◊ô◊ï◊ë◊ô◊ù ◊ó◊ï◊ñ◊®◊ô◊ù ◊ï◊ì◊§◊ï◊°◊ô◊ù ◊ó◊ï◊ì◊©◊ô◊ô◊ù.</p></div><div className="rounded-full px-4 py-2 text-sm font-semibold" style={{ backgroundColor: activeTheme.soft, color: activeTheme.text }}>◊û◊™◊¢◊ì◊õ◊ü ◊ê◊ï◊ò◊ï◊û◊ò◊ô◊™</div></div><div className="mt-5 grid gap-4">{realInsights.map((insight, index) => <div key={insight} className="flex items-start gap-4 rounded-[24px] border border-neutral-200 bg-white p-5 shadow-sm"><div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-neutral-100 text-lg font-semibold text-neutral-500">{index % 3 === 0 ? '‚óî' : index % 3 === 1 ? '‚ñ≤' : '‚ú¶'}</div><div className="flex-1 text-sm leading-7 text-neutral-700 no-orphans">{noSingleWordLine(insight)}</div></div>)}</div></Section> : null}
            {preferences.showRecurringDetection ? <Section><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">◊ñ◊ô◊î◊ï◊ô ◊ó◊ô◊ï◊ë◊ô◊ù ◊ß◊ë◊ï◊¢◊ô◊ù</h2><p className="mt-2 text-sm text-neutral-500">◊ñ◊ô◊î◊ï◊ô ◊û◊†◊ï◊ô◊ô◊ù, ◊ë◊ô◊ò◊ï◊ó◊ô◊ù, ◊°◊ú◊ï◊ú◊® ◊ï◊©◊õ◊ô◊®◊ï◊™ ◊ú◊§◊ô ◊û◊ô◊ú◊ï◊™ ◊û◊§◊™◊ó ◊ï◊ó◊ñ◊®◊î ◊ë◊ô◊ü ◊ó◊ï◊ì◊©◊ô◊ù.</p><div className="mt-5 space-y-3">{recurringTransactions.length ? recurringTransactions.map((item) => <div key={item.id} className="flex justify-between rounded-2xl bg-neutral-50 p-4 text-sm"><span>{item.merchant}</span><strong>{SHEKEL.format(item.amount)}</strong></div>) : <EmptyState title="◊ê◊ô◊ü ◊¢◊ì◊ô◊ô◊ü ◊ó◊ô◊ï◊ë◊ô◊ù ◊ß◊ë◊ï◊¢◊ô◊ù" text="◊î◊¢◊ú◊ô ◊§◊ô◊®◊ï◊ò◊ô◊ù ◊©◊ú ◊õ◊û◊î ◊ó◊ï◊ì◊©◊ô◊ù ◊õ◊ì◊ô ◊©◊†◊ï◊õ◊ú ◊ú◊ñ◊î◊ï◊™ ◊û◊†◊ï◊ô◊ô◊ù ◊ï◊™◊©◊ú◊ï◊û◊ô◊ù ◊ó◊ï◊ñ◊®◊ô◊ù ◊ë◊¶◊ï◊®◊î ◊ó◊õ◊û◊î." />}</div></Section> : null}
          </section>
        ) : null}

        {activeTab === 'settings' ? (
          <>
            <Section>
              <div className="flex flex-col gap-2">
                <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊î◊™◊ê◊û◊î ◊ê◊ô◊©◊ô◊™</h2>
                <p className="text-sm leading-7 text-neutral-500">
                  ◊õ◊ê◊ü ◊û◊í◊ì◊ô◊®◊ô◊ù ◊ê◊ô◊ö ◊î◊ò◊ï◊§◊° ◊ï◊î◊ì◊©◊ë◊ï◊®◊ì ◊ô◊™◊†◊î◊í◊ï: ◊©◊û◊ï◊™, ◊ô◊¢◊ì◊ô◊ù ◊ï◊û◊î ◊ô◊ï◊¶◊í ◊ë◊û◊°◊ö ◊î◊®◊ê◊©◊ô.
                </p>
              </div>

              <div className="mt-6 grid gap-6 lg:grid-cols-2">
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">◊§◊®◊ò◊ô ◊î◊ë◊ô◊™</h3>
                  <div className="mt-4 grid gap-3">
                    <label className="text-sm font-semibold text-neutral-600">
                      ◊©◊ù ◊î◊ì◊©◊ë◊ï◊®◊ì
                      <Field
                        value={monthData.dashboardTitle}
                        onChange={(event) => updateMonthField('dashboardTitle', event.target.value)}
                        className="mt-2 w-full"
                      />
                    </label>
                    <label className="text-sm font-semibold text-neutral-600">
                      ◊û◊ñ◊î◊î ◊ë◊ô◊™ / Household ID
                      <Field
                        value={householdProfileId}
                        onChange={(event) => updatePreference('householdProfileId', event.target.value || DEFAULT_SUPABASE_PROFILE_ID)}
                        className="mt-2 w-full"
                      />
                    </label>
                    <div className="grid gap-3 md:grid-cols-2">
                      <label className="text-sm font-semibold text-neutral-600">
                        ◊û◊©◊™◊û◊©/◊™ ◊®◊ê◊©◊ï◊ü/◊î
                        <Field
                          value={preferences.primaryPerson}
                          onChange={(event) => updatePreference('primaryPerson', event.target.value)}
                          className="mt-2 w-full"
                        />
                      </label>
                      <label className="text-sm font-semibold text-neutral-600">
                        ◊û◊©◊™◊û◊©/◊™ ◊©◊†◊ô/◊î
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
                            setCloudStatus('◊©◊ï◊û◊®...');
                            await saveFinanceStateToSupabase(months, learnedRules, preferences, householdProfileId, supabaseConfig);
                            setCloudStatus('◊†◊©◊û◊® ◊ë◊¢◊†◊ü');
                          } catch (error) {
                            setCloudStatus(`◊©◊í◊ô◊ê◊î ◊ë◊©◊û◊ô◊®◊î: ${error?.message || '◊ú◊ê ◊ô◊ì◊ï◊¢'}`);
                          }
                        }}
                      >
                        ◊©◊û◊ï◊® ◊î◊í◊ì◊®◊ï◊™ ◊ë◊¢◊†◊ü
                      </PrimaryButton>
                      <GhostButton onClick={() => window.location.reload()}>
                        ◊®◊¢◊†◊ü ◊ó◊ô◊ë◊ï◊®
                      </GhostButton>
                    </div>
                  </div>
                </div>

                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">◊ô◊¢◊ì◊ô◊ù ◊ó◊ï◊ì◊©◊ô◊ô◊ù</h3>
                  <div className="mt-4 grid gap-3 md:grid-cols-2">
                    <label className="text-sm font-semibold text-neutral-600">
                      ◊ô◊¢◊ì ◊î◊ï◊¶◊ê◊ï◊™ ◊ó◊ï◊ì◊©◊ô
                      <Field
                        type="number"
                        value={preferences.monthlyBudgetTarget}
                        onChange={(event) => updatePreference('monthlyBudgetTarget', event.target.value)}
                        className="mt-2 w-full"
                      />
                    </label>
                    <label className="text-sm font-semibold text-neutral-600">
                      ◊ô◊¢◊ì ◊©◊ô◊¢◊ï◊® ◊ó◊ô◊°◊õ◊ï◊ü ◊ë◊ê◊ó◊ï◊ñ◊ô◊ù
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
                <div className="mt-4 grid gap-3 md:grid-cols-2 lg:grid-cols-3">
                  {[
                    ['showMonthlyStory', 'Monthly Story'],
                    ['showFinancialHealth', 'Financial Health'],
                    ['showCategoryChart', '◊í◊®◊£ ◊ß◊ò◊í◊ï◊®◊ô◊ï◊™'],
                    ['showTrendChart', '◊í◊®◊£ ◊û◊í◊û◊î'],
                    ['showSmartInsightCards', '◊õ◊®◊ò◊ô◊°◊ô ◊™◊ï◊ë◊†◊ï◊™'],
                    ['showRecurringDetection', '◊ñ◊ô◊î◊ï◊ô ◊ó◊ô◊ï◊ë◊ô◊ù ◊ß◊ë◊ï◊¢◊ô◊ù'],
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

              <div className="mt-6 grid gap-6 lg:grid-cols-2">
                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">Theme Mood</h3>
                  <div className="mt-4 grid grid-cols-2 gap-3">
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
                            <span>◊ô◊¢◊ì ◊ó◊ô◊°◊õ◊ï◊ü {formatPercent(config.savingsTarget)}</span>
                            <span>◊î◊™◊®◊ê◊î ◊ë÷æ{config.budgetWarningAt}%</span>
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
                    <div className="mt-1 text-neutral-500">{setupHealth.localStorage ? '◊§◊¢◊ô◊ú' : '◊ú◊ê ◊ñ◊û◊ô◊ü'}</div>
                  </div>
                  <div className="rounded-2xl bg-neutral-50 p-4 text-sm">
                    <strong>Supabase ENV</strong>
                    <div className="mt-1 text-neutral-500">{setupHealth.supabaseEnv ? '◊û◊ï◊í◊ì◊®' : '◊ú◊ê ◊û◊ï◊í◊ì◊®'}</div>
                  </div>
                  <div className="rounded-2xl bg-neutral-50 p-4 text-sm">
                    <strong>Excel Parser</strong>
                    <div className="mt-1 text-neutral-500">{setupHealth.xlsxParser ? '◊§◊¢◊ô◊ú' : '◊ó◊°◊® xlsx'}</div>
                  </div>
                  <div className="rounded-2xl bg-neutral-50 p-4 text-sm">
                    <strong>Household</strong>
                    <div className="mt-1 text-neutral-500">{setupHealth.householdProfileId}</div>
                  </div>
                </div>
                <p className="mt-4 text-xs leading-6 text-neutral-500">
                  ◊ñ◊î ◊ú◊ê ◊ì◊û◊ï: ◊õ◊ú ◊°◊ò◊ò◊ï◊° ◊õ◊ê◊ü ◊û◊©◊ß◊£ ◊ó◊ô◊ë◊ï◊® ◊ê◊û◊ô◊™◊ô ◊ë◊ß◊ï◊ì. ◊ê◊ù Supabase ENV ◊ú◊ê ◊û◊ï◊í◊ì◊®, ◊î◊û◊¢◊®◊õ◊™ ◊¢◊ï◊ë◊ì◊™ ◊ë◊û◊¶◊ë LocalStorage + JSON Backup.
                </p>
              </div>

              <div className="mt-6 grid gap-6 lg:grid-cols-2">
                <div className="rounded-[24px] border border-neutral-200 bg-white p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">Smart Notifications</h3>
                  <div className="mt-4 space-y-3">
                    {[
                      ['budget80', '◊î◊™◊®◊ê◊î ◊ú◊§◊ô ◊û◊¶◊ë ◊§◊ô◊†◊†◊°◊ô'],
                      ['woltSpike', '◊î◊™◊®◊ê◊î ◊õ◊©◊ï◊ï◊ú◊ò ◊¢◊ï◊ú◊î ◊û◊©◊û◊¢◊ï◊™◊ô◊™'],
                      ['savingsDrop', '◊î◊™◊®◊ê◊î ◊õ◊©◊©◊ô◊¢◊ï◊® ◊î◊ó◊ô◊°◊õ◊ï◊ü ◊ô◊ï◊®◊ì'],
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
                    <PrimaryButton theme={activeTheme} onClick={exportBackup}>◊ô◊ô◊¶◊ï◊ê ◊í◊ô◊ë◊ï◊ô JSON</PrimaryButton>
                    <label className="cursor-pointer rounded-xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-center text-sm font-semibold text-neutral-700">
                      ◊ô◊ô◊ë◊ï◊ê ◊í◊ô◊ë◊ï◊ô
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
                    <GhostButton onClick={resetCurrentMonth}>◊ê◊ô◊§◊ï◊° ◊ó◊ï◊ì◊© ◊†◊ï◊õ◊ó◊ô</GhostButton>
                  </div>
                  <p className="mt-4 text-xs leading-6 text-neutral-500">
                    Cloud Sync ◊¢◊ï◊ë◊ì ◊®◊ß ◊ê◊ù Supabase ◊û◊ï◊í◊ì◊®. Local Only ◊©◊ï◊û◊® ◊ë◊ì◊§◊ì◊§◊ü. ◊ô◊ô◊¶◊ï◊ê/◊ô◊ô◊ë◊ï◊ê JSON ◊¢◊ï◊ë◊ì ◊™◊û◊ô◊ì.
                  </p>
                </div>
              </div>
            </Section>

            <Section>
              <h2 className="text-3xl font-semibold tracking-tight text-neutral-950">◊ó◊ô◊§◊ï◊© ◊ï◊§◊ô◊ú◊ò◊®◊ô◊ù</h2>
              <div className="mt-5 grid gap-3">
                <Field value={searchTerm} onChange={(event) => setSearchTerm(event.target.value)} placeholder="◊ó◊ô◊§◊ï◊© ◊ë◊ô◊™ ◊¢◊°◊ß, ◊ú◊û◊©◊ú ◊ï◊ï◊ú◊ò" />
                <SelectField value={categoryFilter} onChange={(event) => setCategoryFilter(event.target.value)}>
                  <option>◊î◊õ◊ï◊ú</option>
                  {EXPENSE_CATEGORIES.map((category) => <option key={category}>{category}</option>)}
                </SelectField>
                <div className="grid gap-3 md:grid-cols-2">
                  <Field value={minAmount} onChange={(event) => setMinAmount(event.target.value)} type="number" placeholder="◊°◊õ◊ï◊ù ◊û◊ô◊†◊ô◊û◊ï◊ù" />
                  <Field value={maxAmount} onChange={(event) => setMaxAmount(event.target.value)} type="number" placeholder="◊°◊õ◊ï◊ù ◊û◊ß◊°◊ô◊û◊ï◊ù" />
                </div>
              </div>
            </Section>
          </>
        ) : null}
      </div>
    </div>
  );
}

