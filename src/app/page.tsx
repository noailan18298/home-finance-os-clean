// @ts-nocheck

'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

const STORAGE_KEY = 'family-finance-os-stable-v14';
const DEFAULT_SUPABASE_PROFILE_ID = 'default-household';
const APP_BUILD_MARKER = 'finance-dashboard-build-v14';

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
  'משכנתא / שכירות',
  'ארנונה',
  'חשמל',
  'מים',
  'אינטרנט + סלולר',
  'ביטוחים',
  'סופר / מזון',
  'מסעדות / וולט',
  'תחבורה / דלק',
  'בריאות',
  'קניות',
  'בידור / מנויים',
  'חיסכון / השקעות',
  'מס הכנסה',
  'מע״מ',
  'ביטוח לאומי',
  'הוצאות עסקיות',
  'אחר',
];

const CATEGORY_BUDGETS = {
  'סופר / מזון': 4000,
  'מסעדות / וולט': 800,
  'תחבורה / דלק': 1800,
  קניות: 1200,
  בריאות: 800,
  'בידור / מנויים': 600,
  אחר: 1000,
};

const MERCHANT_CATEGORY_MAP = {
  wolt: 'מסעדות / וולט',
  tenbis: 'מסעדות / וולט',
  shufersal: 'סופר / מזון',
  שופרסל: 'סופר / מזון',
  רמי: 'סופר / מזון',
  victory: 'סופר / מזון',
  ויקטורי: 'סופר / מזון',
  yellow: 'תחבורה / דלק',
  דור: 'תחבורה / דלק',
  פז: 'תחבורה / דלק',
  fox: 'קניות',
  zara: 'קניות',
  superpharm: 'בריאות',
  סופרפארם: 'בריאות',
  כללית: 'בריאות',
  netflix: 'בידור / מנויים',
  spotify: 'בידור / מנויים',
  icloud: 'בידור / מנויים',
  google: 'בידור / מנויים',
  apple: 'בידור / מנויים',
};

const RECURRING_KEYWORDS = [
  'netflix', 'spotify', 'icloud', 'google', 'apple', 'cellcom', 'partner', 'pelephone', 'hot', 'yes',
  'ביטוח', 'הראל', 'מגדל', 'כלל', 'סלקום', 'פרטנר', 'פלאפון', 'שכירות',
];

const SHEKEL = new Intl.NumberFormat('he-IL', { style: 'currency', currency: 'ILS', maximumFractionDigits: 0 });

function getPublicEnv(key) {
  if (typeof process !== 'undefined' && process.env && process.env[key]) return process.env[key];
  return '';
}

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

const SUPABASE_URL = getPublicEnv('NEXT_PUBLIC_SUPABASE_URL');
const SUPABASE_ANON_KEY = getPublicEnv('NEXT_PUBLIC_SUPABASE_ANON_KEY');

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
  return new Date().toISOString().slice(0, 7);
}

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const number = Number(String(value || '').replace(/[₪,]/g, '').replace(/\s/g, '').trim());
  return Number.isFinite(number) ? number : 0;
}

function formatPercent(value) {
  return `${Math.round(value || 0)}%`;
}

function monthLabel(monthKey) {
  if (!monthKey) return '';
  const [year, month] = monthKey.split('-');
  return `${month}/${year}`;
}

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

function detectCategory(merchant = '', learnedRules = {}) {
  const normalized = normalizeMerchantName(merchant);
  for (const [key, category] of Object.entries(learnedRules || {})) {
    if (normalized.includes(normalizeMerchantName(key))) return category;
  }
  for (const [key, category] of Object.entries(MERCHANT_CATEGORY_MAP)) {
    if (normalized.includes(normalizeMerchantName(key))) return category;
  }
  return 'אחר';
}

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

function normalizeImportedRows(rows, learnedRules = {}) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const firstRow = rows[0].join(' ').toLowerCase();
  const hasHeader = ['date', 'תאריך', 'amount', 'סכום', 'merchant', 'בית עסק', 'שם בית עסק'].some((word) => firstRow.includes(word));
  const dataRows = hasHeader ? rows.slice(1) : rows;

  return dataRows
    .map((row) => {
      const cells = row.map((cell) => String(cell || '').trim());
      const date = cells[0] || '';
      const merchant = cells[1] || cells[0] || 'עסקה';
      const rawAmount = cells[2] || cells[cells.length - 1] || '0';
      const amount = Math.abs(toNumber(rawAmount));
      return { id: makeId('tx'), date, merchant, amount, category: detectCategory(merchant, learnedRules) };
    })
    .filter((transaction) => transaction.amount > 0);
}

function parseCsvText(text, learnedRules = {}) {
  const carriageReturn = String.fromCharCode(13);
  const lineFeed = String.fromCharCode(10);
  const normalizedText = String(text || '').split(carriageReturn).join('');
  const rawLines = normalizedText.split(lineFeed);
  const lines = rawLines.map((line) => line.trim()).filter(Boolean);
  return normalizeImportedRows(lines.map((line) => splitCsvLine(line)), learnedRules);
}

function parseExcelArrayBuffer(buffer, learnedRules = {}) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) return [];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1, raw: false, defval: '' });
  return normalizeImportedRows(rows, learnedRules);
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

function getCompareMonthKeys(months, selectedMonth, periodId = 'previous') {
  const period = COMPARE_PERIODS.find((item) => item.id === periodId) || COMPARE_PERIODS[0];
  const sortedMonths = Object.keys(months || {}).filter((month) => month < selectedMonth).sort((a, b) => b.localeCompare(a));
  if (period.id === 'all') return sortedMonths;
  return sortedMonths.slice(0, period.months);
}

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
      return { month, total: totals.credit + totals.manual };
    });
}

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

async function loadFinanceStateFromSupabase(profileId = DEFAULT_SUPABASE_PROFILE_ID) {
  if (!SUPABASE_URL || !SUPABASE_ANON_KEY) return null;
  const safeProfileId = profileId || DEFAULT_SUPABASE_PROFILE_ID;
  const url = `${SUPABASE_URL}/rest/v1/finance_app_state?profile_id=eq.${encodeURIComponent(safeProfileId)}&select=months,learned_rules`;
  const response = await fetch(url, { headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${SUPABASE_ANON_KEY}` } });
  if (!response.ok) throw new Error('Supabase load failed');
  const rows = await response.json();
  return Array.isArray(rows) ? rows[0] || null : null;
}

async function saveFinanceStateToSupabase(months, learnedRules, profileId = DEFAULT_SUPABASE_PROFILE_ID) {
  if (!SUPABASE_URL || !SUPABASE_ANON_KEY) return;
  const response = await fetch(`${SUPABASE_URL}/rest/v1/finance_app_state`, {
    method: 'POST',
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
      'Content-Type': 'application/json',
      Prefer: 'resolution=merge-duplicates',
    },
    body: JSON.stringify({ profile_id: profileId || DEFAULT_SUPABASE_PROFILE_ID, months, learned_rules: learnedRules, updated_at: new Date().toISOString() }),
  });
  if (!response.ok) throw new Error('Supabase save failed');
}

function createDefaultMonth() {
  return {
    dashboardTitle: 'מערכת פיננסית משפחתית',
    emergencyFund: 0,
    lastSalaryImport: '',
    attachedDocuments: [],
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
    preferences: {
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
      notifications: { budget80: true, woltSpike: true, savingsDrop: true },
    },
  };
}

function normalizeMonthData(data) {
  const base = createDefaultMonth();
  const safe = data || {};
  const safePreferences = safe.preferences || {};
  return {
    ...base,
    ...safe,
    incomes: Array.isArray(safe.incomes) ? safe.incomes : base.incomes,
    manualExpenses: Array.isArray(safe.manualExpenses) ? safe.manualExpenses : base.manualExpenses,
    savingsProducts: Array.isArray(safe.savingsProducts) ? safe.savingsProducts : base.savingsProducts,
    savingGoals: Array.isArray(safe.savingGoals) ? safe.savingGoals : base.savingGoals,
    creditCards: (Array.isArray(safe.creditCards) ? safe.creditCards : base.creditCards).map((card) => ({ transactions: [], pendingTransactions: [], importedFile: '', ...card })),
    attachedDocuments: Array.isArray(safe.attachedDocuments) ? safe.attachedDocuments : base.attachedDocuments,
    selfEmployed: { ...base.selfEmployed, ...(safe.selfEmployed || {}) },
    preferences: {
      ...base.preferences,
      ...safePreferences,
      notifications: { ...base.preferences.notifications, ...(safePreferences.notifications || {}) },
    },
  };
}

function getInitialMonths() {
  const currentMonth = getCurrentMonthKey();
  const saved = safeJsonParse(getStorageItem(STORAGE_KEY), null);
  if (saved && typeof saved === 'object' && !Array.isArray(saved)) {
    return saved[currentMonth] ? saved : { ...saved, [currentMonth]: createDefaultMonth() };
  }
  return { [currentMonth]: createDefaultMonth() };
}

function getInitialLearnedRules() {
  const saved = safeJsonParse(getStorageItem(`${STORAGE_KEY}-rules`), {});
  return saved && typeof saved === 'object' && !Array.isArray(saved) ? saved : {};
}

function runSmokeTests() {
  console.assert(APP_BUILD_MARKER === 'finance-dashboard-build-v14', 'build marker failed');
  console.assert(getPublicEnv('THIS_ENV_SHOULD_NOT_EXIST') === '', 'safe env fallback failed');
  console.assert(toNumber('₪1,250') === 1250, 'currency parsing failed');
  console.assert(detectCategory('Wolt TLV') === 'מסעדות / וולט', 'wolt category failed');
  console.assert(detectCategory('My Shop', { shop: 'קניות' }) === 'קניות', 'learned rule failed');
  console.assert(splitCsvLine('a,b,c').length === 3, 'csv split failed');
  console.assert(parseCsvText(['date,merchant,amount', '2026-01-01,Wolt,55'].join(String.fromCharCode(10))).length === 1, 'csv parse failed');
  console.assert(parseCsvText(['date,merchant,amount', '2026-01-01,Wolt,55'].join(String.fromCharCode(13) + String.fromCharCode(10))).length === 1, 'csv CRLF parse failed');
  console.assert(splitCsvLine('"a,b",c').length === 2, 'quoted csv parsing failed');
  console.assert(getCategoryTotals([{ category: 'קניות', amount: 10 }, { category: 'קניות', amount: 20 }]).קניות === 30, 'category totals failed');
  console.assert(buildRealInsights([{ merchant: 'Wolt', category: 'מסעדות / וולט', amount: 900 }], [], 0, 'Survival').some((insight) => insight.includes('Survival')), 'real budget insight failed');
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
  console.assert(TABS.some((tab) => tab.id === 'insights' && tab.label === 'תובנות חכמות'), 'smart insights tab label failed');
  console.assert(normalizeMonthData({ preferences: { showTrendChart: false } }).preferences.showTrendChart === false, 'preferences override failed');
  console.assert(normalizeMonthData({}).preferences.householdProfileId === DEFAULT_SUPABASE_PROFILE_ID, 'household profile default failed');
  console.assert(getSafeTheme('Missing').accent === THEME_STYLES.Sage.accent, 'theme fallback failed');
  console.assert(getSafeTheme('Dark').page.includes('111111'), 'dark theme page exists');
  console.assert(noSingleWordLine('אחת שתיים שלוש').includes(String.fromCharCode(160)), 'no orphan text helper failed');
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
    <div className={`min-h-[220px] rounded-[24px] border ${toneClass} p-5 shadow-sm transition hover:shadow-md`}>
      <div className="text-center text-xs font-semibold uppercase tracking-widest text-neutral-400">{title}</div>
      <div className="mt-6 text-center text-3xl font-semibold tracking-tight text-neutral-950">{value}</div>
      <div className={`mt-6 px-2 text-center text-sm font-medium leading-7 ${noteClass} no-single-word-lines`}>{noSingleWordLine(note)}</div>
    </div>
  );
}

function Section({ children, className = '' }) {
  return <section className={`rounded-[28px] border border-neutral-200 bg-white p-8 shadow-sm ${className}`}>{children}</section>;
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

function InputRow({ children }) {
  return <div className="grid gap-3 rounded-[24px] border border-neutral-200 p-4 md:grid-cols-[1fr_160px_44px]">{children}</div>;
}

function TransactionEditorTable({ rows, cardId, mode, onUpdate, onRemove }) {
  const isPending = mode === 'pending';
  return (
    <div className="mt-5 max-h-[900px] overflow-auto rounded-[24px] border border-neutral-200 bg-white">
      <div className="min-w-[720px]">
        <div className="sticky top-0 z-10 grid grid-cols-[110px_minmax(180px,1fr)_170px_120px_44px] bg-neutral-100 px-5 py-4 text-sm font-semibold text-neutral-700">
          <div>תאריך</div>
          <div>{isPending ? 'בית עסק' : 'עסקה'}</div>
          <div>{isPending ? 'קטגוריה' : 'קטגוריה לומדת'}</div>
          <div>סכום</div>
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
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonthKey());
  const [months, setMonths] = useState(getInitialMonths);
  const [learnedRules, setLearnedRules] = useState(getInitialLearnedRules);
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('הכול');
  const [minAmount, setMinAmount] = useState('');
  const [maxAmount, setMaxAmount] = useState('');
  const [cloudStatus, setCloudStatus] = useState('טוען מהענן…');
  const [hasLoadedCloud, setHasLoadedCloud] = useState(false);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [comparePeriod, setComparePeriod] = useState('previous');

  const monthData = normalizeMonthData(months[selectedMonth]);
  const activeTheme = getSafeTheme(monthData.preferences.themeMood);
  const isDark = monthData.preferences.themeMood === 'Dark';
  const modeConfig = getFinancialModeConfig(monthData.preferences.financialMode);
  const householdProfileId = monthData.preferences.householdProfileId || DEFAULT_SUPABASE_PROFILE_ID;
  const setupHealth = {
    localStorage: typeof window !== 'undefined',
    supabaseEnv: Boolean(SUPABASE_URL && SUPABASE_ANON_KEY),
    householdProfileId,
    xlsxParser: Boolean(XLSX && XLSX.read),
  };

  useEffect(() => {
    async function loadCloudState() {
      try {
        const data = await loadFinanceStateFromSupabase(householdProfileId);
        if (data?.months) {
          setMonths((current) => {
            const nextMonths = data.months && typeof data.months === 'object' && !Array.isArray(data.months) ? data.months : current;
            return nextMonths[selectedMonth] ? nextMonths : { ...nextMonths, [selectedMonth]: current[selectedMonth] || createDefaultMonth() };
          });
        }
        if (data?.learned_rules) setLearnedRules(data.learned_rules);
        setCloudStatus(data?.months ? 'מסונכרן מהענן' : 'אין עדיין נתוני ענן, עובדים מקומית');
      } catch {
        setCloudStatus('ענן לא זמין כרגע, עובדים מקומית');
      } finally {
        setHasLoadedCloud(true);
      }
    }
    loadCloudState();
  }, [householdProfileId, selectedMonth]);

  useEffect(() => {
    setStorageItem(STORAGE_KEY, JSON.stringify(months));
    setStorageItem(`${STORAGE_KEY}-rules`, JSON.stringify(learnedRules));
  }, [months, learnedRules]);

  useEffect(() => {
    if (!hasLoadedCloud) return;
    const saveTimeout = setTimeout(async () => {
      try {
        if (monthData.preferences.syncMode === 'Cloud Sync' || monthData.preferences.syncMode === 'Auto Backup') {
          await saveFinanceStateToSupabase(months, learnedRules, householdProfileId);
          setCloudStatus(SUPABASE_URL && SUPABASE_ANON_KEY ? 'נשמר בענן' : 'לא הוגדר Supabase, נשמר מקומית');
        } else {
          setCloudStatus('Local Only: נשמר רק בדפדפן');
        }
      } catch {
        setCloudStatus('לא נשמר בענן, נשמר מקומית');
      }
    }, 900);
    return () => clearTimeout(saveTimeout);
  }, [months, learnedRules, hasLoadedCloud, monthData.preferences.syncMode, householdProfileId]);

  function setSelectedMonthData(nextData) {
    setMonths((current) => ({ ...current, [selectedMonth]: nextData }));
  }

  function ensureMonth(monthKey) {
    setSelectedMonth(monthKey);
    if (!months[monthKey]) setMonths((current) => ({ ...current, [monthKey]: createDefaultMonth() }));
  }

  function updateMonthField(field, value) {
    setSelectedMonthData({ ...monthData, [field]: value });
  }

  function updateRow(section, id, field, value) {
    const numericFields = ['amount', 'monthlyDeposit', 'currentBalance', 'targetAmount', 'currentAmount'];
    setSelectedMonthData({
      ...monthData,
      [section]: monthData[section].map((row) => (row.id === id ? { ...row, [field]: numericFields.includes(field) ? toNumber(value) : value } : row)),
    });
  }

  function removeRow(section, id) {
    setSelectedMonthData({ ...monthData, [section]: monthData[section].filter((row) => row.id !== id) });
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

  function updatePreference(field, value) {
    const numericFields = ['monthlyBudgetTarget', 'savingsRateTarget'];
    setSelectedMonthData({ ...monthData, preferences: { ...monthData.preferences, [field]: numericFields.includes(field) ? toNumber(value) : value } });
  }

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
  }

  function removeAttachedDocument(documentId) {
    setSelectedMonthData({ ...monthData, attachedDocuments: (monthData.attachedDocuments || []).filter((document) => document.id !== documentId) });
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

  async function importCreditFile(cardId, file) {
    const lower = file.name.toLowerCase();
    let importedTransactions = [];
    if (lower.endsWith('.csv')) importedTransactions = parseCsvText(await file.text(), learnedRules);
    else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) importedTransactions = parseExcelArrayBuffer(await file.arrayBuffer(), learnedRules);
    else {
      alert('נא להעלות CSV או Excel');
      return;
    }
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, importedFile: file.name, pendingTransactions: importedTransactions } : card)),
    });
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
      creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, transactions: [...(card.transactions || []), { id: makeId('tx'), date: '', merchant: 'עסקה חדשה', category: 'אחר', amount: 0 }] } : card)),
    });
  }

  function removeTransaction(cardId, transactionId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, transactions: (card.transactions || []).filter((transaction) => transaction.id !== transactionId) } : card)),
    });
  }

  const totalIncome = monthData.incomes.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const allCreditTransactions = useMemo(() => monthData.creditCards.flatMap((card) => card.transactions || []), [monthData.creditCards]);
  const totalCreditCards = allCreditTransactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalManualExpenses = monthData.manualExpenses.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalSavingsProducts = monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const totalSavingGoals = monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const totalPlannedSavings = totalSavingsProducts + totalSavingGoals;
  const includeSelfEmployed = Boolean(monthData.preferences.includeSelfEmployed);
  const selfEmployedVatDue = Math.max(0, toNumber(monthData.selfEmployed.vatCollected) - toNumber(monthData.selfEmployed.vatPaidOnExpenses));
  const rawSelfEmployedPayments = selfEmployedVatDue + toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance) + toNumber(monthData.selfEmployed.businessExpenses);
  const totalSelfEmployedPayments = includeSelfEmployed ? rawSelfEmployedPayments : 0;
  const totalExpenses = totalCreditCards + totalManualExpenses + totalPlannedSavings + totalSelfEmployedPayments;
  const monthlySavings = totalIncome - totalExpenses;
  const savingsRate = totalIncome ? (monthlySavings / totalIncome) * 100 : 0;
  const emergencyMonths = toNumber(monthData.emergencyFund) / (totalExpenses || 1);
  const totalAssets = toNumber(monthData.emergencyFund) + monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.currentBalance), 0) + monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.currentAmount), 0);
  const categoryTotals = useMemo(() => getCategoryTotals(allCreditTransactions), [allCreditTransactions]);
  const recurringTransactions = useMemo(() => detectRecurringTransactions(allCreditTransactions, months, selectedMonth), [allCreditTransactions, months, selectedMonth]);
  const monthlyCompare = useMemo(() => getMonthlyCompare(months, selectedMonth, comparePeriod), [months, selectedMonth, comparePeriod]);
  const trend = useMemo(() => getMonthlyTrend(months), [months]);
  const maxTrend = Math.max(1, ...trend.map((item) => item.total));
  const burnRate = trend.length ? trend.reduce((sum, item) => sum + item.total, 0) / trend.length : 0;
  const cashFlow = totalPlannedSavings;
  const topCategories = useMemo(() => Object.entries(categoryTotals).sort((a, b) => toNumber(b[1]) - toNumber(a[1])).slice(0, 6), [categoryTotals]);
  const filteredTransactions = useMemo(() => {
    const normalizedSearch = normalizeMerchantName(searchTerm);
    return allCreditTransactions.filter((transaction) => {
      const merchantMatch = normalizeMerchantName(transaction.merchant).includes(normalizedSearch);
      const categoryMatch = categoryFilter === 'הכול' || transaction.category === categoryFilter;
      const amount = toNumber(transaction.amount);
      return merchantMatch && categoryMatch && (minAmount === '' || amount >= toNumber(minAmount)) && (maxAmount === '' || amount <= toNumber(maxAmount));
    });
  }, [allCreditTransactions, searchTerm, categoryFilter, minAmount, maxAmount]);

  const financialHealthScore = calculateFinancialHealthScore(allCreditTransactions, modeConfig) || 0;
  const monthlyBudgetTarget = toNumber(monthData.preferences.monthlyBudgetTarget);
  const effectiveBudgetTarget = monthlyBudgetTarget ? monthlyBudgetTarget / modeConfig.strictness : 0;
  const budgetUsageRate = effectiveBudgetTarget ? (totalExpenses / effectiveBudgetTarget) * 100 : 0;
  const targetSavingsRate = toNumber(monthData.preferences.savingsRateTarget) || modeConfig.savingsTarget;
  const realInsights = useMemo(
    () => buildRealInsights(allCreditTransactions, recurringTransactions, totalIncome, monthData.preferences.financialMode, { savingsRate, burnRate, cashFlow, totalAssets }),
    [allCreditTransactions, recurringTransactions, totalIncome, monthData.preferences.financialMode, savingsRate, burnRate, cashFlow, totalAssets]
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

  const activeNotifications = [
    monthData.preferences.notifications?.budget80 && budgetUsageRate >= modeConfig.budgetWarningAt ? `הגעתם ל־${modeConfig.budgetWarningAt}% מהתקציב לפי מצב ${modeConfig.label}.` : null,
    monthData.preferences.notifications?.woltSpike && (categoryTotals['מסעדות / וולט'] || 0) > (CATEGORY_BUDGETS['מסעדות / וולט'] || 0) ? 'וולט חרג מהתקציב שהוגדר.' : null,
    monthData.preferences.notifications?.savingsDrop && savingsRate < targetSavingsRate ? 'שיעור החיסכון נמוך מהיעד שהוגדר.' : null,
  ].filter(Boolean);

  const monthlyStory = totalIncome
    ? `במצב ${modeConfig.label}, החודש הוצאתם ${SHEKEL.format(totalExpenses)} שהם ${formatPercent((totalExpenses / totalIncome) * 100)} מההכנסה. יעד החיסכון למצב הזה הוא ${formatPercent(targetSavingsRate)}, והיתרה אחרי הכול היא ${SHEKEL.format(monthlySavings)}.`
    : `מצב ${modeConfig.label} פעיל. התחילו להזין הכנסות והוצאות כדי לקבל סיפור פיננסי חודשי מותאם.`;

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

  const pieChart = topCategories.length
    ? `conic-gradient(${topCategories.map(([, amount], index) => {
        const start = topCategories.slice(0, index).reduce((sum, [, value]) => sum + value, 0) / (totalCreditCards || 1);
        const end = topCategories.slice(0, index + 1).reduce((sum, [, value]) => sum + value, 0) / (totalCreditCards || 1);
        const colors = [activeTheme.accent, '#111111', '#737373', '#a3a3a3', '#d4d4d4', '#f5f5f5'];
        return `${colors[index % colors.length]} ${start * 100}% ${end * 100}%`;
      }).join(', ')})`
    : 'conic-gradient(#dddddd 0% 100%)';

  return (
    <div dir="rtl" className={`min-h-screen p-6 text-right transition-colors duration-300 ${activeTheme.page} ${isDark ? 'theme-dark' : ''}`} style={{ fontFamily: 'Circular, Arial, Helvetica, sans-serif' }}>
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

      <div className="mx-auto max-w-7xl space-y-7">
        <div className="dark-nav sticky top-0 z-40 rounded-2xl border border-neutral-200 bg-white/95 p-2 shadow-sm backdrop-blur-xl" style={isDark ? { backgroundColor: 'rgba(18, 18, 18, 0.96)', borderColor: '#333333' } : undefined}>
          <div className="flex flex-wrap gap-1">
            {TABS.map((tab) => (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveTab(tab.id)}
                className={`rounded-xl px-4 py-2.5 text-sm font-semibold transition ${activeTab === tab.id ? 'text-white shadow-sm' : 'text-neutral-600 hover:bg-neutral-100 hover:text-neutral-950'}`}
                style={activeTab === tab.id ? { backgroundColor: activeTheme.accent } : undefined}
              >
                {tab.label}
              </button>
            ))}
          </div>
        </div>

        <section className="hero-banner overflow-hidden rounded-3xl border border-neutral-200 bg-white shadow-sm" style={isDark ? { backgroundColor: '#151515', borderColor: '#333333' } : undefined}>
          <div className={`p-8 ${isDark ? 'text-white' : 'text-neutral-950'}`} style={isDark ? { backgroundColor: '#151515' } : undefined}>
            <div className="flex flex-col gap-6 lg:flex-row lg:items-center lg:justify-between">
              <div>
                <input value={monthData.dashboardTitle} onChange={(event) => updateMonthField('dashboardTitle', event.target.value)} className="w-full max-w-3xl rounded-xl border border-transparent bg-transparent px-0 py-2 text-4xl font-semibold tracking-tight text-neutral-950 outline-none transition placeholder:text-neutral-400 md:text-5xl" placeholder="שם הדשבורד המשפחתי" />
                <p className="mt-4 max-w-4xl text-base leading-8 text-neutral-500 no-orphans no-single-word-lines">{noSingleWordLine('ממלאים הכנסות, הוצאות, אשראי, עצמאי, קרנות ויעדים. המערכת מחשבת תזרים, חיסכון ותובנות אמיתיות.')}</p>
                <div className="nowrap-chip mt-4 inline-flex max-w-full rounded-full px-4 py-2 text-sm font-semibold no-orphans" style={{ backgroundColor: activeTheme.soft, color: activeTheme.text }}>{noSingleWordLine(`${modeInsight[monthData.preferences.financialMode] || modeInsight.Stable} יעד חיסכון: ${formatPercent(targetSavingsRate)} | התראה ב־${modeConfig.budgetWarningAt}%`)}</div>
                <div className="nowrap-chip mt-5 inline-flex max-w-full rounded-full border border-neutral-200 bg-neutral-50 px-4 py-2 text-sm font-medium text-neutral-600 no-orphans">{noSingleWordLine(cloudStatus)}</div>
              </div>
              <div className="rounded-2xl border border-neutral-200 bg-neutral-50 p-5">
                <label className="text-xs font-semibold uppercase tracking-widest text-neutral-400">חודש</label>
                <input type="month" value={selectedMonth} onChange={(event) => ensureMonth(event.target.value)} className="mt-2 w-full rounded-xl border border-neutral-300 bg-white px-4 py-3 text-lg font-semibold text-neutral-900 outline-none focus:border-neutral-900 focus:ring-2 focus:ring-neutral-100" />
              </div>
            </div>
          </div>

          <div className="dark-surface grid grid-cols-1 gap-4 border-t border-neutral-100 bg-white p-6 md:grid-cols-2 xl:grid-cols-7" style={isDark ? { backgroundColor: '#151515', borderColor: '#333333' } : undefined}>
            <StatCard title="סה״כ הכנסות" value={SHEKEL.format(totalIncome)} note="כל מקורות ההכנסה" tone="good" />
            <StatCard title="סה״כ הוצאות" value={SHEKEL.format(totalExpenses)} note={effectiveBudgetTarget ? `${formatPercent(budgetUsageRate)} מתוך יעד ${modeConfig.label}` : `${totalIncome ? formatPercent((totalExpenses / totalIncome) * 100) : '0%'} מההכנסה`} tone={(effectiveBudgetTarget && totalExpenses > effectiveBudgetTarget) || (totalIncome && totalExpenses > totalIncome) ? 'danger' : budgetUsageRate >= modeConfig.budgetWarningAt ? 'warn' : 'neutral'} />
            <StatCard title="סה״כ אשראי" value={SHEKEL.format(totalCreditCards)} note="מכרטיסי האשראי" />
            <StatCard title="עצמאי" value={SHEKEL.format(totalSelfEmployedPayments)} note={includeSelfEmployed ? 'כלול בתזרים המשפחתי' : 'לא כלול בתזרים'} />
            <StatCard title="חסכונות" value={SHEKEL.format(totalPlannedSavings)} note="קרנות, פנסיה ויעדים" tone="good" />
            <StatCard title="יתרה אחרי הכול" value={SHEKEL.format(monthlySavings)} note={`${formatPercent(savingsRate)} חיסכון / יעד ${formatPercent(targetSavingsRate)}`} tone={monthlySavings >= 0 && savingsRate >= targetSavingsRate ? 'good' : monthlySavings < 0 ? 'danger' : 'neutral'} />
            <StatCard title="שווי שהוזן" value={SHEKEL.format(totalAssets)} note={`${emergencyMonths.toFixed(1)} חודשי חירום`} />
          </div>
        </section>

        {activeTab === 'dashboard' ? (
          <>
            {(monthData.preferences.showMonthlyStory || monthData.preferences.showFinancialHealth || activeNotifications.length > 0) ? (
              <Section>
                <div className="grid gap-6 lg:grid-cols-[1.4fr_0.6fr]">
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

                  {monthData.preferences.showMonthlyStory ? (
                    <div>
                      <div className="text-xs font-semibold uppercase tracking-widest text-neutral-400">MONTHLY STORY</div>
                      <h2 className="mt-4 text-4xl font-semibold leading-tight tracking-tight text-neutral-950">הסיפור של החודש שלכם</h2>
                      <p className="mt-5 max-w-3xl text-lg leading-9 text-neutral-600 no-orphans">{noSingleWordLine(monthlyStory)}</p>
                      <div className="mt-8 grid gap-3 md:grid-cols-3">
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">Burn Rate</div><div className="mt-2 text-xl font-semibold text-neutral-950">{SHEKEL.format(burnRate)}</div></div>
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">Cash Flow לחיסכון</div><div className="mt-2 text-xl font-semibold text-neutral-950">{SHEKEL.format(cashFlow)}</div></div>
                        <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 px-5 py-4"><div className="text-xs font-semibold text-neutral-400">שיעור חיסכון</div><div className="mt-2 text-xl font-semibold text-neutral-950">{formatPercent(savingsRate)}</div></div>
                      </div>
                    </div>
                  ) : null}

                  {monthData.preferences.showFinancialHealth ? (
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

            {(monthData.preferences.showCategoryChart || monthData.preferences.showTrendChart) ? (
              <section className="grid gap-6 lg:grid-cols-3">
                {monthData.preferences.showCategoryChart ? (
                  <Section>
                    <div className="flex items-center justify-between gap-4"><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">התפלגות הוצאות לפי קטגוריות</h2><span className="rounded-full border border-neutral-200 bg-white px-3 py-1 text-xs font-medium text-neutral-500">Heatmap</span></div>
                    <div className="mx-auto mt-6 h-56 w-56 rounded-full" style={{ background: pieChart }} />
                    <div className="mt-6 space-y-2">
                      {topCategories.length ? topCategories.map(([category, amount]) => <div key={category} className={`flex justify-between rounded-2xl border px-4 py-3 text-sm ${getBudgetHeatColor(category, amount)}`}><span>{category}</span><strong>{SHEKEL.format(amount)}</strong></div>) : <EmptyState title="אין עדיין קטגוריות" text="העלי פירוט אשראי כדי לראות התפלגות צבעונית לפי קטגוריות." />}
                    </div>
                  </Section>
                ) : null}

                {monthData.preferences.showTrendChart ? (
                  <Section className="lg:col-span-2">
                    <div className="flex items-center justify-between gap-4"><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">מגמת הוצאות חודשית</h2><span className="rounded-full border border-neutral-200 bg-white px-3 py-1 text-xs font-medium text-neutral-500">Trend</span></div>
                    <div className="mt-6 flex h-64 items-end gap-3 rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                      {trend.map((item) => <div key={item.month} className="flex flex-1 flex-col items-center gap-2"><div className="w-full rounded-t-xl bg-neutral-800" style={{ height: `${Math.max(4, (item.total / maxTrend) * 200)}px` }} /><span className="text-xs text-neutral-500">{monthLabel(item.month)}</span></div>)}
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
                  <div key={product.id} className="grid gap-3 rounded-[24px] border border-neutral-200 p-4 md:grid-cols-[1fr_140px_120px_150px_150px_44px]">
                    <Field value={product.name} onChange={(event) => updateRow('savingsProducts', product.id, 'name', event.target.value)} />
                    <SelectField value={product.type} onChange={(event) => updateRow('savingsProducts', product.id, 'type', event.target.value)}><option>קרן השתלמות</option><option>פנסיה</option><option>קופת גמל</option><option>חיסכון</option><option>השקעות</option></SelectField>
                    <Field value={product.owner} onChange={(event) => updateRow('savingsProducts', product.id, 'owner', event.target.value)} />
                    <Field type="number" value={product.monthlyDeposit} onChange={(event) => updateRow('savingsProducts', product.id, 'monthlyDeposit', event.target.value)} />
                    <Field type="number" value={product.currentBalance} onChange={(event) => updateRow('savingsProducts', product.id, 'currentBalance', event.target.value)} />
                    <GhostButton onClick={() => removeRow('savingsProducts', product.id)} className="px-0">×</GhostButton>
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
                      <Field value={goal.name} onChange={(event) => updateRow('savingGoals', goal.id, 'name', event.target.value)} className="w-full font-semibold" />
                      <div className="mt-3 grid gap-3">
                        <Field type="number" value={goal.targetAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'targetAmount', event.target.value)} placeholder="יעד" />
                        <Field type="number" value={goal.currentAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'currentAmount', event.target.value)} placeholder="נצבר" />
                        <Field type="number" value={goal.monthlyDeposit} onChange={(event) => updateRow('savingGoals', goal.id, 'monthlyDeposit', event.target.value)} placeholder="הפקדה חודשית" />
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
          <section className="grid gap-6 lg:grid-cols-2">
            <Section>
              <div className="flex items-center justify-between gap-4"><div><h2 className="text-3xl font-semibold tracking-tight text-neutral-950">הכנסות</h2><p className="mt-2 text-sm text-neutral-500">אפשר לצרף תלוש PDF לתיעוד החודש. את סכום הנטו מזינים ידנית כדי שהמערכת תישאר 100% אמינה וללא OCR דמו.</p></div><div className="flex gap-3"><label className="cursor-pointer rounded-xl bg-neutral-950 px-4 py-3 text-sm font-semibold text-white transition hover:bg-neutral-800">צירוף תלוש לחודש<input type="file" accept="application/pdf" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) attachSalarySlipFile(file); }} /></label><PrimaryButton theme={activeTheme} onClick={addIncome}>+ הוספה</PrimaryButton></div></div>
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
              <div className="mt-6 grid gap-4 md:grid-cols-2">{[['owner', 'בעל העסק', 'text'], ['salaryTransferToHousehold', 'העברה / משכורת לחשבון המשותף', 'number'], ['grossRevenue', 'הכנסה עסקית ברוטו', 'number'], ['vatCollected', 'מע״מ שנגבה מלקוחות', 'number'], ['vatPaidOnExpenses', 'מע״מ על הוצאות מוכרות', 'number'], ['incomeTaxAdvance', 'מקדמת מס הכנסה', 'number'], ['nationalInsurance', 'ביטוח לאומי', 'number'], ['businessExpenses', 'הוצאות עסקיות ששולמו החודש', 'number']].map(([field, label, type]) => <label key={field} className="text-sm font-semibold text-neutral-600">{label}<Field type={type} value={monthData.selfEmployed[field]} onChange={(event) => updateSelfEmployedField(field, event.target.value)} className="mt-2 w-full" /></label>)}</div>
              <div className="mt-6 grid gap-4 md:grid-cols-3"><StatCard title="מע״מ צפוי" value={SHEKEL.format(selfEmployedVatDue)} note="נגבה פחות מוכר" /><StatCard title="מס + ביטוח" value={SHEKEL.format(toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance))} note="תשלומי חובה" /><StatCard title="סה״כ עצמאי" value={SHEKEL.format(totalSelfEmployedPayments)} note={includeSelfEmployed ? 'כלול בבית' : 'מחוץ לבית'} /></div>
            </Section>
          </section>
        ) : null}

        {activeTab === 'insights' ? (
          <section className="grid gap-6 lg:grid-cols-2">
            {monthData.preferences.showSmartInsightCards ? <Section><div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between"><div><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">תובנות חכמות</h2><p className="mt-2 text-sm text-neutral-500">תובנות מחושבות ישירות מהנתונים: חריגות, תקציבים, בתי עסק מובילים, חיובים חוזרים ודפוסים חודשיים.</p></div><div className="rounded-full px-4 py-2 text-sm font-semibold" style={{ backgroundColor: activeTheme.soft, color: activeTheme.text }}>מתעדכן אוטומטית</div></div><div className="mt-5 grid gap-4">{realInsights.map((insight, index) => <div key={insight} className="flex items-start gap-4 rounded-[24px] border border-neutral-200 bg-white p-5 shadow-sm"><div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-neutral-100 text-lg font-semibold text-neutral-500">{index % 3 === 0 ? '◔' : index % 3 === 1 ? '▲' : '✦'}</div><div className="flex-1 text-sm leading-7 text-neutral-700 no-orphans">{noSingleWordLine(insight)}</div></div>)}</div></Section> : null}
            {monthData.preferences.showRecurringDetection ? <Section><h2 className="text-2xl font-semibold tracking-tight text-neutral-950">זיהוי חיובים קבועים</h2><p className="mt-2 text-sm text-neutral-500">זיהוי מנויים, ביטוחים, סלולר ושכירות לפי מילות מפתח וחזרה בין חודשים.</p><div className="mt-5 space-y-3">{recurringTransactions.length ? recurringTransactions.map((item) => <div key={item.id} className="flex justify-between rounded-2xl bg-neutral-50 p-4 text-sm"><span>{item.merchant}</span><strong>{SHEKEL.format(item.amount)}</strong></div>) : <EmptyState title="אין עדיין חיובים קבועים" text="העלי פירוטים של כמה חודשים כדי שנוכל לזהות מנויים ותשלומים חוזרים בצורה חכמה." />}</div></Section> : null}
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

              <div className="mt-6 grid gap-6 lg:grid-cols-2">
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
                    <div className="grid gap-3 md:grid-cols-2">
                      <label className="text-sm font-semibold text-neutral-600">
                        משתמש/ת ראשון/ה
                        <Field
                          value={monthData.preferences.primaryPerson}
                          onChange={(event) => updatePreference('primaryPerson', event.target.value)}
                          className="mt-2 w-full"
                        />
                      </label>
                      <label className="text-sm font-semibold text-neutral-600">
                        משתמש/ת שני/ה
                        <Field
                          value={monthData.preferences.secondaryPerson}
                          onChange={(event) => updatePreference('secondaryPerson', event.target.value)}
                          className="mt-2 w-full"
                        />
                      </label>
                    </div>
                  </div>
                </div>

                <div className="rounded-[24px] border border-neutral-200 bg-neutral-50 p-5">
                  <h3 className="text-lg font-semibold text-neutral-950">יעדים חודשיים</h3>
                  <div className="mt-4 grid gap-3 md:grid-cols-2">
                    <label className="text-sm font-semibold text-neutral-600">
                      יעד הוצאות חודשי
                      <Field
                        type="number"
                        value={monthData.preferences.monthlyBudgetTarget}
                        onChange={(event) => updatePreference('monthlyBudgetTarget', event.target.value)}
                        className="mt-2 w-full"
                      />
                    </label>
                    <label className="text-sm font-semibold text-neutral-600">
                      יעד שיעור חיסכון באחוזים
                      <Field
                        type="number"
                        value={monthData.preferences.savingsRateTarget}
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
                    ['showCategoryChart', 'גרף קטגוריות'],
                    ['showTrendChart', 'גרף מגמה'],
                    ['showSmartInsightCards', 'כרטיסי תובנות'],
                    ['showRecurringDetection', 'זיהוי חיובים קבועים'],
                  ].map(([field, label]) => (
                    <label key={field} className="flex cursor-pointer items-center justify-between gap-4 rounded-2xl border border-neutral-200 bg-neutral-50 px-4 py-3 text-sm font-semibold text-neutral-700">
                      <span>{label}</span>
                      <input
                        type="checkbox"
                        checked={Boolean(monthData.preferences[field])}
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
                      const isSelected = monthData.preferences.themeMood === themeName;
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
                      const isSelected = monthData.preferences.financialMode === mode;
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

              <div className="mt-6 grid gap-6 lg:grid-cols-2">
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
                          checked={Boolean(monthData.preferences.notifications?.[field])}
                          onChange={(event) => updatePreference('notifications', { ...monthData.preferences.notifications, [field]: event.target.checked })}
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
                      const isSelected = monthData.preferences.syncMode === mode;
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
                <div className="grid gap-3 md:grid-cols-2">
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
