'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';

const STORAGE_KEY = 'family-finance-os-full-intelligence-v1';
const SUPABASE_PROFILE_ID = 'default-household';

const SHEKEL = new Intl.NumberFormat('he-IL', {
  style: 'currency',
  currency: 'ILS',
  maximumFractionDigits: 0,
});

const defaultExpenseCategories = [
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

const categoryBudgets = {
  'סופר / מזון': 4000,
  'מסעדות / וולט': 800,
  'תחבורה / דלק': 1800,
  קניות: 1200,
  בריאות: 800,
  'בידור / מנויים': 600,
  אחר: 1000,
};

const baseMerchantCategoryMap = {
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

const recurringKeywords = [
  'netflix',
  'spotify',
  'icloud',
  'google',
  'apple',
  'cellcom',
  'partner',
  'pelephone',
  'hot',
  'yes',
  'ביטוח',
  'הראל',
  'מגדל',
  'כלל',
  'סלקום',
  'פרטנר',
  'פלאפון',
  'שכירות',
];

function makeId(prefix = 'id') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function getCurrentMonthKey() {
  return new Date().toISOString().slice(0, 7);
}

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const cleaned = String(value || '').replace(/[₪,]/g, '').replace(/\s/g, '').trim();
  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
}

function monthLabel(monthKey) {
  if (!monthKey) return '';
  const [year, month] = monthKey.split('-');
  return `${month}/${year}`;
}

function formatPercent(value) {
  return `${Math.round(value)}%`;
}

function normalizeMerchantName(merchant = '') {
  return String(merchant)
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[.,:;|()\[\]{}]/g, '')
    .trim();
}

function detectCategory(merchant = '', learnedRules = {}) {
  const lower = normalizeMerchantName(merchant);

  for (const [key, category] of Object.entries(learnedRules)) {
    if (lower.includes(normalizeMerchantName(key))) return category;
  }

  for (const [key, category] of Object.entries(baseMerchantCategoryMap)) {
    if (lower.includes(normalizeMerchantName(key))) return category;
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
  const hasHeader =
    firstRow.includes('date') ||
    firstRow.includes('תאריך') ||
    firstRow.includes('amount') ||
    firstRow.includes('סכום') ||
    firstRow.includes('merchant') ||
    firstRow.includes('בית עסק') ||
    firstRow.includes('שם בית עסק');

  const dataRows = hasHeader ? rows.slice(1) : rows;

  return dataRows
    .map((row) => {
      const compactRow = row.map((cell) => String(cell || '').trim());
      const date = compactRow[0] || '';
      const merchant = compactRow[1] || compactRow[0] || 'עסקה';
      const rawAmount = compactRow[2] || compactRow[compactRow.length - 1] || '0';
      const amount = Math.abs(toNumber(rawAmount));

      return {
        id: makeId('tx'),
        date,
        merchant,
        amount,
        category: detectCategory(merchant, learnedRules),
      };
    })
    .filter((transaction) => transaction.amount > 0);
}

function parseCsvText(text, learnedRules = {}) {
  const cleanText = String(text || '').replace(/\r/g, '');
  const lines = cleanText.split('\n').map((line) => line.trim()).filter(Boolean);
  if (lines.length === 0) return [];
  const rows = lines.map(splitCsvLine).filter((row) => row.length > 0);
  return normalizeImportedRows(rows, learnedRules);
}

function parseExcelArrayBuffer(buffer, learnedRules = {}) {
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) return [];

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: '',
  });

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

function calculateFinancialHealthScore(transactions) {
  if (!transactions.length) return null;

  const total = transactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const categoryTotals = getCategoryTotals(transactions);
  const merchantTotals = getMerchantTotals(transactions);
  const uncategorizedAmount = categoryTotals['אחר'] || 0;
  const largestTransaction = Math.max(...transactions.map((item) => toNumber(item.amount)));
  const largestMerchantAmount = Math.max(...Object.values(merchantTotals).map(toNumber));

  let score = 100;

  if (uncategorizedAmount / total > 0.2) score -= 15;
  if (largestTransaction / total > 0.25) score -= 12;
  if (largestMerchantAmount / total > 0.35) score -= 10;

  Object.entries(categoryBudgets).forEach(([category, budget]) => {
    const spent = categoryTotals[category] || 0;
    if (spent > budget) score -= 8;
    else if (spent >= budget * 0.8) score -= 4;
  });

  return Math.max(0, Math.min(100, score));
}

function detectRecurringTransactions(transactions, historicalMonths = {}, selectedMonth = '') {
  const historicalMerchants = new Set();

  Object.entries(historicalMonths).forEach(([month, data]) => {
    if (month === selectedMonth) return;
    (data.creditCards || []).forEach((card) => {
      (card.transactions || []).forEach((transaction) => {
        historicalMerchants.add(normalizeMerchantName(transaction.merchant));
      });
    });
  });

  return transactions.filter((transaction) => {
    const normalized = normalizeMerchantName(transaction.merchant);
    const keywordHit = recurringKeywords.some((keyword) => normalized.includes(normalizeMerchantName(keyword)));
    const historicalHit = historicalMerchants.has(normalized);
    return keywordHit || historicalHit;
  });
}

function getMonthlyTrend(months) {
  return Object.entries(months)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([month, data]) => {
      const creditTotal = (data.creditCards || []).reduce(
        (sum, card) => sum + (card.transactions || []).reduce((innerSum, item) => innerSum + toNumber(item.amount), 0),
        0
      );
      const manualTotal = (data.manualExpenses || []).reduce((sum, item) => sum + toNumber(item.amount), 0);
      return { month, total: creditTotal + manualTotal };
    });
}

function buildRealInsights(transactions, recurringTransactions = [], totalIncome = 0) {
  if (!transactions.length) return ['עדיין אין נתוני אשראי. העלו CSV או Excel בתוך אזור פירוט האשראי כדי לקבל תובנות.'];

  const total = transactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const categoryTotals = getCategoryTotals(transactions);
  const merchantTotals = getMerchantTotals(transactions);
  const sortedCategories = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]);
  const sortedMerchants = Object.entries(merchantTotals).sort((a, b) => b[1] - a[1]);
  const averageTransaction = total / transactions.length;
  const healthScore = calculateFinancialHealthScore(transactions);
  const insights = [];

  insights.push(`ציון בריאות הוצאות אשראי: ${healthScore}/100. הציון מחושב לפי חריגות תקציב, ריכוזיות, עסקאות גדולות וסיווגים חסרים.`);

  const [topCategory, topCategoryAmount] = sortedCategories[0] || [];
  if (topCategory) insights.push(`הקטגוריה הגדולה ביותר באשראי היא ${topCategory}: ${SHEKEL.format(topCategoryAmount)}, שהם ${Math.round((topCategoryAmount / total) * 100)}% מהחיובים.`);

  const [topMerchant, topMerchantAmount] = sortedMerchants[0] || [];
  if (topMerchant) insights.push(`בית העסק הדומיננטי ביותר הוא ${topMerchant}: ${SHEKEL.format(topMerchantAmount)}, כלומר ${Math.round((topMerchantAmount / total) * 100)}% מהחיובים.`);

  insights.push(`גובה עסקת אשראי ממוצעת: ${SHEKEL.format(averageTransaction)}.`);

  if (totalIncome > 0) insights.push(`חיובי האשראי הם ${formatPercent((total / totalIncome) * 100)} מההכנסה שהוזנה החודש.`);

  Object.entries(categoryBudgets).forEach(([category, budget]) => {
    const spent = categoryTotals[category] || 0;
    if (spent > budget) insights.push(`${category} חרגה מהתקציב: ${SHEKEL.format(spent)} מתוך ${SHEKEL.format(budget)}. חריגה של ${SHEKEL.format(spent - budget)}.`);
    else if (spent >= budget * 0.8) insights.push(`${category} מתקרבת לתקציב: ${SHEKEL.format(spent)} מתוך ${SHEKEL.format(budget)}.`);
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

  if (healthScore >= 85) insights.push('לא נמצאו חריגות משמעותיות לפי החוקים שהוגדרו. נראה שהחודש מאוזן יחסית באשראי.');

  return insights;
}

async function loadFinanceStateFromSupabase() {
  const { data, error } = await supabase
    .from('finance_app_state')
    .select('months, learned_rules')
    .eq('profile_id', SUPABASE_PROFILE_ID)
    .maybeSingle();

  if (error) throw error;
  return data;
}

async function saveFinanceStateToSupabase(months, learnedRules) {
  const { error } = await supabase
    .from('finance_app_state')
    .upsert({
      profile_id: SUPABASE_PROFILE_ID,
      months,
      learned_rules: learnedRules,
      updated_at: new Date().toISOString(),
    });

  if (error) throw error;
}

function runSmokeTests() {
  console.assert(toNumber('₪1,250') === 1250, 'currency parsing failed');
  console.assert(detectCategory('Wolt TLV') === 'מסעדות / וולט', 'wolt category failed');
  console.assert(detectCategory('My Shop', { shop: 'קניות' }) === 'קניות', 'learned rule failed');
  console.assert(splitCsvLine('a,b,c').length === 3, 'csv split failed');
  console.assert(parseCsvText('date,merchant,amount\n2026-01-01,Wolt,55').length === 1, 'csv parse failed');
  console.assert(getCategoryTotals([{ category: 'קניות', amount: 10 }, { category: 'קניות', amount: 20 }]).קניות === 30, 'category totals failed');
  console.assert(buildRealInsights([{ merchant: 'Wolt', category: 'מסעדות / וולט', amount: 900 }]).some((insight) => insight.includes('חרגה')), 'real budget insight failed');
}

if (typeof window !== 'undefined') runSmokeTests();

function createDefaultMonth() {
  return {
    dashboardTitle: 'מערכת פיננסית משפחתית 🇮🇱',
    emergencyFund: 0,
    lastSalaryImport: '',
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
      grossRevenue: 0,
      vatCollected: 0,
      vatPaidOnExpenses: 0,
      incomeTaxAdvance: 0,
      nationalInsurance: 0,
      businessExpenses: 0,
    },
  };
}

function normalizeMonthData(data) {
  const base = createDefaultMonth();
  const safe = data || {};

  return {
    ...base,
    ...safe,
    incomes: safe.incomes || base.incomes,
    manualExpenses: safe.manualExpenses || base.manualExpenses,
    savingsProducts: safe.savingsProducts || base.savingsProducts,
    savingGoals: safe.savingGoals || base.savingGoals,
    creditCards: (safe.creditCards || base.creditCards).map((card) => ({ transactions: [], pendingTransactions: [], importedFile: '', ...card })),
    selfEmployed: { ...base.selfEmployed, ...(safe.selfEmployed || {}) },
  };
}

function StatCard({ title, value, note }) {
  return (
    <div className="rounded-3xl border border-slate-100 bg-white p-6 shadow-sm">
      <div className="text-sm font-medium text-slate-500">{title}</div>
      <div className="mt-3 text-3xl font-bold text-slate-900">{value}</div>
      <div className="mt-2 text-sm text-[#5f7d66]">{note}</div>
    </div>
  );
}

function Section({ children, className = '' }) {
  return <section className={`rounded-[32px] border border-slate-200 bg-white p-7 shadow-sm ${className}`}>{children}</section>;
}

export default function PersonalIsraeliFamilyFinanceDashboard() {
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonthKey());
  const [months, setMonths] = useState(() => {
    try {
      const saved = window.localStorage.getItem(STORAGE_KEY);
      return saved ? JSON.parse(saved) : { [getCurrentMonthKey()]: createDefaultMonth() };
    } catch {
      return { [getCurrentMonthKey()]: createDefaultMonth() };
    }
  });
  const [learnedRules, setLearnedRules] = useState(() => {
    try {
      return JSON.parse(window.localStorage.getItem(`${STORAGE_KEY}-rules`) || '{}');
    } catch {
      return {};
    }
  });
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('הכול');
  const [minAmount, setMinAmount] = useState('');
  const [maxAmount, setMaxAmount] = useState('');
  const [aiInsights, setAiInsights] = useState([]);
  const [aiStatus, setAiStatus] = useState('idle');
  const [aiError, setAiError] = useState('');
  const [cloudStatus, setCloudStatus] = useState('טוען מהענן…');
  const [hasLoadedCloud, setHasLoadedCloud] = useState(false);

  const monthData = normalizeMonthData(months[selectedMonth]);

  useEffect(() => {
    async function loadCloudState() {
      try {
        const data = await loadFinanceStateFromSupabase();

        if (data?.months) {
          setMonths(data.months);
        }

        if (data?.learned_rules) {
          setLearnedRules(data.learned_rules);
        }

        setCloudStatus(data?.months ? 'מסונכרן מהענן' : 'אין עדיין נתוני ענן, עובדים מקומית');
      } catch {
        setCloudStatus('ענן לא זמין כרגע, עובדים מקומית');
      } finally {
        setHasLoadedCloud(true);
      }
    }

    loadCloudState();
  }, []);

  useEffect(() => {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(months));
      window.localStorage.setItem(`${STORAGE_KEY}-rules`, JSON.stringify(learnedRules));
    } catch {
      // Local storage can be blocked. The app still works in memory.
    }
  }, [months, learnedRules]);

  useEffect(() => {
    if (!hasLoadedCloud) return;

    const saveTimeout = window.setTimeout(async () => {
      try {
        await saveFinanceStateToSupabase(months, learnedRules);
        setCloudStatus('נשמר בענן');
      } catch {
        setCloudStatus('לא נשמר בענן, נשמר מקומית');
      }
    }, 900);

    return () => window.clearTimeout(saveTimeout);
  }, [months, learnedRules, hasLoadedCloud]);

  function setSelectedMonthData(nextData) {
    setMonths((current) => ({ ...current, [selectedMonth]: nextData }));
  }

  function ensureMonth(monthKey) {
    setSelectedMonth(monthKey);
    if (!months[monthKey]) {
      setMonths((current) => ({ ...current, [monthKey]: createDefaultMonth() }));
    }
  }

  function updateMonthField(field, value) {
    setSelectedMonthData({ ...monthData, [field]: value });
  }

  function updateRow(section, id, field, value) {
    const numericFields = ['amount', 'monthlyDeposit', 'currentBalance', 'targetAmount', 'currentAmount'];
    setSelectedMonthData({
      ...monthData,
      [section]: monthData[section].map((row) =>
        row.id === id ? { ...row, [field]: numericFields.includes(field) ? toNumber(value) : value } : row
      ),
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
    const numericFields = ['grossRevenue', 'vatCollected', 'vatPaidOnExpenses', 'incomeTaxAdvance', 'nationalInsurance', 'businessExpenses'];
    setSelectedMonthData({
      ...monthData,
      selfEmployed: { ...monthData.selfEmployed, [field]: numericFields.includes(field) ? toNumber(value) : value },
    });
  }

  async function importSalarySlipFile(file) {
    setSelectedMonthData({ ...monthData, lastSalaryImport: file.name });

    try {
      const formData = new FormData();
      formData.append('file', file);

      const response = await fetch('/api/salary-slip-ocr', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error('OCR failed');
      }

      const data = await response.json();

      if (data?.salaryNet) {
        const updatedIncomes = [...monthData.incomes];

        if (updatedIncomes[0]) {
          updatedIncomes[0] = {
            ...updatedIncomes[0],
            amount: toNumber(data.salaryNet),
          };
        }

        setSelectedMonthData({
          ...monthData,
          lastSalaryImport: file.name,
          incomes: updatedIncomes,
        });
      }
    } catch {
      alert('התלוש נשמר, אבל OCR עדיין לא מחובר בשרת. אפשר להזין ידנית בינתיים.');
    }
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

    if (lower.endsWith('.csv')) {
      importedTransactions = parseCsvText(await file.text(), learnedRules);
    } else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
      importedTransactions = parseExcelArrayBuffer(await file.arrayBuffer(), learnedRules);
    } else {
      alert('נא להעלות CSV או Excel');
      return;
    }

    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId ? { ...card, importedFile: file.name, pendingTransactions: importedTransactions } : card
      ),
    });
    setAiInsights([]);
    setAiError('');
  }

  function updatePendingTransaction(cardId, transactionId, field, value) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId
          ? {
              ...card,
              pendingTransactions: (card.pendingTransactions || []).map((transaction) =>
                transaction.id === transactionId ? { ...transaction, [field]: field === 'amount' ? toNumber(value) : value } : transaction
              ),
            }
          : card
      ),
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

    if (changedTransaction) {
      const normalizedMerchant = normalizeMerchantName(changedTransaction.merchant);
      setLearnedRules((current) => ({ ...current, [normalizedMerchant]: newCategory }));
      setSelectedMonthData({
        ...monthData,
        creditCards: updatedCards.map((card) => ({
          ...card,
          transactions: (card.transactions || []).map((transaction) =>
            normalizeMerchantName(transaction.merchant) === normalizedMerchant ? { ...transaction, category: newCategory } : transaction
          ),
        })),
      });
    }
  }

  function removePendingTransaction(cardId, transactionId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId ? { ...card, pendingTransactions: (card.pendingTransactions || []).filter((transaction) => transaction.id !== transactionId) } : card
      ),
    });
  }

  function approvePendingTransactions(cardId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId ? { ...card, transactions: [...(card.transactions || []), ...(card.pendingTransactions || [])], pendingTransactions: [] } : card
      ),
    });
  }

  function addTransaction(cardId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId ? { ...card, transactions: [...(card.transactions || []), { id: makeId('tx'), date: '', merchant: 'עסקה חדשה', category: 'אחר', amount: 0 }] } : card
      ),
    });
  }

  function removeTransaction(cardId, transactionId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId ? { ...card, transactions: (card.transactions || []).filter((transaction) => transaction.id !== transactionId) } : card
      ),
    });
  }

  async function generateAiInsights() {
    if (!allCreditTransactions.length) {
      setAiError('צריך להעלות קובץ אשראי לפני שמריצים AI.');
      return;
    }

    setAiStatus('loading');
    setAiError('');

    try {
      const response = await fetch('/api/ai-insights', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ transactions: allCreditTransactions, categoryBudgets, month: selectedMonth }),
      });

      if (!response.ok) throw new Error('AI endpoint failed');
      const data = await response.json();
      setAiInsights(Array.isArray(data.insights) ? data.insights : []);
      setAiStatus('ready');
    } catch {
      setAiStatus('error');
      setAiError('OpenAI לא זמין כרגע דרך השרת. מוצגות תובנות מקומיות אמיתיות שמחושבות מהנתונים.');
    }
  }

  const totalIncome = monthData.incomes.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const allCreditTransactions = useMemo(() => monthData.creditCards.flatMap((card) => card.transactions || []), [monthData.creditCards]);
  const totalCreditCards = allCreditTransactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalManualExpenses = monthData.manualExpenses.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const totalSavingsProducts = monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const totalSavingGoals = monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0);
  const totalPlannedSavings = totalSavingsProducts + totalSavingGoals;
  const selfEmployedVatDue = Math.max(0, toNumber(monthData.selfEmployed.vatCollected) - toNumber(monthData.selfEmployed.vatPaidOnExpenses));
  const totalSelfEmployedPayments = selfEmployedVatDue + toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance) + toNumber(monthData.selfEmployed.businessExpenses);
  const totalExpenses = totalCreditCards + totalManualExpenses + totalPlannedSavings + totalSelfEmployedPayments;
  const monthlySavings = totalIncome - totalExpenses;
  const savingsRate = totalIncome ? (monthlySavings / totalIncome) * 100 : 0;
  const emergencyMonths = toNumber(monthData.emergencyFund) / (totalExpenses || 1);
  const totalAssets = toNumber(monthData.emergencyFund) + monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.currentBalance), 0) + monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.currentAmount), 0);

  const categoryTotals = useMemo(() => getCategoryTotals(allCreditTransactions), [allCreditTransactions]);
  const merchantTotals = useMemo(() => getMerchantTotals(allCreditTransactions), [allCreditTransactions]);
  const recurringTransactions = useMemo(() => detectRecurringTransactions(allCreditTransactions, months, selectedMonth), [allCreditTransactions, months, selectedMonth]);
  const realInsights = useMemo(() => buildRealInsights(allCreditTransactions, recurringTransactions, totalIncome), [allCreditTransactions, recurringTransactions, totalIncome]);
  const healthScore = useMemo(() => calculateFinancialHealthScore(allCreditTransactions), [allCreditTransactions]);
  const trend = useMemo(() => getMonthlyTrend(months), [months]);
  const maxTrend = Math.max(1, ...trend.map((item) => item.total));
  const burnRate = trend.length ? trend.reduce((sum, item) => sum + item.total, 0) / trend.length : 0;
  const cashFlow = totalPlannedSavings;
  const topMerchants = useMemo(() => Object.entries(merchantTotals).sort((a, b) => b[1] - a[1]).slice(0, 5), [merchantTotals]);
  const topCategories = useMemo(() => Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]).slice(0, 6), [categoryTotals]);

  const filteredTransactions = useMemo(() => {
    const normalizedSearch = normalizeMerchantName(searchTerm);
    return allCreditTransactions.filter((transaction) => {
      const merchantMatch = normalizeMerchantName(transaction.merchant).includes(normalizedSearch);
      const categoryMatch = categoryFilter === 'הכול' || transaction.category === categoryFilter;
      const amount = toNumber(transaction.amount);
      const minMatch = minAmount === '' || amount >= toNumber(minAmount);
      const maxMatch = maxAmount === '' || amount <= toNumber(maxAmount);
      return merchantMatch && categoryMatch && minMatch && maxMatch;
    });
  }, [allCreditTransactions, searchTerm, categoryFilter, minAmount, maxAmount]);

  const pieChart = topCategories.length
    ? `conic-gradient(${topCategories
        .map(([, amount], index) => {
          const start = topCategories.slice(0, index).reduce((sum, [, value]) => sum + value, 0) / (totalCreditCards || 1);
          const end = topCategories.slice(0, index + 1).reduce((sum, [, value]) => sum + value, 0) / (totalCreditCards || 1);
          const colors = ['#7a9b76', '#9ebc8a', '#d0b88a', '#b9856f', '#8aa0a2', '#c9c0a8'];
          return `${colors[index % colors.length]} ${start * 100}% ${end * 100}%`;
        })
        .join(', ')})`
    : 'conic-gradient(#eef3ef 0% 100%)';

  return (
    <div dir="rtl" className="min-h-screen bg-gradient-to-br from-[#f4f1ea] via-[#faf8f4] to-[#eef3ef] p-6 text-right text-slate-800">
      <div className="mx-auto max-w-7xl space-y-8">
        <section className="overflow-hidden rounded-[32px] border border-[#d9d4c7] bg-[#fcfbf8] shadow-[0_10px_40px_rgba(0,0,0,0.05)]">
          <div className="bg-gradient-to-l from-[#6b8f71] to-[#9ebc8a] p-8 text-[#f8f6f1]">
            <div className="flex flex-col gap-6 lg:flex-row lg:items-center lg:justify-between">
              <div>
                <input value={monthData.dashboardTitle} onChange={(event) => updateMonthField('dashboardTitle', event.target.value)} className="w-full max-w-3xl rounded-2xl border border-white/20 bg-white/10 px-4 py-3 text-4xl font-bold tracking-tight text-white outline-none transition placeholder:text-white/70 focus:bg-white/20 md:text-5xl" placeholder="שם הדשבורד המשפחתי" />
                <p className="mt-4 max-w-2xl text-base text-[#f3f0e8]">ממלאים הכנסות, הוצאות, אשראי, עצמאי, קרנות ויעדים. המערכת מחשבת תזרים, חיסכון ותובנות אמיתיות.</p>
                <div className="mt-4 inline-flex rounded-2xl bg-white/15 px-4 py-2 text-sm font-semibold text-white">
                  {cloudStatus}
                </div>
              </div>
              <div className="rounded-3xl bg-[#f7f5ef]/20 p-6 backdrop-blur-sm">
                <label className="text-sm uppercase tracking-widest text-[#f3f0e8]">חודש</label>
                <input type="month" value={selectedMonth} onChange={(event) => ensureMonth(event.target.value)} className="mt-2 w-full rounded-2xl border-0 bg-white px-4 py-3 text-lg font-bold text-slate-900 shadow-sm" />
              </div>
            </div>
          </div>
          <div className="grid gap-5 p-6 md:grid-cols-2 xl:grid-cols-7">
            <StatCard title="סה״כ הכנסות" value={SHEKEL.format(totalIncome)} note="כל מקורות ההכנסה" />
            <StatCard title="סה״כ הוצאות" value={SHEKEL.format(totalExpenses)} note={`${totalIncome ? formatPercent((totalExpenses / totalIncome) * 100) : '0%'} מההכנסה`} />
            <StatCard title="סה״כ אשראי" value={SHEKEL.format(totalCreditCards)} note="מכרטיסי האשראי" />
            <StatCard title="עצמאי" value={SHEKEL.format(totalSelfEmployedPayments)} note="מע״מ, מס וביטוח לאומי" />
            <StatCard title="חסכונות" value={SHEKEL.format(totalPlannedSavings)} note="קרנות, פנסיה ויעדים" />
            <StatCard title="יתרה אחרי הכול" value={SHEKEL.format(monthlySavings)} note={`${formatPercent(savingsRate)} שיעור חיסכון`} />
            <StatCard title="שווי שהוזן" value={SHEKEL.format(totalAssets)} note={`${emergencyMonths.toFixed(1)} חודשי חירום`} />
          </div>
        </section>

        <section className="grid gap-6 lg:grid-cols-2">
          <Section>
            <div className="flex items-center justify-between gap-4">
              <div><h2 className="text-3xl font-bold">הכנסות</h2><p className="mt-2 text-sm text-slate-500">אפשר להעלות גם תלוש PDF. כשה־OCR יחובר בשרת, נטו מהתלוש ייכנס אוטומטית להכנסות.</p></div>
              <div className="flex gap-3"><label className="cursor-pointer rounded-2xl bg-[#7c9780] px-4 py-3 text-sm font-semibold text-white shadow-sm">שמירת תלוש לחודש<input type="file" accept="application/pdf" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) importSalarySlipFile(file); }} /></label><button onClick={addIncome} className="rounded-2xl bg-[#7a9b76] px-4 py-3 text-sm font-semibold text-white shadow-sm">+ הוספה</button></div>
            </div>
            {monthData.lastSalaryImport ? <div className="mt-4 rounded-2xl bg-[#f3f5ef] p-4 text-sm text-[#4f6854]">תלוש שנקלט לחודש: {monthData.lastSalaryImport}</div> : null}
            <div className="mt-6 space-y-3">{monthData.incomes.map((income) => <div key={income.id} className="grid gap-3 rounded-2xl border border-slate-100 p-4 md:grid-cols-[1fr_160px_44px]"><input value={income.name} onChange={(event) => updateRow('incomes', income.id, 'name', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" /><input type="number" value={income.amount} onChange={(event) => updateRow('incomes', income.id, 'amount', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" /><button onClick={() => removeRow('incomes', income.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button></div>)}</div>
          </Section>

          <Section>
            <h2 className="text-3xl font-bold">עצמאי: מע״מ, מס הכנסה וביטוח לאומי</h2>
            <p className="mt-2 text-sm text-slate-500">אזור לאורן כעצמאי. נספר בנפרד כדי שלא יתערבב עם הוצאות הבית.</p>
            <div className="mt-6 grid gap-4 md:grid-cols-2">{[
              ['owner', 'בעל העסק', 'text'], ['grossRevenue', 'הכנסה עסקית ברוטו', 'number'], ['vatCollected', 'מע״מ שנגבה מלקוחות', 'number'], ['vatPaidOnExpenses', 'מע״מ על הוצאות מוכרות', 'number'], ['incomeTaxAdvance', 'מקדמת מס הכנסה', 'number'], ['nationalInsurance', 'ביטוח לאומי', 'number'], ['businessExpenses', 'הוצאות עסקיות ששולמו החודש', 'number'],
            ].map(([field, label, type]) => <label key={field} className="text-sm font-bold text-slate-600">{label}<input type={type} value={monthData.selfEmployed[field]} onChange={(event) => updateSelfEmployedField(field, event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" /></label>)}</div>
            <div className="mt-6 grid gap-4 md:grid-cols-3"><StatCard title="מע״מ צפוי" value={SHEKEL.format(selfEmployedVatDue)} note="נגבה פחות מוכר" /><StatCard title="מס + ביטוח" value={SHEKEL.format(toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance))} note="תשלומי חובה" /><StatCard title="סה״כ עצמאי" value={SHEKEL.format(totalSelfEmployedPayments)} note="כולל הוצאות עסקיות" /></div>
          </Section>
        </section>

        <Section>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-3xl font-bold">סיכום כרטיסי אשראי</h2><p className="mt-2 text-sm text-slate-500">כאן מעלים CSV/Excel לכל כרטיס, בודקים קטגוריות, ואז מאשרים הכנסה להוצאות.</p></div><button onClick={addCreditCard} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת כרטיס</button></div>
          <div className="mt-7 grid gap-6 xl:grid-cols-2">{monthData.creditCards.map((card) => {
            const cardTotal = (card.transactions || []).reduce((sum, item) => sum + toNumber(item.amount), 0);
            return <div key={card.id} className="rounded-3xl border border-[#d8e2d2] bg-[#f3f5ef]/60 p-5"><div className="grid gap-3 md:grid-cols-[1fr_1fr_110px_44px]"><input value={card.name} onChange={(event) => updateCreditCard(card.id, 'name', event.target.value)} className="rounded-xl border border-[#d8e2d2] bg-white px-4 py-3 text-sm" placeholder="שם הכרטיס" /><input value={card.owner} onChange={(event) => updateCreditCard(card.id, 'owner', event.target.value)} className="rounded-xl border border-[#d8e2d2] bg-white px-4 py-3 text-sm" placeholder="בעל/ת הכרטיס" /><div className="rounded-xl bg-white px-4 py-3 text-sm font-bold text-[#5d755f]">{SHEKEL.format(cardTotal)}</div><button onClick={() => removeCreditCard(card.id)} className="rounded-xl bg-white text-sm font-bold text-slate-500">×</button></div>
              <div className="mt-5 rounded-2xl border border-dashed border-[#d8e2d2] bg-white p-4"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><div className="text-sm font-bold text-[#4f6854]">העלאת CSV / Excel של פירוט אשראי</div><div className="mt-1 text-xs text-slate-500">העלאה נמצאת כאן, בתוך הכרטיס הרלוונטי.</div></div><label className="cursor-pointer rounded-2xl bg-[#7a9b76] px-4 py-3 text-sm font-semibold text-white shadow-sm">העלאת קובץ<input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) importCreditFile(card.id, file); }} /></label></div>{card.importedFile ? <div className="mt-3 rounded-xl bg-[#f3f5ef] px-4 py-3 text-sm text-[#4f6854]">נקלט קובץ: <strong>{card.importedFile}</strong></div> : null}</div>
              {(card.pendingTransactions || []).length > 0 ? <div className="mt-4 rounded-2xl border border-[#d8e2d2] bg-[#fcfbf8] p-4"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><div className="text-sm font-bold text-[#4f6854]">עסקאות שזוהו לאישור</div><div className="mt-1 text-xs text-slate-500">בדקו סכומים וקטגוריות לפני שהן נכנסות להוצאות.</div></div><button onClick={() => approvePendingTransactions(card.id)} className="rounded-2xl bg-[#7a9b76] px-4 py-3 text-sm font-semibold text-white shadow-sm">אשר והכנס להוצאות</button></div><div className="mt-4 overflow-x-auto rounded-2xl border border-[#d8e2d2] bg-white"><div className="min-w-[760px]"><div className="grid grid-cols-[120px_1fr_170px_140px_44px] bg-[#e5eee2]/60 px-4 py-3 text-xs font-bold text-[#4f6854]"><div>תאריך</div><div>בית עסק</div><div>קטגוריה</div><div>סכום</div><div></div></div>{(card.pendingTransactions || []).map((transaction) => <div key={transaction.id} className="grid grid-cols-[120px_1fr_170px_140px_44px] gap-3 border-t border-[#d8e2d2] p-3"><input value={transaction.date || ''} onChange={(event) => updatePendingTransaction(card.id, transaction.id, 'date', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm" /><input value={transaction.merchant} onChange={(event) => updatePendingTransaction(card.id, transaction.id, 'merchant', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm" /><select value={transaction.category} onChange={(event) => updatePendingTransaction(card.id, transaction.id, 'category', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm">{defaultExpenseCategories.map((category) => <option key={category}>{category}</option>)}</select><input type="number" value={transaction.amount} onChange={(event) => updatePendingTransaction(card.id, transaction.id, 'amount', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm" /><button onClick={() => removePendingTransaction(card.id, transaction.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button></div>)}</div></div></div> : null}
              <div className="mt-5 overflow-x-auto rounded-2xl border border-[#d8e2d2] bg-white"><div className="min-w-[700px]"><div className="grid grid-cols-[130px_1fr_170px_140px_44px] bg-[#e5eee2]/60 px-4 py-3 text-xs font-bold text-[#4f6854]"><div>תאריך</div><div>עסקה</div><div>קטגוריה לומדת</div><div>סכום</div><div></div></div>{(card.transactions || []).map((transaction) => <div key={transaction.id} className="grid grid-cols-[130px_1fr_170px_140px_44px] gap-3 border-t border-[#d8e2d2] p-3"><div className="px-2 py-3 text-sm text-slate-500">{transaction.date}</div><input value={transaction.merchant} readOnly className="rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm" /><select value={transaction.category} onChange={(event) => updateTransactionCategory(transaction.id, event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm">{defaultExpenseCategories.map((category) => <option key={category}>{category}</option>)}</select><div className="px-3 py-3 text-sm font-bold">{SHEKEL.format(transaction.amount)}</div><button onClick={() => removeTransaction(card.id, transaction.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button></div>)}</div></div><button onClick={() => addTransaction(card.id)} className="mt-4 rounded-2xl bg-[#4d5b52] px-4 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת עסקה</button>
            </div>;
          })}</div>
        </Section>

        <section className="grid gap-6 lg:grid-cols-3">
          <Section><h2 className="text-2xl font-bold text-[#4d5b52]">Pie Chart לפי קטגוריות אשראי</h2><div className="mx-auto mt-6 h-56 w-56 rounded-full" style={{ background: pieChart }} /><div className="mt-6 space-y-2">{topCategories.map(([category, amount]) => <div key={category} className="flex justify-between rounded-xl bg-slate-50 px-4 py-3 text-sm"><span>{category}</span><strong>{SHEKEL.format(amount)}</strong></div>)}</div></Section>
          <Section className="lg:col-span-2"><h2 className="text-2xl font-bold text-[#4d5b52]">Trend Line לפי חודשים</h2><div className="mt-6 flex h-64 items-end gap-3 rounded-3xl bg-[#eef3ef] p-5">{trend.map((item) => <div key={item.month} className="flex flex-1 flex-col items-center gap-2"><div className="w-full rounded-t-xl bg-[#7a9b76]" style={{ height: `${Math.max(4, (item.total / maxTrend) * 200)}px` }} /><span className="text-xs text-slate-500">{monthLabel(item.month)}</span></div>)}</div><div className="mt-4 rounded-2xl bg-white p-4 text-sm text-slate-600">Burn Rate ממוצע: <strong>{SHEKEL.format(burnRate)}</strong> | Cash Flow לחיסכון: <strong>{SHEKEL.format(cashFlow)}</strong></div></Section>
        </section>

        <section className="grid gap-6 lg:grid-cols-2">
          <Section><div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between"><div><h2 className="text-2xl font-bold text-[#4d5b52]">תובנות AI אמיתיות</h2><p className="mt-2 text-sm text-slate-500">תובנות מקומיות מחושבות מהנתונים. כפתור AI מפעיל OpenAI אמיתי דרך השרת אחרי שה־API key מוגדר ב־Vercel.</p></div><button onClick={generateAiInsights} disabled={aiStatus === 'loading' || allCreditTransactions.length === 0} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm disabled:opacity-50">{aiStatus === 'loading' ? 'מנתח…' : 'נתח עם AI'}</button></div>{aiError ? <div className="mt-4 rounded-2xl bg-amber-50 p-4 text-sm text-amber-800">{aiError}</div> : null}<div className="mt-5 space-y-3">{(aiInsights.length > 0 ? aiInsights : realInsights).map((insight) => <div key={insight} className="rounded-2xl bg-[#eef3ef] p-4 text-sm leading-relaxed text-slate-700">{insight}</div>)}</div></Section>
          <Section><h2 className="text-2xl font-bold text-[#4d5b52]">Recurring Detection</h2><p className="mt-2 text-sm text-slate-500">זיהוי מנויים, ביטוחים, סלולר ושכירות לפי מילות מפתח וחזרה בין חודשים.</p><div className="mt-5 space-y-3">{recurringTransactions.length ? recurringTransactions.map((item) => <div key={item.id} className="flex justify-between rounded-2xl bg-slate-50 p-4 text-sm"><span>{item.merchant}</span><strong>{SHEKEL.format(item.amount)}</strong></div>) : <div className="rounded-2xl bg-slate-50 p-4 text-sm text-slate-400">לא זוהו עסקאות חוזרות עדיין</div>}</div></Section>
        </section>

        <Section>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-3xl font-bold">קרנות, פנסיה וחסכונות</h2><p className="mt-2 text-sm text-slate-500">הפרשות חודשיות לקרן השתלמות, פנסיה וחסכונות קבועים.</p></div><button onClick={addSavingsProduct} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת חיסכון</button></div>
          <div className="mt-7 grid gap-3">{monthData.savingsProducts.map((product) => <div key={product.id} className="grid gap-3 rounded-2xl border border-slate-100 p-4 md:grid-cols-[1fr_140px_120px_150px_150px_44px]"><input value={product.name} onChange={(event) => updateRow('savingsProducts', product.id, 'name', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm" /><select value={product.type} onChange={(event) => updateRow('savingsProducts', product.id, 'type', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm"><option>קרן השתלמות</option><option>פנסיה</option><option>קופת גמל</option><option>חיסכון</option><option>השקעות</option></select><input value={product.owner} onChange={(event) => updateRow('savingsProducts', product.id, 'owner', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm" /><input type="number" value={product.monthlyDeposit} onChange={(event) => updateRow('savingsProducts', product.id, 'monthlyDeposit', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm" /><input type="number" value={product.currentBalance} onChange={(event) => updateRow('savingsProducts', product.id, 'currentBalance', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm" /><button onClick={() => removeRow('savingsProducts', product.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button></div>)}</div>
        </Section>

        <Section>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-3xl font-bold">יעדי חיסכון</h2><p className="mt-2 text-sm text-slate-500">טיסה ליפן, חתונה, קרן חירום וכל יעד אחר.</p></div><button onClick={addSavingGoal} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white">+ הוספת יעד</button></div>
          <div className="mt-6 grid gap-5 md:grid-cols-3">{monthData.savingGoals.map((goal) => { const progress = toNumber(goal.targetAmount) ? Math.min(100, Math.round((toNumber(goal.currentAmount) / toNumber(goal.targetAmount)) * 100)) : 0; return <div key={goal.id} className="rounded-3xl bg-[#eef3ef] p-5"><input value={goal.name} onChange={(event) => updateRow('savingGoals', goal.id, 'name', event.target.value)} className="w-full rounded-xl border border-slate-200 px-4 py-3 font-bold" /><div className="mt-3 grid gap-3"><input type="number" value={goal.targetAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'targetAmount', event.target.value)} placeholder="יעד" className="rounded-xl border border-slate-200 px-4 py-3" /><input type="number" value={goal.currentAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'currentAmount', event.target.value)} placeholder="נצבר" className="rounded-xl border border-slate-200 px-4 py-3" /><input type="number" value={goal.monthlyDeposit} onChange={(event) => updateRow('savingGoals', goal.id, 'monthlyDeposit', event.target.value)} placeholder="הפקדה חודשית" className="rounded-xl border border-slate-200 px-4 py-3" /></div><div className="mt-4 flex justify-between text-sm font-bold"><span>{progress}%</span><button onClick={() => removeRow('savingGoals', goal.id)} className="text-red-500">מחיקה</button></div><div className="mt-2 h-3 overflow-hidden rounded-full bg-white"><div className="h-full rounded-full bg-[#7a9b76]" style={{ width: `${progress}%` }} /></div></div>; })}</div>
        </Section>

        <section className="grid gap-6 lg:grid-cols-2"><Section><h2 className="text-3xl font-bold">קרן חירום</h2><p className="mt-2 text-sm text-slate-500">מלאו סכום חיסכון נזיל נוכחי.</p><input type="number" value={monthData.emergencyFund} onChange={(event) => updateMonthField('emergencyFund', toNumber(event.target.value))} className="mt-6 w-full rounded-2xl border border-slate-200 px-5 py-4 text-xl font-bold outline-none focus:border-[#8fa88c]" /></Section><Section><h2 className="text-3xl font-bold">חיפוש ופילטרים</h2><div className="mt-5 grid gap-3"><input value={searchTerm} onChange={(event) => setSearchTerm(event.target.value)} placeholder="חיפוש בית עסק, למשל וולט" className="rounded-2xl border border-slate-200 px-4 py-3" /><select value={categoryFilter} onChange={(event) => setCategoryFilter(event.target.value)} className="rounded-2xl border border-slate-200 px-4 py-3"><option>הכול</option>{defaultExpenseCategories.map((category) => <option key={category}>{category}</option>)}</select><div className="grid gap-3 md:grid-cols-2"><input value={minAmount} onChange={(event) => setMinAmount(event.target.value)} type="number" placeholder="סכום מינימום" className="rounded-2xl border border-slate-200 px-4 py-3" /><input value={maxAmount} onChange={(event) => setMaxAmount(event.target.value)} type="number" placeholder="סכום מקסימום" className="rounded-2xl border border-slate-200 px-4 py-3" /></div></div></Section></section>

        <Section>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-3xl font-bold">הוצאות ידניות</h2><p className="mt-2 text-sm text-slate-500">הוצאות שלא נכנסות מכרטיסי האשראי. הטבלה רחבה ואפשר לגלול אופקית.</p></div><button onClick={addManualExpense} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת הוצאה</button></div>
          <div className="mt-7 overflow-x-auto rounded-3xl border border-slate-100 bg-white"><div className="min-w-[920px]"><div className="grid grid-cols-[minmax(320px,1fr)_180px_180px_60px] gap-3 bg-[#f3f5ef] px-5 py-4 text-sm font-bold text-[#4f6854]"><div>קטגוריה</div><div>סוג</div><div>סכום</div><div></div></div>{monthData.manualExpenses.map((expense) => <div key={expense.id} className="grid grid-cols-[minmax(320px,1fr)_180px_180px_60px] gap-3 border-t border-slate-100 p-4"><input value={expense.category} onChange={(event) => updateRow('manualExpenses', expense.id, 'category', event.target.value)} className="w-full rounded-xl border border-slate-200 px-4 py-4 text-base outline-none focus:border-[#8fa88c]" /><select value={expense.type} onChange={(event) => updateRow('manualExpenses', expense.id, 'type', event.target.value)} className="w-full rounded-xl border border-slate-200 px-4 py-4 text-base outline-none focus:border-[#8fa88c]"><option>קבועה</option><option>משתנה</option><option>חיסכון</option><option>חד פעמית</option></select><input type="number" value={expense.amount} onChange={(event) => updateRow('manualExpenses', expense.id, 'amount', event.target.value)} className="w-full rounded-xl border border-slate-200 px-4 py-4 text-base outline-none focus:border-[#8fa88c]" /><button onClick={() => removeRow('manualExpenses', expense.id)} className="rounded-xl bg-slate-100 text-lg font-bold text-slate-500">×</button></div>)}</div></div>
        </Section>

        <Section>
          <h2 className="text-3xl font-bold">כל עסקאות האשראי המסוננות</h2>
          <div className="mt-6 overflow-x-auto rounded-3xl border border-slate-100 bg-white"><div className="min-w-[1200px]"><div className="grid grid-cols-[160px_1fr_240px_180px_180px] bg-[#eef3ef] px-6 py-4 text-sm font-bold text-[#4d5b52]"><div>תאריך</div><div>בית עסק</div><div>קטגוריה לומדת</div><div>סכום</div><div>זיהוי</div></div>{filteredTransactions.map((transaction) => { const isRecurring = recurringTransactions.some((item) => item.id === transaction.id); return <div key={transaction.id} className="grid grid-cols-[160px_1fr_240px_180px_180px] gap-4 border-t border-slate-100 px-6 py-4"><div>{transaction.date}</div><div>{transaction.merchant}</div><select value={transaction.category} onChange={(event) => updateTransactionCategory(transaction.id, event.target.value)} className="rounded-xl border border-slate-200 px-3 py-2 text-sm">{defaultExpenseCategories.map((category) => <option key={category}>{category}</option>)}</select><div className="font-bold">{SHEKEL.format(transaction.amount)}</div><div>{isRecurring ? '🔁 חוזר' : '—'}</div></div>; })}{filteredTransactions.length === 0 ? <div className="p-16 text-center text-slate-400">אין עסקאות להצגה 🌿</div> : null}</div></div>
        </Section>
      </div>
    </div>
  );
}
