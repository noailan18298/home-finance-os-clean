'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

const STORAGE_KEY = 'family-finance-os-months-hebrew-v4';

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
  'קניות': 1200,
  'בריאות': 800,
  'אחר': 1000,
};

const merchantCategoryMap = {
  wolt: 'מסעדות / וולט',
  tenbis: 'מסעדות / וולט',
  shufersal: 'סופר / מזון',
  רמי: 'סופר / מזון',
  victory: 'סופר / מזון',
  yellow: 'תחבורה / דלק',
  דור: 'תחבורה / דלק',
  פז: 'תחבורה / דלק',
  fox: 'קניות',
  zara: 'קניות',
  superpharm: 'בריאות',
  כללית: 'בריאות',
};

function makeId(prefix = 'id') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function getCurrentMonthKey() {
  return new Date().toISOString().slice(0, 7);
}

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

  const cleaned = String(value || '')
    .replace(/[₪,]/g, '')
    .replace(/\s/g, '')
    .trim();

  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
}

function detectCategory(merchant = '') {
  const lower = String(merchant).toLowerCase();

  for (const key of Object.keys(merchantCategoryMap)) {
    if (lower.includes(key.toLowerCase())) {
      return merchantCategoryMap[key];
    }
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

function normalizeImportedRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];

  const firstRow = rows[0].join(' ').toLowerCase();

  const hasHeader =
    firstRow.includes('date') ||
    firstRow.includes('תאריך') ||
    firstRow.includes('amount') ||
    firstRow.includes('סכום') ||
    firstRow.includes('merchant') ||
    firstRow.includes('בית עסק');

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
        category: detectCategory(merchant),
      };
    })
    .filter((transaction) => transaction.amount > 0);
}

function parseCsvText(text) {
  const cleanText = String(text || '').replace(/\r/g, '');

  const lines = cleanText
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length === 0) {
    return [];
  }

  const rows = lines.map(splitCsvLine).filter((row) => row.length > 0);

  return normalizeImportedRows(rows);
}

function parseExcelArrayBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array' });

  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    return [];
  }

  const sheet = workbook.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: '',
  });

  return normalizeImportedRows(rows);
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

function buildRealInsights(transactions) {
  if (!transactions.length) {
    return [
      'עדיין אין נתונים אמיתיים. העלו קובץ CSV או Excel מאתר האשראי כדי לקבל תובנות.',
    ];
  }

  const total = transactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
  const categoryTotals = getCategoryTotals(transactions);
  const merchantTotals = getMerchantTotals(transactions);
  const sortedCategories = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]);
  const sortedMerchants = Object.entries(merchantTotals).sort((a, b) => b[1] - a[1]);
  const averageTransaction = total / transactions.length;
  const healthScore = calculateFinancialHealthScore(transactions);
  const insights = [];

  insights.push(`ציון בריאות הוצאות לחודש הזה: ${healthScore}/100. הציון מחושב לפי חריגות תקציב, ריכוזיות, עסקאות גדולות וסיווגים חסרים.`);

  const [topCategory, topCategoryAmount] = sortedCategories[0] || [];
  if (topCategory) {
    insights.push(
      `הקטגוריה הגדולה ביותר היא ${topCategory}: ${SHEKEL.format(topCategoryAmount)}, שהם ${Math.round((topCategoryAmount / total) * 100)}% מסך החיובים.`
    );
  }

  const [topMerchant, topMerchantAmount] = sortedMerchants[0] || [];
  if (topMerchant) {
    insights.push(
      `בית העסק עם הסכום הגבוה ביותר הוא ${topMerchant}: ${SHEKEL.format(topMerchantAmount)}, כלומר ${Math.round((topMerchantAmount / total) * 100)}% מהחיובים.`
    );
  }

  insights.push(`גובה עסקה ממוצעת בקובץ: ${SHEKEL.format(averageTransaction)}.`);

  Object.entries(categoryBudgets).forEach(([category, budget]) => {
    const spent = categoryTotals[category] || 0;
    if (spent > budget) {
      insights.push(
        `${category} חרגה מהתקציב: ${SHEKEL.format(spent)} מתוך ${SHEKEL.format(budget)}. חריגה של ${SHEKEL.format(spent - budget)}.`
      );
    } else if (spent >= budget * 0.8) {
      insights.push(
        `${category} מתקרבת לתקציב: ${SHEKEL.format(spent)} מתוך ${SHEKEL.format(budget)}.`
      );
    }
  });

  const uncategorized = categoryTotals['אחר'] || 0;
  if (uncategorized > 0) {
    insights.push(
      `${SHEKEL.format(uncategorized)} עדיין מסווגים כ״אחר״. מומלץ לעבור על העסקאות האלו כדי לשפר את הדיוק של הניתוח.`
    );
  }

  const largeTransactions = transactions
    .filter((transaction) => toNumber(transaction.amount) >= Math.max(500, total * 0.08))
    .sort((a, b) => toNumber(b.amount) - toNumber(a.amount));

  if (largeTransactions.length > 0) {
    insights.push(
      `זוהו ${largeTransactions.length} עסקאות גדולות יחסית. הגדולה ביותר: ${largeTransactions[0].merchant} בסך ${SHEKEL.format(largeTransactions[0].amount)}.`
    );
  }

  const repeatedMerchants = sortedMerchants.filter(([, amount]) => amount >= averageTransaction * 2).slice(0, 3);
  if (repeatedMerchants.length > 0) {
    insights.push(
      `בתי עסק שכדאי לבדוק לעומק: ${repeatedMerchants.map(([merchant, amount]) => `${merchant} (${SHEKEL.format(amount)})`).join(', ')}.`
    );
  }

  if (healthScore >= 85) {
    insights.push('לא נמצאו חריגות משמעותיות לפי החוקים שהוגדרו. נראה שהחודש מאוזן יחסית.');
  }

  return insights;
}

function runSmokeTests() {
  console.assert(toNumber('₪1,250') === 1250, 'currency parsing failed');
  console.assert(detectCategory('Wolt TLV') === 'מסעדות / וולט', 'wolt category failed');
  console.assert(splitCsvLine('a,b,c').length === 3, 'csv split failed');
  console.assert(
    parseCsvText('date,merchant,amount\n2026-01-01,Wolt,55').length === 1,
    'csv parse failed'
  );
  console.assert(
    normalizeImportedRows([
      ['תאריך', 'בית עסק', 'סכום'],
      ['01/01/2026', 'Yellow', '300'],
    ])[0].category === 'תחבורה / דלק',
    'excel normalize failed'
  );
  console.assert(getCategoryTotals([{ category: 'קניות', amount: 10 }, { category: 'קניות', amount: 20 }]).קניות === 30, 'category totals failed');
  console.assert(buildRealInsights([{ merchant: 'Wolt', category: 'מסעדות / וולט', amount: 900 }]).some((insight) => insight.includes('חרגה')), 'real budget insight failed');
  console.assert(calculateFinancialHealthScore([{ merchant: 'Wolt', category: 'מסעדות / וולט', amount: 900 }]) < 100, 'health score should react to budget exceptions');
}

if (typeof window !== 'undefined') {
  runSmokeTests();
}

export default function PersonalIsraeliFamilyFinanceDashboard() {
  const [transactions, setTransactions] = useState([]);
  const [fileName, setFileName] = useState('');
  const [aiInsights, setAiInsights] = useState([]);
  const [aiStatus, setAiStatus] = useState('idle');
  const [aiError, setAiError] = useState('');

  async function importFile(file) {
    const lower = file.name.toLowerCase();
    setFileName(file.name);
    setAiInsights([]);
    setAiError('');

    if (lower.endsWith('.csv')) {
      const text = await file.text();
      setTransactions(parseCsvText(text));
      return;
    }

    if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
      const buffer = await file.arrayBuffer();
      setTransactions(parseExcelArrayBuffer(buffer));
      return;
    }

    alert('נא להעלות CSV או Excel');
  }

  async function generateAiInsights() {
    if (!transactions.length) {
      setAiError('צריך להעלות קובץ לפני שמריצים AI.');
      return;
    }

    setAiStatus('loading');
    setAiError('');

    try {
      const response = await fetch('/api/ai-insights', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          transactions,
          categoryBudgets,
          totals: {
            total,
            transactionCount: transactions.length,
          },
        }),
      });

      if (!response.ok) {
        throw new Error('AI endpoint failed');
      }

      const data = await response.json();
      setAiInsights(Array.isArray(data.insights) ? data.insights : []);
      setAiStatus('ready');
    } catch (error) {
      setAiStatus('error');
      setAiError('ה־AI עדיין לא מחובר בשרת. בינתיים מוצגות התובנות המקומיות שמחושבות מהנתונים.');
    }
  }

  const total = useMemo(() => {
    return transactions.reduce((sum, item) => sum + toNumber(item.amount), 0);
  }, [transactions]);

  const categoryTotals = useMemo(() => getCategoryTotals(transactions), [transactions]);
  const merchantTotals = useMemo(() => getMerchantTotals(transactions), [transactions]);
  const realInsights = useMemo(() => buildRealInsights(transactions), [transactions]);
  const healthScore = useMemo(() => calculateFinancialHealthScore(transactions), [transactions]);
  const topMerchants = useMemo(() => {
    return Object.entries(merchantTotals)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
  }, [merchantTotals]);

  return (
    <div
      dir="rtl"
      className="min-h-screen bg-gradient-to-br from-[#f4f1ea] via-[#faf8f4] to-[#eef3ef] p-6 text-slate-800"
    >
      <div className="mx-auto max-w-7xl space-y-8">
        <div className="rounded-[32px] border border-[#d9d4c7] bg-white p-8 shadow-sm">
          <div className="flex flex-col gap-6 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <h1 className="text-5xl font-bold text-[#4d5b52]">
                מערכת פיננסית משפחתית 🇮🇱
              </h1>
              <p className="mt-4 text-lg text-slate-500">
                העלאת CSV ו־Excel אמיתיים מכרטיסי אשראי, עם תובנות שמחושבות רק מהנתונים שלכם
              </p>
            </div>

            <label className="cursor-pointer rounded-2xl bg-[#7a9b76] px-6 py-4 text-white shadow-sm transition hover:bg-[#6f8c6b]">
              העלאת קובץ
              <input
                type="file"
                className="hidden"
                accept=".csv,.xlsx,.xls"
                onChange={(event) => {
                  const file = event.target.files?.[0];
                  if (file) importFile(file);
                }}
              />
            </label>
          </div>

          {fileName ? (
            <div className="mt-6 rounded-2xl bg-[#eef3ef] px-5 py-4 text-sm text-[#4d5b52]">
              קובץ שנקלט: <strong>{fileName}</strong>
            </div>
          ) : null}
        </div>

        <div className="grid gap-6 md:grid-cols-3">
          <div className="rounded-3xl bg-white p-6 shadow-sm">
            <div className="text-sm text-slate-500">סה״כ עסקאות</div>
            <div className="mt-3 text-4xl font-bold">
              {transactions.length}
            </div>
          </div>

          <div className="rounded-3xl bg-white p-6 shadow-sm">
            <div className="text-sm text-slate-500">סה״כ חיובים</div>
            <div className="mt-3 text-4xl font-bold">
              {SHEKEL.format(total)}
            </div>
          </div>

          <div className="rounded-3xl bg-white p-6 shadow-sm">
            <div className="text-sm text-slate-500">קטגוריות שזוהו</div>
            <div className="mt-3 text-4xl font-bold">
              {Object.keys(categoryTotals).length}
            </div>
          </div>

          <div className="rounded-3xl bg-white p-6 shadow-sm md:col-span-3">
            <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
              <div>
                <div className="text-sm text-slate-500">ציון AI מקומי</div>
                <div className="mt-3 text-5xl font-bold text-[#4d5b52]">
                  {healthScore === null ? '—' : `${healthScore}/100`}
                </div>
              </div>
              <div className="max-w-2xl text-sm leading-relaxed text-slate-500">
                הציון לא מגיע מטקסט דמו. הוא מחושב מתוך הקובץ לפי חריגות תקציב, עסקאות גדולות, בתי עסק דומיננטיים וקטגוריות שלא סווגו.
              </div>
            </div>
            <div className="mt-5 h-3 overflow-hidden rounded-full bg-[#eef3ef]">
              <div className="h-full rounded-full bg-[#7a9b76]" style={{ width: `${healthScore === null ? 0 : healthScore}%` }} />
            </div>
          </div>
        </div>

        <div className="grid gap-6 lg:grid-cols-2">
          <div className="rounded-[32px] border border-slate-200 bg-white p-6 shadow-sm">
            <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
              <div>
                <h2 className="text-2xl font-bold text-[#4d5b52]">תובנות AI אמיתיות</h2>
                <p className="mt-2 text-sm text-slate-500">
                  קודם מוצגות תובנות מקומיות שמחושבות מהנתונים. אחרי חיבור OpenAI בשרת, הכפתור יפיק ניתוח AI מלא.
                </p>
              </div>
              <button
                onClick={generateAiInsights}
                disabled={aiStatus === 'loading' || transactions.length === 0}
                className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-[#6f8c6b] disabled:cursor-not-allowed disabled:opacity-50"
              >
                {aiStatus === 'loading' ? 'מנתח…' : 'נתח עם AI'}
              </button>
            </div>

            {aiError ? (
              <div className="mt-4 rounded-2xl bg-amber-50 p-4 text-sm text-amber-800">
                {aiError}
              </div>
            ) : null}

            <div className="mt-5 space-y-3">
              {(aiInsights.length > 0 ? aiInsights : realInsights).map((insight) => (
                <div key={insight} className="rounded-2xl bg-[#eef3ef] p-4 text-sm leading-relaxed text-slate-700">
                  {insight}
                </div>
              ))}
            </div>
          </div>

          <div className="rounded-[32px] border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-2xl font-bold text-[#4d5b52]">בתי עסק מובילים</h2>
            <p className="mt-2 text-sm text-slate-500">
              מחושב לפי הסכום המצטבר בקובץ שהועלה.
            </p>
            <div className="mt-5 space-y-3">
              {topMerchants.length > 0 ? topMerchants.map(([merchant, amount]) => (
                <div key={merchant} className="flex items-center justify-between rounded-2xl bg-slate-50 p-4 text-sm">
                  <span className="font-bold">{merchant}</span>
                  <span>{SHEKEL.format(amount)}</span>
                </div>
              )) : (
                <div className="rounded-2xl bg-slate-50 p-4 text-sm text-slate-400">אין עדיין נתונים</div>
              )}
            </div>
          </div>
        </div>

        <div className="rounded-[32px] border border-slate-200 bg-white p-6 shadow-sm">
          <h2 className="text-2xl font-bold text-[#4d5b52]">תקציב מול בפועל לפי קטגוריה</h2>
          <div className="mt-5 grid gap-4 md:grid-cols-2 xl:grid-cols-3">
            {Object.entries(categoryBudgets).map(([category, budget]) => {
              const spent = categoryTotals[category] || 0;
              const percent = budget ? Math.round((spent / budget) * 100) : 0;
              const status = percent < 80 ? '🟢' : percent < 100 ? '🟠' : '🔴';

              return (
                <div key={category} className="rounded-3xl bg-slate-50 p-5">
                  <div className="flex items-center justify-between">
                    <div className="font-bold">{category}</div>
                    <div>{status}</div>
                  </div>
                  <div className="mt-3 text-2xl font-bold">{SHEKEL.format(spent)}</div>
                  <div className="mt-1 text-sm text-slate-500">מתוך {SHEKEL.format(budget)}</div>
                  <div className="mt-4 h-3 overflow-hidden rounded-full bg-white">
                    <div className="h-full rounded-full bg-[#7a9b76]" style={{ width: `${Math.min(percent, 100)}%` }} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        <div className="overflow-x-auto rounded-[32px] border border-slate-200 bg-white shadow-sm">
          <div className="min-w-[1100px]">
            <div className="grid grid-cols-[160px_1fr_220px_180px] bg-[#eef3ef] px-6 py-4 text-sm font-bold text-[#4d5b52]">
              <div>תאריך</div>
              <div>בית עסק</div>
              <div>קטגוריה</div>
              <div>סכום</div>
            </div>

            {transactions.map((transaction) => (
              <div
                key={transaction.id}
                className="grid grid-cols-[160px_1fr_220px_180px] gap-4 border-t border-slate-100 px-6 py-4"
              >
                <div>{transaction.date}</div>
                <div>{transaction.merchant}</div>
                <div>{transaction.category}</div>
                <div className="font-bold">
                  {SHEKEL.format(transaction.amount)}
                </div>
              </div>
            ))}

            {transactions.length === 0 ? (
              <div className="p-16 text-center text-slate-400">
                עדיין לא נטענו עסקאות 🌿
              </div>
            ) : null}
          </div>
        </div>
      </div>
    </div>
  );
}
