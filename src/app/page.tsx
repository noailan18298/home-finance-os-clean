import { useEffect, useMemo, useState } from 'react';

const STORAGE_KEY = 'family-finance-os-months-hebrew-v2';

const SHEKEL = new Intl.NumberFormat('he-IL', {
  style: 'currency',
  currency: 'ILS',
  maximumFractionDigits: 0,
});

const COLORS = {
  page: 'bg-gradient-to-br from-[#f4f1ea] via-[#faf8f4] to-[#eef3ef]',
  card: 'bg-[#fcfbf8]',
  border: 'border-[#d9d4c7]',
  sage: 'bg-[#7a9b76]',
  sageDark: 'bg-[#4d5b52]',
  sageText: 'text-[#5f7d66]',
  soft: 'bg-[#f3f5ef]',
};

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
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function formatPercent(value) {
  return `${Math.round(value)}%`;
}

function detectCategory(merchant = '') {
  const lower = merchant.toLowerCase();

  for (const key of Object.keys(merchantCategoryMap)) {
    if (lower.includes(key.toLowerCase())) {
      return merchantCategoryMap[key];
    }
  }

  return 'אחר';
}

function createDefaultMonth() {
  return {
    dashboardTitle: 'מערכת ניהול פיננסי משפחתית 🇮🇱',
    emergencyFund: 0,
    lastSalaryImport: '',
    incomes: [
      { id: makeId('income'), name: 'משכורת 1', amount: 0 },
      { id: makeId('income'), name: 'משכורת 2', amount: 0 },
      { id: makeId('income'), name: 'קצבת ילדים / ביטוח לאומי', amount: 0 },
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
      { id: makeId('saving'), name: 'חיסכון כללי', type: 'חיסכון', owner: 'משפחה', monthlyDeposit: 0, currentBalance: 0 },
    ],
    savingGoals: [
      { id: makeId('goal'), name: 'טיול משפחתי', targetAmount: 30000, currentAmount: 0, monthlyDeposit: 0, targetDate: '2026-12' },
    ],
    creditCards: [
      { id: makeId('card'), name: 'כרטיס אשראי 1', owner: 'נועה', importedFile: '', transactions: [] },
      { id: makeId('card'), name: 'כרטיס אשראי 2', owner: 'אורן', importedFile: '', transactions: [] },
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

function normalizeMonthData(monthData) {
  const safe = monthData || {};

  return {
    emergencyFund: safe.emergencyFund || 0,
    dashboardTitle: safe.dashboardTitle || 'מערכת ניהול פיננסי משפחתית 🇮🇱',
    lastSalaryImport: safe.lastSalaryImport || '',
    incomes: safe.incomes || [],
    manualExpenses: safe.manualExpenses || safe.expenses || [],
    savingsProducts: safe.savingsProducts || [],
    savingGoals: safe.savingGoals || [],
    creditCards: safe.creditCards || [],
    selfEmployed: safe.selfEmployed || {
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

function runSmokeTests() {
  console.assert(toNumber('12') === 12, 'toNumber should parse numeric strings');
  console.assert(toNumber('not-a-number') === 0, 'toNumber should return zero for invalid values');
  console.assert(detectCategory('Wolt Tel Aviv') === 'מסעדות / וולט', 'Wolt should map to restaurants');
  console.assert(detectCategory('Unknown merchant') === 'אחר', 'Unknown merchants should map to other');
  console.assert(createDefaultMonth().incomes.length > 0, 'Default month should include income rows');
  console.assert(createDefaultMonth().dashboardTitle.length > 0, 'Default month should include a dashboard title');
}

if (typeof window !== 'undefined') {
  runSmokeTests();
}

function StatCard({ title, value, note }) {
  return (
    <div className="rounded-3xl border border-slate-100 bg-slate-50 p-6 shadow-sm">
      <div className="text-sm font-medium text-slate-500">{title}</div>
      <div className="mt-3 text-3xl font-bold text-slate-900">{value}</div>
      <div className="mt-2 text-sm text-emerald-700">{note}</div>
    </div>
  );
}

function Section({ children, className = '' }) {
  return (
    <section className={`rounded-[32px] border border-slate-200 bg-white p-7 shadow-sm ${className}`}>
      {children}
    </section>
  );
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

  useEffect(() => {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(months));
    } catch {
      // Local storage can fail in private/sandboxed browsers. The dashboard still works in memory.
    }
  }, [months]);

  const monthData = normalizeMonthData(months[selectedMonth]);

  const creditCardExpenses = useMemo(() => {
    const byCategory = {};

    monthData.creditCards.forEach((card) => {
      (card.transactions || []).forEach((transaction) => {
        const category = transaction.category || 'אחר';
        byCategory[category] = (byCategory[category] || 0) + toNumber(transaction.amount);
      });
    });

    return Object.entries(byCategory).map(([category, amount]) => ({
      id: `credit-${category}`,
      category,
      amount,
      type: 'אשראי',
      source: 'credit',
    }));
  }, [monthData.creditCards]);

  const allExpenses = useMemo(
    () => [...monthData.manualExpenses, ...creditCardExpenses],
    [monthData.manualExpenses, creditCardExpenses]
  );

  const totalIncome = useMemo(
    () => monthData.incomes.reduce((sum, item) => sum + toNumber(item.amount), 0),
    [monthData.incomes]
  );

  const totalCreditCards = useMemo(
    () => creditCardExpenses.reduce((sum, item) => sum + toNumber(item.amount), 0),
    [creditCardExpenses]
  );

  const totalManualExpenses = useMemo(
    () => monthData.manualExpenses.reduce((sum, item) => sum + toNumber(item.amount), 0),
    [monthData.manualExpenses]
  );

  const totalSavingsProducts = useMemo(
    () => monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0),
    [monthData.savingsProducts]
  );

  const totalSavingGoals = useMemo(
    () => monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.monthlyDeposit), 0),
    [monthData.savingGoals]
  );

  const categoryBudgets = {
    'סופר / מזון': 4000,
    'מסעדות / וולט': 800,
    קניות: 1200,
    'תחבורה / דלק': 1800,
  };

  const categoryTotals = allExpenses.reduce((acc, expense) => {
    acc[expense.category] = (acc[expense.category] || 0) + toNumber(expense.amount);
    return acc;
  }, {});

  const totalAssets =
    toNumber(monthData.emergencyFund) +
    monthData.savingsProducts.reduce((sum, item) => sum + toNumber(item.currentBalance), 0) +
    monthData.savingGoals.reduce((sum, item) => sum + toNumber(item.currentAmount), 0);

  const selfEmployedVatDue = Math.max(
    0,
    toNumber(monthData.selfEmployed.vatCollected) - toNumber(monthData.selfEmployed.vatPaidOnExpenses)
  );

  const totalSelfEmployedPayments =
    selfEmployedVatDue +
    toNumber(monthData.selfEmployed.incomeTaxAdvance) +
    toNumber(monthData.selfEmployed.nationalInsurance) +
    toNumber(monthData.selfEmployed.businessExpenses);

  const totalPlannedSavings = totalSavingsProducts + totalSavingGoals;
  const totalExpenses = totalManualExpenses + totalCreditCards + totalPlannedSavings + totalSelfEmployedPayments;
  const monthlySavings = totalIncome - totalExpenses;
  const savingsRate = totalIncome ? (monthlySavings / totalIncome) * 100 : 0;
  const emergencyMonths = toNumber(monthData.emergencyFund) / (totalExpenses || 1);

  const insightCards = useMemo(() => {
    const groceries = allExpenses.find((item) => item.category.includes('סופר') || item.category.includes('מזון'));
    const restaurants = allExpenses.find((item) => item.category.includes('וולט') || item.category.includes('מסעדות'));
    const fixedTotal = allExpenses
      .filter((item) => item.type === 'קבועה')
      .reduce((sum, item) => sum + toNumber(item.amount), 0);

    const insights = [];

    if (monthlySavings > 0) {
      insights.push(`החודש אתם צפויים לחסוך ${SHEKEL.format(monthlySavings)}. זה המקום שבו האקסל מחייך בשקט.`);
    } else if (totalIncome > 0) {
      insights.push(`החודש במינוס של ${SHEKEL.format(Math.abs(monthlySavings))}. כדאי להתחיל מבדיקת הוצאות משתנות ואשראי.`);
    } else {
      insights.push('התחילו בהזנת הכנסות, הוצאות וכרטיסי אשראי. אחרי זה הדשבורד יתחיל להפיק תובנות.');
    }

    if (totalCreditCards > 0) {
      insights.push(`סה״כ חיובי אשראי החודש: ${SHEKEL.format(totalCreditCards)}. הסכום הזה מוזן אוטומטית להוצאות לפי קטגוריות.`);
    }

    if (totalPlannedSavings > 0) {
      insights.push(`החודש אתם מפרישים ${SHEKEL.format(totalPlannedSavings)} לקרנות, פנסיה ויעדי חיסכון. זה כסף עם שליחות, לא כסף שנעלם.`);
    }

    if (totalSelfEmployedPayments > 0) {
      insights.push(`תשלומי העצמאי של ${monthData.selfEmployed.owner || 'אורן'} החודש: ${SHEKEL.format(totalSelfEmployedPayments)}. כדאי להפריד אותם מהוצאות הבית כדי לא לטעות בתחושת הכסף הפנוי.`);
    }

    if (restaurants && toNumber(restaurants.amount) > 0) {
      insights.push(`מסעדות / וולט בקצב שנתי: ${SHEKEL.format(toNumber(restaurants.amount) * 12)}. שווה להגדיר תקרה חודשית.`);
    }

    if (groceries && toNumber(groceries.amount) > 0 && totalIncome > 0) {
      insights.push(`סופר / מזון הם ${formatPercent((toNumber(groceries.amount) / totalIncome) * 100)} מההכנסה החודשית.`);
    }

    if (fixedTotal > 0 && totalIncome > 0) {
      insights.push(`הוצאות קבועות הן ${formatPercent((fixedTotal / totalIncome) * 100)} מההכנסה.`);
    }

    if (emergencyMonths > 0) {
      insights.push(`קרן החירום מכסה ${emergencyMonths.toFixed(1)} חודשים לפי קצב ההוצאות הנוכחי.`);
    }

    return insights.slice(0, 8);
  }, [allExpenses, monthlySavings, totalIncome, totalCreditCards, totalPlannedSavings, totalSelfEmployedPayments, monthData.selfEmployed.owner, emergencyMonths]);

  function setSelectedMonthData(nextData) {
    setMonths((current) => ({ ...current, [selectedMonth]: nextData }));
  }

  function ensureMonthExists() {
    if (months[selectedMonth]) return;
    setMonths((current) => ({ ...current, [selectedMonth]: createDefaultMonth() }));
  }

  function duplicatePreviousMonth() {
    const keys = Object.keys(months).sort();
    const previousKey = keys.filter((key) => key < selectedMonth).pop();
    if (!previousKey || months[selectedMonth]) return;

    const previous = normalizeMonthData(months[previousKey]);
    setMonths((current) => ({
      ...current,
      [selectedMonth]: {
        ...previous,
        incomes: previous.incomes.map((item) => ({ ...item, id: makeId('income') })),
        manualExpenses: previous.manualExpenses.map((item) => ({ ...item, id: makeId('expense') })),
        savingsProducts: previous.savingsProducts.map((item) => ({ ...item, id: makeId('saving') })),
        savingGoals: previous.savingGoals.map((item) => ({ ...item, id: makeId('goal') })),
        creditCards: previous.creditCards.map((card) => ({ ...card, id: makeId('card'), importedFile: '', transactions: [] })),
      },
    }));
  }

  function updateMonthField(field, value) {
    setSelectedMonthData({ ...monthData, [field]: value });
  }

  function updateSelfEmployedField(field, value) {
    const numericFields = ['grossRevenue', 'vatCollected', 'vatPaidOnExpenses', 'incomeTaxAdvance', 'nationalInsurance', 'businessExpenses'];
    setSelectedMonthData({
      ...monthData,
      selfEmployed: {
        ...monthData.selfEmployed,
        [field]: numericFields.includes(field) ? toNumber(value) : value,
      },
    });
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
    setSelectedMonthData({
      ...monthData,
      incomes: [...monthData.incomes, { id: makeId('income'), name: 'הכנסה חדשה', amount: 0 }],
    });
  }

  function addManualExpense() {
    setSelectedMonthData({
      ...monthData,
      manualExpenses: [...monthData.manualExpenses, { id: makeId('expense'), category: 'הוצאה חדשה', amount: 0, type: 'משתנה' }],
    });
  }

  function addSavingsProduct() {
    setSelectedMonthData({
      ...monthData,
      savingsProducts: [
        ...monthData.savingsProducts,
        { id: makeId('saving'), name: 'חיסכון חדש', type: 'חיסכון', owner: 'משפחה', monthlyDeposit: 0, currentBalance: 0 },
      ],
    });
  }

  function addSavingGoal() {
    setSelectedMonthData({
      ...monthData,
      savingGoals: [
        ...monthData.savingGoals,
        { id: makeId('goal'), name: 'יעד חדש', targetAmount: 0, currentAmount: 0, monthlyDeposit: 0, targetDate: selectedMonth },
      ],
    });
  }

  function addCreditCard() {
    setSelectedMonthData({
      ...monthData,
      creditCards: [...monthData.creditCards, { id: makeId('card'), name: 'כרטיס אשראי חדש', owner: '', importedFile: '', transactions: [] }],
    });
  }

  function updateCreditCard(cardId, field, value) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) => (card.id === cardId ? { ...card, [field]: value } : card)),
    });
  }

  function removeCreditCard(cardId) {
    setSelectedMonthData({ ...monthData, creditCards: monthData.creditCards.filter((card) => card.id !== cardId) });
  }

  function importMockSalarySlip(fileName) {
    const extractedSalary = {
      employee: 'נועה',
      netSalary: 21450,
      pensionDeposit: 2650,
      studyFundDeposit: 1570,
      bonus: 1200,
    };

    const updatedIncomes = [
      ...monthData.incomes,
      { id: makeId('income'), name: `שכר נטו - ${extractedSalary.employee}`, amount: extractedSalary.netSalary },
    ];

    if (extractedSalary.bonus > 0) {
      updatedIncomes.push({ id: makeId('income'), name: `בונוס - ${extractedSalary.employee}`, amount: extractedSalary.bonus });
    }

    const updatedSavings = [
      ...monthData.savingsProducts,
      {
        id: makeId('saving'),
        name: `קרן השתלמות ${extractedSalary.employee}`,
        type: 'קרן השתלמות',
        owner: extractedSalary.employee,
        monthlyDeposit: extractedSalary.studyFundDeposit,
        currentBalance: 0,
      },
      {
        id: makeId('saving'),
        name: `פנסיה ${extractedSalary.employee}`,
        type: 'פנסיה',
        owner: extractedSalary.employee,
        monthlyDeposit: extractedSalary.pensionDeposit,
        currentBalance: 0,
      },
    ];

    setSelectedMonthData({
      ...monthData,
      lastSalaryImport: fileName,
      incomes: updatedIncomes,
      savingsProducts: updatedSavings,
    });
  }

  function parseCsvText(text) {
    const lines = text
      .split(/
?
/)
      .map((line) => line.trim())
      .filter(Boolean);

    if (lines.length === 0) return [];

    const rows = lines.map((line) => {
      const separator = line.includes(';') ? ';' : ',';
      return line.split(separator).map((cell) => cell.trim().replace(/^"|"$/g, ''));
    });

    const firstRow = rows[0].join(' ').toLowerCase();
    const hasHeader = firstRow.includes('date') || firstRow.includes('תאריך') || firstRow.includes('amount') || firstRow.includes('סכום');
    const dataRows = hasHeader ? rows.slice(1) : rows;

    return dataRows
      .map((row) => {
        const date = row[0] || '';
        const merchant = row[1] || row[0] || 'עסקה ללא שם';
        const rawAmount = row[2] || row[1] || '0';
        const amount = toNumber(String(rawAmount).replace(',', '').replace('₪', ''));

        return {
          id: makeId('import'),
          date,
          merchant,
          amount: Math.abs(amount),
          category: detectCategory(merchant),
        };
      })
      .filter((transaction) => transaction.amount > 0);
  }

  function importCreditFile(cardId, file) {
    const fileName = file.name || '';
    const lower = fileName.toLowerCase();

    if (!lower.endsWith('.csv')) {
      alert('כרגע ההעלאה המדויקת תומכת בקבצי CSV. אפשר לייצא CSV מאתר האשראי ולהעלות כאן.');
      return;
    }

    const reader = new FileReader();

    reader.onload = () => {
      const text = String(reader.result || '');
      const importedTransactions = parseCsvText(text);

      setSelectedMonthData({
        ...monthData,
        creditCards: monthData.creditCards.map((card) =>
          card.id === cardId
            ? {
                ...card,
                importedFile: fileName,
                pendingTransactions: importedTransactions,
              }
            : card
        ),
      });
    };

    reader.readAsText(file, 'utf-8');
  }

  function updatePendingTransaction(cardId, transactionId, field, value) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId
          ? {
              ...card,
              pendingTransactions: (card.pendingTransactions || []).map((transaction) =>
                transaction.id === transactionId
                  ? { ...transaction, [field]: field === 'amount' ? toNumber(value) : value }
                  : transaction
              ),
            }
          : card
      ),
    });
  }

  function removePendingTransaction(cardId, transactionId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId
          ? {
              ...card,
              pendingTransactions: (card.pendingTransactions || []).filter((transaction) => transaction.id !== transactionId),
            }
          : card
      ),
    });
  }

  function approvePendingTransactions(cardId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId
          ? {
              ...card,
              transactions: [...(card.transactions || []), ...(card.pendingTransactions || [])],
              pendingTransactions: [],
            }
          : card
      ),
    });
  }

  function addTransaction(cardId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId
          ? {
              ...card,
              transactions: [...(card.transactions || []), { id: makeId('tx'), merchant: 'עסקה חדשה', category: 'אחר', amount: 0 }],
            }
          : card
      ),
    });
  }

  function updateTransaction(cardId, transactionId, field, value) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId
          ? {
              ...card,
              transactions: (card.transactions || []).map((transaction) =>
                transaction.id === transactionId
                  ? { ...transaction, [field]: field === 'amount' ? toNumber(value) : value }
                  : transaction
              ),
            }
          : card
      ),
    });
  }

  function removeTransaction(cardId, transactionId) {
    setSelectedMonthData({
      ...monthData,
      creditCards: monthData.creditCards.map((card) =>
        card.id === cardId
          ? { ...card, transactions: (card.transactions || []).filter((transaction) => transaction.id !== transactionId) }
          : card
      ),
    });
  }

  if (!months[selectedMonth]) {
    return (
      <div dir="rtl" className={`min-h-screen ${COLORS.page} p-6 text-right text-slate-800`}>
        <div className="mx-auto max-w-3xl rounded-[32px] border border-[#d9d4c7] bg-[#fcfbf8] p-8 shadow-sm">
          <h1 className="text-4xl font-bold">חודש חדש</h1>
          <p className="mt-3 text-slate-600">בחרת חודש שעדיין לא קיים בדשבורד.</p>
          <input
            type="month"
            value={selectedMonth}
            onChange={(event) => setSelectedMonth(event.target.value)}
            className="mt-6 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-lg font-bold"
          />
          <div className="mt-6 flex gap-3">
            <button onClick={ensureMonthExists} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white">
              התחלה ריקה
            </button>
            <button onClick={duplicatePreviousMonth} className="rounded-2xl bg-[#4d5b52] px-5 py-3 text-sm font-semibold text-white">
              העתקה מחודש קודם
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div dir="rtl" className={`min-h-screen ${COLORS.page} p-6 text-right text-slate-800`}>
      <div className="mx-auto max-w-7xl space-y-8">
        <section className="overflow-hidden rounded-[32px] border border-[#d9d4c7] bg-[#fcfbf8] shadow-[0_10px_40px_rgba(0,0,0,0.05)]">
          <div className="bg-gradient-to-l from-[#6b8f71] to-[#9ebc8a] p-8 text-[#f8f6f1]">
            <div className="flex flex-col gap-6 lg:flex-row lg:items-center lg:justify-between">
              <div>
                <input
                  value={monthData.dashboardTitle}
                  onChange={(event) => updateMonthField('dashboardTitle', event.target.value)}
                  className="w-full max-w-3xl rounded-2xl border border-white/20 bg-white/10 px-4 py-3 text-4xl font-bold tracking-tight text-white outline-none transition placeholder:text-white/70 focus:bg-white/20 md:text-5xl"
                  placeholder="שם הדשבורד המשפחתי"
                />
                <p className="mt-4 max-w-2xl text-base text-[#f3f0e8]">
                  ממלאים הכנסות, הוצאות וכרטיסי אשראי בכל חודש. הדשבורד מחשב אוטומטית חיסכון, קצב שריפה, כיסוי חירום ותובנות.
                </p>
              </div>
              <div className="rounded-3xl bg-[#f7f5ef]/20 p-6 backdrop-blur-sm">
                <label className="text-sm uppercase tracking-widest text-[#f3f0e8]">חודש</label>
                <input
                  type="month"
                  value={selectedMonth}
                  onChange={(event) => setSelectedMonth(event.target.value)}
                  className="mt-2 w-full rounded-2xl border-0 bg-white px-4 py-3 text-lg font-bold text-slate-900 shadow-sm"
                />
              </div>
            </div>
          </div>

          <div className="grid gap-5 p-6 md:grid-cols-2 xl:grid-cols-7">
            <StatCard title="סה״כ הכנסות" value={SHEKEL.format(totalIncome)} note="כל מקורות ההכנסה" />
            <StatCard title="סה״כ הוצאות" value={SHEKEL.format(totalExpenses)} note={`${totalIncome ? formatPercent((totalExpenses / totalIncome) * 100) : '0%'} מההכנסה`} />
            <StatCard title="סה״כ אשראי" value={SHEKEL.format(totalCreditCards)} note="מהכרטיסים" />
            <StatCard title="עצמאי" value={SHEKEL.format(totalSelfEmployedPayments)} note="מע״מ, מס וביטוח לאומי" />
            <StatCard title="הפרשות וחסכונות" value={SHEKEL.format(totalPlannedSavings)} note="קרנות, פנסיה ויעדים" />
            <StatCard title="יתרה אחרי הכול" value={SHEKEL.format(monthlySavings)} note={`${formatPercent(savingsRate)} שיעור חיסכון`} />
            <StatCard title="קרן חירום" value={SHEKEL.format(toNumber(monthData.emergencyFund))} note={`${emergencyMonths.toFixed(1)} חודשי כיסוי`} />
          </div>
        </section>

        <section className="grid gap-6 lg:grid-cols-2">
          <Section>
            <div className="flex items-center justify-between gap-4">
              <div>
                <h2 className="text-3xl font-bold">הכנסות</h2>
                <p className="mt-2 text-sm text-slate-500">הוסיפו משכורות, בונוסים, פרילנס או כל הכנסה אחרת.</p>
              </div>
              <div className="flex gap-3">
                <label className="cursor-pointer rounded-2xl bg-[#7c9780] px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-[#6d876f]">
                  העלאת תלוש
                  <input
                    type="file"
                    accept="application/pdf"
                    className="hidden"
                    onChange={(event) => {
                      const file = event.target.files?.[0];
                      if (file) importMockSalarySlip(file.name);
                    }}
                  />
                </label>
                <button onClick={addIncome} className="rounded-2xl bg-[#7a9b76] px-4 py-3 text-sm font-semibold text-white shadow-sm">
                  + הוספה
                </button>
              </div>
            </div>

            {monthData.lastSalaryImport ? (
              <div className="mt-4 rounded-2xl bg-[#f3f5ef] p-4 text-sm text-[#4f6854]">תלוש אחרון שנקלט: {monthData.lastSalaryImport}</div>
            ) : null}

            <div className="mt-6 space-y-3">
              {monthData.incomes.map((income) => (
                <div key={income.id} className="grid gap-3 rounded-2xl border border-slate-100 p-4 md:grid-cols-[1fr_160px_44px]">
                  <input value={income.name} onChange={(event) => updateRow('incomes', income.id, 'name', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                  <input type="number" value={income.amount} onChange={(event) => updateRow('incomes', income.id, 'amount', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                  <button onClick={() => removeRow('incomes', income.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button>
                </div>
              ))}
            </div>
          </Section>

          <Section>
            <h2 className="text-3xl font-bold">עצמאי: מע״מ, מס הכנסה וביטוח לאומי</h2>
            <p className="mt-2 text-sm text-slate-500">אזור לאורן כעצמאי. הסכומים האלה נספרים בנפרד כדי שלא יתערבבו עם הוצאות הבית.</p>

            <div className="mt-6 grid gap-4 md:grid-cols-2">
              <label className="text-sm font-bold text-slate-600">
                בעל העסק
                <input value={monthData.selfEmployed.owner} onChange={(event) => updateSelfEmployedField('owner', event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
              </label>
              <label className="text-sm font-bold text-slate-600">
                הכנסה עסקית ברוטו
                <input type="number" value={monthData.selfEmployed.grossRevenue} onChange={(event) => updateSelfEmployedField('grossRevenue', event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
              </label>
              <label className="text-sm font-bold text-slate-600">
                מע״מ שנגבה מלקוחות
                <input type="number" value={monthData.selfEmployed.vatCollected} onChange={(event) => updateSelfEmployedField('vatCollected', event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
              </label>
              <label className="text-sm font-bold text-slate-600">
                מע״מ על הוצאות מוכרות
                <input type="number" value={monthData.selfEmployed.vatPaidOnExpenses} onChange={(event) => updateSelfEmployedField('vatPaidOnExpenses', event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
              </label>
              <label className="text-sm font-bold text-slate-600">
                מקדמת מס הכנסה
                <input type="number" value={monthData.selfEmployed.incomeTaxAdvance} onChange={(event) => updateSelfEmployedField('incomeTaxAdvance', event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
              </label>
              <label className="text-sm font-bold text-slate-600">
                ביטוח לאומי
                <input type="number" value={monthData.selfEmployed.nationalInsurance} onChange={(event) => updateSelfEmployedField('nationalInsurance', event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
              </label>
              <label className="text-sm font-bold text-slate-600 md:col-span-2">
                הוצאות עסקיות ששולמו החודש
                <input type="number" value={monthData.selfEmployed.businessExpenses} onChange={(event) => updateSelfEmployedField('businessExpenses', event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
              </label>
            </div>

            <div className="mt-6 grid gap-4 md:grid-cols-3">
              <div className="rounded-2xl bg-[#f3f5ef] p-5">
                <div className="text-sm text-slate-500">מע״מ צפוי לתשלום</div>
                <div className="mt-2 text-2xl font-bold text-[#5f7d66]">{SHEKEL.format(selfEmployedVatDue)}</div>
              </div>
              <div className="rounded-2xl bg-[#f3f5ef] p-5">
                <div className="text-sm text-slate-500">מס + ביטוח לאומי</div>
                <div className="mt-2 text-2xl font-bold text-[#5f7d66]">{SHEKEL.format(toNumber(monthData.selfEmployed.incomeTaxAdvance) + toNumber(monthData.selfEmployed.nationalInsurance))}</div>
              </div>
              <div className="rounded-2xl bg-[#4d5b52] p-5 text-white">
                <div className="text-sm text-slate-200">סה״כ תשלומי עצמאי</div>
                <div className="mt-2 text-2xl font-bold">{SHEKEL.format(totalSelfEmployedPayments)}</div>
              </div>
            </div>
          </Section>

          <Section>
            <h2 className="text-3xl font-bold">קרן חירום</h2>
            <p className="mt-2 text-sm text-slate-500">מלאו את סכום החיסכון הנזיל הנוכחי פעם בחודש.</p>
            <input type="number" value={monthData.emergencyFund} onChange={(event) => updateMonthField('emergencyFund', toNumber(event.target.value))} className="mt-6 w-full rounded-2xl border border-slate-200 px-5 py-4 text-xl font-bold outline-none focus:border-[#8fa88c]" />
            <div className="mt-6 rounded-3xl bg-[#4d5b52] p-6 text-white">
              <div className="text-sm text-slate-200">כיסוי נוכחי</div>
              <div className="mt-2 text-5xl font-bold">{emergencyMonths.toFixed(1)}</div>
              <div className="mt-2 text-sm text-slate-200">חודשים לפי ההוצאות של החודש</div>
            </div>
          </Section>
        </section>

        <Section>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-3xl font-bold">קרנות, פנסיה וחסכונות</h2>
              <p className="mt-2 text-sm text-slate-500">הפרשות חודשיות לקרן השתלמות, פנסיה וחסכונות קבועים.</p>
            </div>
            <button onClick={addSavingsProduct} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת חיסכון</button>
          </div>

          <div className="mt-7 overflow-x-auto rounded-2xl border border-slate-100">
            <div className="min-w-[900px]">
              <div className="grid grid-cols-[1fr_150px_120px_150px_150px_44px] bg-slate-50 px-4 py-3 text-xs font-bold text-slate-500">
                <div>שם</div><div>סוג</div><div>בעלים</div><div>הפקדה חודשית</div><div>יתרה נוכחית</div><div></div>
              </div>
              {monthData.savingsProducts.map((product) => (
                <div key={product.id} className="grid grid-cols-[1fr_150px_120px_150px_150px_44px] gap-3 border-t border-slate-100 p-3">
                  <input value={product.name} onChange={(event) => updateRow('savingsProducts', product.id, 'name', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                  <select value={product.type} onChange={(event) => updateRow('savingsProducts', product.id, 'type', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm outline-none focus:border-[#8fa88c]"><option>קרן השתלמות</option><option>פנסיה</option><option>קופת גמל</option><option>חיסכון</option><option>השקעות</option></select>
                  <input value={product.owner} onChange={(event) => updateRow('savingsProducts', product.id, 'owner', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                  <input type="number" value={product.monthlyDeposit} onChange={(event) => updateRow('savingsProducts', product.id, 'monthlyDeposit', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                  <input type="number" value={product.currentBalance} onChange={(event) => updateRow('savingsProducts', product.id, 'currentBalance', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                  <button onClick={() => removeRow('savingsProducts', product.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button>
                </div>
              ))}
            </div>
          </div>
        </Section>

        <Section>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-3xl font-bold">יעדי חיסכון</h2>
              <p className="mt-2 text-sm text-slate-500">חיסכון למטרה ספציפית כמו טיול, רכב או שיפוץ.</p>
            </div>
            <button onClick={addSavingGoal} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת יעד</button>
          </div>

          <div className="mt-7 grid gap-5 md:grid-cols-2 xl:grid-cols-3">
            {monthData.savingGoals.map((goal) => {
              const progress = toNumber(goal.targetAmount) ? Math.min(100, Math.round((toNumber(goal.currentAmount) / toNumber(goal.targetAmount)) * 100)) : 0;
              const remaining = Math.max(0, toNumber(goal.targetAmount) - toNumber(goal.currentAmount));
              return (
                <div key={goal.id} className="rounded-3xl border border-[#d7e2d5] bg-[#f1f5f0]/70 p-5">
                  <div className="flex items-center justify-between gap-3">
                    <input value={goal.name} onChange={(event) => updateRow('savingGoals', goal.id, 'name', event.target.value)} className="w-full rounded-xl border border-[#d7e2d5] bg-white px-4 py-3 text-sm font-bold outline-none focus:border-[#8aa28d]" />
                    <button onClick={() => removeRow('savingGoals', goal.id)} className="rounded-xl bg-white px-3 py-3 text-sm font-bold text-slate-500">×</button>
                  </div>
                  <div className="mt-4 grid gap-3">
                    <label className="text-xs font-bold text-slate-500">סכום יעד</label>
                    <input type="number" value={goal.targetAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'targetAmount', event.target.value)} className="rounded-xl border border-[#d7e2d5] bg-white px-4 py-3 text-sm outline-none focus:border-[#8aa28d]" />
                    <label className="text-xs font-bold text-slate-500">נצבר עד עכשיו</label>
                    <input type="number" value={goal.currentAmount} onChange={(event) => updateRow('savingGoals', goal.id, 'currentAmount', event.target.value)} className="rounded-xl border border-[#d7e2d5] bg-white px-4 py-3 text-sm outline-none focus:border-[#8aa28d]" />
                    <label className="text-xs font-bold text-slate-500">הפקדה חודשית</label>
                    <input type="number" value={goal.monthlyDeposit} onChange={(event) => updateRow('savingGoals', goal.id, 'monthlyDeposit', event.target.value)} className="rounded-xl border border-[#d7e2d5] bg-white px-4 py-3 text-sm outline-none focus:border-[#8aa28d]" />
                    <label className="text-xs font-bold text-slate-500">תאריך יעד</label>
                    <input type="month" value={goal.targetDate} onChange={(event) => updateRow('savingGoals', goal.id, 'targetDate', event.target.value)} className="rounded-xl border border-[#d7e2d5] bg-white px-4 py-3 text-sm outline-none focus:border-[#8aa28d]" />
                  </div>
                  <div className="mt-5">
                    <div className="flex justify-between text-sm font-bold text-emerald-900"><span>{progress}% הושלם</span><span>נשאר {SHEKEL.format(remaining)}</span></div>
                    <div className="mt-2 h-3 overflow-hidden rounded-full bg-white"><div className="h-full rounded-full bg-[#7c9780]" style={{ width: `${progress}%` }} /></div>
                  </div>
                </div>
              );
            })}
          </div>
        </Section>

        <Section>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-3xl font-bold">סיכום כרטיסי אשראי</h2>
              <p className="mt-2 text-sm text-slate-500">הוסיפו עסקאות לפי כרטיס וקטגוריה, או העלו CSV/Excel מאתר האשראי.</p>
            </div>
            <button onClick={addCreditCard} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת כרטיס</button>
          </div>

          <div className="mt-7 grid gap-6 xl:grid-cols-2">
            {monthData.creditCards.map((card) => {
              const cardTotal = (card.transactions || []).reduce((sum, item) => sum + toNumber(item.amount), 0);
              return (
                <div key={card.id} className="rounded-3xl border border-[#d8e2d2] bg-[#f3f5ef]/60 p-5">
                  <div className="grid gap-3 md:grid-cols-[1fr_1fr_110px_44px]">
                    <input value={card.name} onChange={(event) => updateCreditCard(card.id, 'name', event.target.value)} className="rounded-xl border border-[#d8e2d2] bg-white px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" placeholder="שם הכרטיס" />
                    <input value={card.owner} onChange={(event) => updateCreditCard(card.id, 'owner', event.target.value)} className="rounded-xl border border-[#d8e2d2] bg-white px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" placeholder="בעל/ת הכרטיס" />
                    <div className="rounded-xl bg-white px-4 py-3 text-sm font-bold text-[#5d755f]">{SHEKEL.format(cardTotal)}</div>
                    <button onClick={() => removeCreditCard(card.id)} className="rounded-xl bg-white text-sm font-bold text-slate-500">×</button>
                  </div>

                  <div className="mt-5 rounded-2xl border border-dashed border-[#d8e2d2] bg-white p-4">
                    <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                      <div>
                        <div className="text-sm font-bold text-[#4f6854]">העלאת CSV של פירוט אשראי</div>
                        <div className="mt-1 text-xs text-slate-500">ייצאו CSV מאתר האשראי, העלו כאן, בדקו את הקטגוריות ואז אשרו הכנסת עסקאות.</div>
                      </div>
                      <label className="cursor-pointer rounded-2xl bg-[#7a9b76] px-4 py-3 text-sm font-semibold text-white shadow-sm transition hover:bg-[#688864]">
                        העלאת CSV
                        <input type="file" accept=".csv,text/csv" className="hidden" onChange={(event) => { const file = event.target.files?.[0]; if (file) importCreditFile(card.id, file); }} />
                      </label>
                    </div>
                    {card.importedFile ? <div className="mt-3 rounded-xl bg-[#f3f5ef] px-4 py-3 text-sm text-[#4f6854]">נקלט קובץ: <strong>{card.importedFile}</strong></div> : null}
                    {(card.pendingTransactions || []).length > 0 ? (
                      <div className="mt-4 rounded-2xl border border-[#d8e2d2] bg-[#fcfbf8] p-4">
                        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                          <div>
                            <div className="text-sm font-bold text-[#4f6854]">עסקאות שזוהו לאישור</div>
                            <div className="mt-1 text-xs text-slate-500">בדקו סכומים וקטגוריות לפני שהן נכנסות להוצאות.</div>
                          </div>
                          <button onClick={() => approvePendingTransactions(card.id)} className="rounded-2xl bg-[#7a9b76] px-4 py-3 text-sm font-semibold text-white shadow-sm">
                            אשר והכנס להוצאות
                          </button>
                        </div>

                        <div className="mt-4 overflow-x-auto rounded-2xl border border-[#d8e2d2] bg-white">
                          <div className="min-w-[760px]">
                            <div className="grid grid-cols-[120px_1fr_170px_140px_44px] bg-[#e5eee2]/60 px-4 py-3 text-xs font-bold text-[#4f6854]">
                              <div>תאריך</div><div>בית עסק</div><div>קטגוריה מוצעת</div><div>סכום</div><div></div>
                            </div>
                            {(card.pendingTransactions || []).map((transaction) => (
                              <div key={transaction.id} className="grid grid-cols-[120px_1fr_170px_140px_44px] gap-3 border-t border-[#d8e2d2] p-3">
                                <input value={transaction.date || ''} onChange={(event) => updatePendingTransaction(card.id, transaction.id, 'date', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                                <input value={transaction.merchant} onChange={(event) => updatePendingTransaction(card.id, transaction.id, 'merchant', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                                <select value={transaction.category} onChange={(event) => updatePendingTransaction(card.id, transaction.id, 'category', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm outline-none focus:border-[#8fa88c]">{defaultExpenseCategories.map((category) => <option key={category}>{category}</option>)}</select>
                                <input type="number" value={transaction.amount} onChange={(event) => updatePendingTransaction(card.id, transaction.id, 'amount', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                                <button onClick={() => removePendingTransaction(card.id, transaction.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    ) : null}
                  </div>

                  <div className="mt-5 overflow-x-auto rounded-2xl border border-[#d8e2d2] bg-white">
                    <div className="min-w-[640px]">
                      <div className="grid grid-cols-[1fr_150px_140px_44px] bg-[#e5eee2]/60 px-4 py-3 text-xs font-bold text-[#4f6854]"><div>עסקה / ספק</div><div>קטגוריה</div><div>סכום</div><div></div></div>
                      {(card.transactions || []).map((transaction) => (
                        <div key={transaction.id} className="grid grid-cols-[1fr_150px_140px_44px] gap-3 border-t border-[#d8e2d2] p-3">
                          <input value={transaction.merchant} onChange={(event) => updateTransaction(card.id, transaction.id, 'merchant', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                          <select value={transaction.category} onChange={(event) => updateTransaction(card.id, transaction.id, 'category', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm outline-none focus:border-[#8fa88c]">{defaultExpenseCategories.map((category) => <option key={category}>{category}</option>)}</select>
                          <input type="number" value={transaction.amount} onChange={(event) => updateTransaction(card.id, transaction.id, 'amount', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                          <button onClick={() => removeTransaction(card.id, transaction.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button>
                        </div>
                      ))}
                    </div>
                  </div>
                  <button onClick={() => addTransaction(card.id)} className="mt-4 rounded-2xl bg-[#4d5b52] px-4 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת עסקה</button>
                </div>
              );
            })}
          </div>
        </Section>

        <Section>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-3xl font-bold">הוצאות ידניות וסיכום לפי קטגוריות</h2>
              <p className="mt-2 text-sm text-slate-500">הוצאות שלא נכנסות מכרטיסי האשראי, כמו משכנתא, העברות וצ׳קים.</p>
            </div>
            <button onClick={addManualExpense} className="rounded-2xl bg-[#7a9b76] px-5 py-3 text-sm font-semibold text-white shadow-sm">+ הוספת הוצאה ידנית</button>
          </div>

          <div className="mt-7 overflow-x-auto rounded-2xl border border-slate-100">
            <div className="min-w-[680px]">
              <div className="grid grid-cols-[1fr_150px_160px_44px] bg-slate-50 px-4 py-3 text-xs font-bold text-slate-500"><div>קטגוריה</div><div>סוג</div><div>סכום</div><div></div></div>
              {monthData.manualExpenses.map((expense) => (
                <div key={expense.id} className="grid grid-cols-[1fr_150px_160px_44px] gap-3 border-t border-slate-100 p-3">
                  <input value={expense.category} onChange={(event) => updateRow('manualExpenses', expense.id, 'category', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                  <select value={expense.type} onChange={(event) => updateRow('manualExpenses', expense.id, 'type', event.target.value)} className="rounded-xl border border-slate-200 px-3 py-3 text-sm outline-none focus:border-[#8fa88c]"><option>קבועה</option><option>משתנה</option><option>חיסכון</option><option>חד פעמית</option></select>
                  <input type="number" value={expense.amount} onChange={(event) => updateRow('manualExpenses', expense.id, 'amount', event.target.value)} className="rounded-xl border border-slate-200 px-4 py-3 text-sm outline-none focus:border-[#8fa88c]" />
                  <button onClick={() => removeRow('manualExpenses', expense.id)} className="rounded-xl bg-slate-100 text-sm font-bold text-slate-500">×</button>
                </div>
              ))}
            </div>
          </div>

          <div className="mt-8 grid gap-4 md:grid-cols-2 xl:grid-cols-3">
            {allExpenses.map((expense) => (
              <div key={`${expense.source || 'manual'}-${expense.id}`} className="rounded-2xl border border-slate-100 bg-slate-50 p-5">
                <div className="text-sm font-medium text-slate-500">{expense.type}</div>
                <div className="mt-1 text-lg font-bold">{expense.category}</div>
                <div className="mt-3 text-2xl font-bold text-[#5f7d66]">{SHEKEL.format(toNumber(expense.amount))}</div>
              </div>
            ))}
          </div>
        </Section>

        <section className="grid gap-6 xl:grid-cols-4">
          <Section className="xl:col-span-2">
            <div className="flex items-center justify-between">
              <div><h2 className="text-2xl font-bold">תחזית סוף חודש 🔮</h2><p className="mt-2 text-sm text-slate-500">תחזית לפי קצב ההוצאות הנוכחי.</p></div>
              <div className="rounded-2xl bg-[#e5eee2] px-4 py-3 text-sm font-bold text-[#4f6854]">{monthlySavings >= 0 ? 'מאוזנים' : 'סכנת מינוס'}</div>
            </div>
            <div className="mt-6 grid gap-4 md:grid-cols-3">
              <StatCard title="תחזית יתרה" value={SHEKEL.format(monthlySavings)} note="לאחר הוצאות וחסכונות" />
              <StatCard title="שווי משפחתי" value={SHEKEL.format(totalAssets)} note="נכסים שהוזנו" />
              <StatCard title="כסף פנוי" value={SHEKEL.format(Math.max(0, monthlySavings))} note="לפני החלטות נוספות" />
            </div>
          </Section>

          <Section>
            <h2 className="text-2xl font-bold">מנויים שזוהו 🧾</h2>
            <div className="mt-5 space-y-3 text-sm">
              {['Netflix', 'Spotify', 'Canva', 'iCloud'].map((item) => (
                <div key={item} className="flex items-center justify-between rounded-2xl bg-slate-50 p-4"><span>{item}</span><span className="rounded-full bg-[#e5eee2] px-3 py-1 text-xs font-bold text-[#5d755f]">מנוי פעיל</span></div>
              ))}
            </div>
          </Section>

          <Section>
            <h2 className="text-2xl font-bold">AI פיננסי 🤖</h2>
            <div className="mt-5 rounded-2xl bg-[#4d5b52] p-5 text-sm leading-relaxed text-white">אם תקטינו את וולט ב־₪400 בחודש, יעד הטיול יוקדם בכמעט 3 חודשים.</div>
            <div className="mt-4 rounded-2xl bg-[#f3f5ef] p-5 text-sm text-slate-700">הוצאות הקניות גבוהות ב־18% מהממוצע של החודשים האחרונים.</div>
          </Section>
        </section>

        <Section>
          <div><h2 className="text-3xl font-bold">תקציב מול בפועל 🎯</h2><p className="mt-2 text-sm text-slate-500">השוואה בין התקציב שהוגדר לבין מה שבאמת יצא.</p></div>
          <div className="mt-7 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
            {Object.entries(categoryBudgets).map(([category, budget]) => {
              const spent = categoryTotals[category] || 0;
              const percent = budget ? Math.round((spent / budget) * 100) : 0;
              const status = percent < 80 ? '🟢' : percent < 100 ? '🟠' : '🔴';
              return (
                <div key={category} className="rounded-3xl border border-slate-100 bg-slate-50 p-5">
                  <div className="flex items-center justify-between"><div className="text-sm font-bold">{category}</div><div>{status}</div></div>
                  <div className="mt-4 text-2xl font-bold text-slate-900">{SHEKEL.format(spent)}</div>
                  <div className="mt-1 text-sm text-slate-500">מתוך {SHEKEL.format(budget)}</div>
                  <div className="mt-4 h-3 overflow-hidden rounded-full bg-white"><div className={`h-full rounded-full ${percent < 80 ? 'bg-[#7c9780]' : percent < 100 ? 'bg-orange-400' : 'bg-red-500'}`} style={{ width: `${Math.min(percent, 100)}%` }} /></div>
                </div>
              );
            })}
          </div>
        </Section>

        <section className="grid gap-6 lg:grid-cols-3">
          <Section className="lg:col-span-2">
            <h2 className="text-2xl font-bold">תובנות אוטומטיות</h2>
            <div className="mt-5 grid gap-4 md:grid-cols-2">
              {insightCards.map((insight) => <div key={insight} className="rounded-2xl bg-[#f3f5ef] p-5 text-sm leading-relaxed text-slate-700">{insight}</div>)}
            </div>
          </Section>

          <div className="rounded-[32px] bg-[#4d5b52] p-6 text-white shadow-sm">
            <h2 className="text-2xl font-bold">טקס חודשי</h2>
            <div className="mt-5 space-y-3 text-sm text-slate-200">
              <div className="rounded-2xl bg-white/10 p-3">1. בוחרים חודש</div>
              <div className="rounded-2xl bg-white/10 p-3">2. ממלאים הכנסות</div>
              <div className="rounded-2xl bg-white/10 p-3">3. מעלים CSV של האשראי</div>
              <div className="rounded-2xl bg-white/10 p-3">4. בודקים ומאשרים קטגוריות</div>
              <div className="rounded-2xl bg-white/10 p-3">5. בודקים חיסכון ותובנות</div>
            </div>
          </div>
        </section>
      </div>
    </div>
  );
}
