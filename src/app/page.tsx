'use client';

import { useEffect, useMemo, useState } from 'react';

export default function Page() {
  const [income, setIncome] = useState(25000);
  const [expenses, setExpenses] = useState(18000);

  const savings = income - expenses;

  useEffect(() => {
    console.log('Dashboard loaded');
  }, []);

  return (
    <main
      style={{
        minHeight: '100vh',
        background: '#f4f1ea',
        padding: '40px',
        direction: 'rtl',
        fontFamily: 'Arial',
      }}
    >
      <div
        style={{
          maxWidth: '900px',
          margin: '0 auto',
          background: 'white',
          borderRadius: '24px',
          padding: '32px',
          boxShadow: '0 10px 30px rgba(0,0,0,0.08)',
        }}
      >
        <h1
          style={{
            fontSize: '42px',
            marginBottom: '10px',
            color: '#4d5b52',
          }}
        >
          הדשבורד הכלכלי שלנו 🌿
        </h1>

        <p
          style={{
            color: '#666',
            marginBottom: '30px',
          }}
        >
          מערכת לניהול פיננסי משפחתי
        </p>

        <div
          style={{
            display: 'grid',
            gap: '20px',
            gridTemplateColumns: '1fr 1fr',
          }}
        >
          <div
            style={{
              background: '#f3f5ef',
              padding: '24px',
              borderRadius: '20px',
            }}
          >
            <h2>הכנסות</h2>

            <input
              type="number"
              value={income}
              onChange={(e) => setIncome(Number(e.target.value))}
              style={{
                width: '100%',
                padding: '14px',
                borderRadius: '12px',
                border: '1px solid #ccc',
                marginTop: '10px',
              }}
            />
          </div>

          <div
            style={{
              background: '#f3f5ef',
              padding: '24px',
              borderRadius: '20px',
            }}
          >
            <h2>הוצאות</h2>

            <input
              type="number"
              value={expenses}
              onChange={(e) => setExpenses(Number(e.target.value))}
              style={{
                width: '100%',
                padding: '14px',
                borderRadius: '12px',
                border: '1px solid #ccc',
                marginTop: '10px',
              }}
            />
          </div>
        </div>

        <div
          style={{
            marginTop: '30px',
            background: '#4d5b52',
            color: 'white',
            padding: '30px',
            borderRadius: '24px',
          }}
        >
          <div style={{ fontSize: '18px', opacity: 0.8 }}>
            חיסכון חודשי
          </div>

          <div
            style={{
              fontSize: '54px',
              fontWeight: 'bold',
              marginTop: '10px',
            }}
          >
            ₪{savings.toLocaleString()}
          </div>
        </div>
      </div>
    </main>
  );
}
