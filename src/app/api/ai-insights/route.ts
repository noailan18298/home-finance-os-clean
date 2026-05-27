import OpenAI from 'openai';
import { NextResponse } from 'next/server';

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

export async function POST(req: Request) {
  try {
    const body = await req.json();

    const transactions = body.transactions || [];
    const totals = body.totals || {};
    const budgets = body.categoryBudgets || {};

    if (!transactions.length) {
      return NextResponse.json({
        insights: ['לא התקבלו עסקאות לניתוח.'],
      });
    }

    const prompt = `
אתה אנליסט פיננסי אישי למשפחה בישראל.

נתוני עסקאות:
${JSON.stringify(transactions, null, 2)}

סיכומים:
${JSON.stringify(totals, null, 2)}

תקציבים:
${JSON.stringify(budgets, null, 2)}

תן:
1. 5 תובנות פיננסיות אמיתיות
2. זיהוי חריגות
3. המלצות לחיסכון
4. קטגוריות בעייתיות
5. בתי עסק חריגים
6. תובנה כללית על מצב החודש

החזר JSON בלבד בפורמט:
{
  "insights": [
    "..."
  ]
}
`;

    const completion = await openai.chat.completions.create({
      model: 'gpt-4.1-mini',
      temperature: 0.4,
      response_format: { type: 'json_object' },
      messages: [
        {
          role: 'system',
          content:
            'אתה AI פיננסי מדויק. אל תמציא נתונים שלא קיימים.',
        },
        {
          role: 'user',
          content: prompt,
        },
      ],
    });

    const content = completion.choices[0]?.message?.content;

    if (!content) {
      throw new Error('No AI response');
    }

    return NextResponse.json(JSON.parse(content));
  } catch (error) {
    console.error(error);

    return NextResponse.json(
      {
        insights: [
          'ה־AI לא הצליח לנתח את הנתונים כרגע.',
        ],
      },
      { status: 500 }
    );
  }
}
