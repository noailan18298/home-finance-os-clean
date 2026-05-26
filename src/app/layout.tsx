import './globals.css';

export const metadata = {
  title: 'Home Finance OS',
  description: 'Family finance dashboard',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="he" dir="rtl">
      <body>{children}</body>
    </html>
  );
}
