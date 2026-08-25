import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Lab AI — tahlil izohlovchi",
  description: "Laboratoriya tahlil natijalarini avtomatik solishtirib, AI izoh yozadi.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="uz">
      <body>{children}</body>
    </html>
  );
}
