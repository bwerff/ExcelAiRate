import type { Metadata } from "next";
import { Inter } from "next/font/google";
import "./globals.css";

const inter = Inter({ subsets: ["latin"] });

export const metadata: Metadata = {
  title: "ExcelAI Rate - Modern Spreadsheet Experience",
  description: "Experience the future of spreadsheets with AI-powered features",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en" className="h-full">
      <body className={`${inter.className} h-full overflow-hidden`}>
        <div id="__next" className="h-full">
          {children}
        </div>
      </body>
    </html>
  );
}
