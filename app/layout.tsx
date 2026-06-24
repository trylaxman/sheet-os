import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel to CRUD Magic Demo",
  description: "Upload any Excel file and instantly turn it into a browser-only CRUD table.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>
        {children}
        <footer className="border-t bg-white">
          <div className="mx-auto max-w-7xl px-6 py-5 text-center">
            <p className="text-sm font-semibold text-gray-800">
              SheetOS — Turn Any Excel Into a Working App
            </p>

            <div className="mt-3 flex flex-wrap items-center justify-center gap-3 text-xs text-gray-600">
              <span>🔒 Privacy First</span>
              <span>•</span>
              <span>⚡ Instant Processing</span>
              <span>•</span>
              <span>📊 Dynamic CRUD Generation</span>
            </div>

            <p className="mt-3 text-xs text-gray-500">
              Files never leave your browser. No uploads. No database. No sign-up.
            </p>
          </div>
        </footer>
      </body>
    </html>
  );
}
