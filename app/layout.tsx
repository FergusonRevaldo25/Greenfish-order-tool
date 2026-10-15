import "./globals.css";
import { Figtree } from "next/font/google";

const font = Figtree({ subsets: ["latin"], display: "swap" });

export const metadata = { title: "Greenfish daily sales sheet" };

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className={font.className}>{children}</body>
    </html>
  );
}
