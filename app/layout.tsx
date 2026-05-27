import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Conversor Bradesco PIX",
  description: "Gerador de remessa CNAB240 Pix para Bradesco",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="pt-BR">
      <body>{children}</body>
    </html>
  );
}
