import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Programas de Ingeniería en Colombia | ACOFI",
  description: "Publicación estadística oficial y directorio nacional de programas de pregrado y posgrado en ingeniería — ACOFI.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es">
      <body className="antialiased font-serif">
        {children}
      </body>
    </html>
  );
}