import "./globals.css";

export const metadata = {
  title: "Gamuda Property Market Intelligence",
  description: "Gamuda property market intelligence dashboard with AI analysis powered by Claude",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
