import "./globals.css";

export const metadata = {
  title: "Penang Property Market Intelligence Dashboard",
  description: "NAPIC Penang property data dashboard with AI analysis powered by Claude",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
