# Penang Property Market Intelligence Dashboard

A Next.js 14 dashboard for Penang property market data (NAPIC 1995-2024) with interactive Recharts visualisations and AI analysis powered by the Anthropic Claude API.

## Features

- **8 data sections** — Overview, Transactions, Residential, Price Index, Commercial, Industrial, Hotel & Tourism, Demographics
- **District filter** — All Penang / Penang Island / Seberang Perai
- **Per-section PDF reports** — AI-written analyst report downloaded as a printable HTML file
- **Forecasting Agent** — Linear regression projections to 2027 with Claude narrative
- **Investment Screener** — AI scores any property across 6 dimensions
- **Data Refresh Agent** — Market intelligence check + NAPIC Excel upload
- **AI Chat** — Conversational analyst powered by Claude Sonnet

## Setup

### 1. Clone and install

```bash
git clone <repo-url>
cd penangresearch
npm install
```

### 2. Add your Anthropic API key

```bash
cp .env.local.example .env.local
```

Edit `.env.local` and replace `your_key_here` with your actual key:

```
ANTHROPIC_API_KEY=sk-ant-...
```

Get a key at https://console.anthropic.com

### 3. Run locally

```bash
npm run dev
```

Open http://localhost:3000

## How it works (API key security)

The Anthropic API key is **never exposed to the browser**. All Claude requests go through the Next.js server-side API route at `/api/claude`, which reads `ANTHROPIC_API_KEY` from the server environment. The frontend only calls `/api/claude`, never the Anthropic API directly.

## Deploy to Vercel

1. Push the repo to GitHub
2. Import the project at https://vercel.com/new
3. In **Environment Variables**, add:
   - `ANTHROPIC_API_KEY` = your key
4. Click **Deploy**

Vercel automatically uses the App Router and the `/api/claude` route as a serverless function.

## PDF Reports

Each data section has a **Download PDF Report** button. Clicking it:
1. Calls Claude to write a professional analyst report for that section
2. Generates a self-contained HTML file with an SVG chart and the AI narrative
3. Downloads the file — open it in a browser and use **Print → Save as PDF**

## Stack

- Next.js 14 (App Router)
- Recharts
- lucide-react
- xlsx (Excel upload)
- @anthropic-ai/sdk (server-side only)
