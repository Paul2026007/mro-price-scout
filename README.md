# MRO Price Scout

AI-powered price comparison for MRO procurement. Upload your SAP / Ariba export, and the tool benchmarks each line item against major industrial suppliers, surfaces the best price, and generates a ready-to-send bulk RFQ.

## What you get
- Drop-in upload for `.xlsx`, `.xls`, and `.csv` SAP/Ariba exports
- Smart column detection (handles `Net Price`, `Material`, `Order Qty`, etc.)
- Side-by-side supplier comparison with savings per unit and total
- Direct supplier search links to verify live pricing
- AI Procurement Insight (executive summary)
- Bulk RFQ email generator
- Excel-ready export panel (HTML table, paste straight into Excel)

## Run locally

```bash
npm install
npm start
```

App opens at `http://localhost:3000`.

## Deploy on Vercel
1. Push this folder to GitHub (see `UPLOAD_INSTRUCTIONS.md`)
2. On vercel.com, click **Add New Project** -> import the GitHub repo
3. In **Settings -> Environment Variables**, add:
   - Key: `REACT_APP_ANTHROPIC_API_KEY`
   - Value: your key from https://console.anthropic.com/
4. Click **Deploy**

The app works without an API key — the AI insight and RFQ generator simply fall back to a non-AI version if the key is missing.

## Project structure
```
mro-price-scout/
  package.json
  vercel.json
  .env.example
  .gitignore
  public/index.html
  src/
    index.js
    index.css
    App.jsx
```
