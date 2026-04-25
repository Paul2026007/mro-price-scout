import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  FileSpreadsheet,
  Search,
  TrendingDown,
  Package,
  ExternalLink,
  Sparkles,
  Mail,
  FileText,
  Table as TableIcon,
  CheckCircle2,
  AlertCircle,
  Loader2,
} from "lucide-react";

// ---------- Sample data ----------
const SAMPLE_ITEMS = [
  { id: "1001", description: "Bearing 6205-2RS Sealed Ball", partNumber: "SKF-6205-2RS", quantity: 24, unitPrice: 18.5 },
  { id: "1002", description: "Industrial Safety Glasses, Clear Lens", partNumber: "3M-1611", quantity: 100, unitPrice: 4.25 },
  { id: "1003", description: "Hex Bolt 1/2-13 x 2 in, Grade 8", partNumber: "FAS-HB-1213-2-G8", quantity: 500, unitPrice: 0.92 },
  { id: "1004", description: "Hydraulic Hose 1/2 in, 3000 PSI, 25 ft", partNumber: "PAR-HH-12-3000-25", quantity: 6, unitPrice: 145.0 },
  { id: "1005", description: "Welding Gloves, Heavy Leather", partNumber: "TILL-50L", quantity: 40, unitPrice: 14.95 },
  { id: "1006", description: "Air Filter, MERV 13, 20x25x1", partNumber: "FILT-AF-2025-13", quantity: 24, unitPrice: 21.4 },
];

// Possible header names from SAP / Ariba exports, mapped to canonical fields
const HEADER_MAP = {
  id: ["item id", "item", "line", "line item", "line no", "po line", "id"],
  description: [
    "description", "item description", "material description", "short text", "product name", "name",
  ],
  partNumber: [
    "part number", "part no", "part #", "part#", "material", "material number", "mfr part", "manufacturer part",
    "supplier part", "vendor part", "sku",
  ],
  quantity: ["quantity", "qty", "order qty", "order quantity", "po qty"],
  unitPrice: [
    "unit price", "net price", "price", "unit cost", "cost", "price/unit", "price per unit",
  ],
};

// Mock supplier dataset — in production these would come from real APIs
const SUPPLIERS = [
  { name: "Grainger", domain: "grainger.com", deliveryDays: 2, rating: 4.7 },
  { name: "MSC Industrial", domain: "mscdirect.com", deliveryDays: 3, rating: 4.6 },
  { name: "Motion Industries", domain: "motionindustries.com", deliveryDays: 4, rating: 4.5 },
  { name: "Amazon Business", domain: "amazon.com", deliveryDays: 1, rating: 4.4 },
  { name: "Fastenal", domain: "fastenal.com", deliveryDays: 3, rating: 4.5 },
  { name: "Zoro", domain: "zoro.com", deliveryDays: 2, rating: 4.3 },
  { name: "Uline", domain: "uline.com", deliveryDays: 1, rating: 4.6 },
];

// ---------- Helpers ----------
const fmt = (n) =>
  Number(n || 0).toLocaleString("en-US", { style: "currency", currency: "USD" });

const normalize = (s) => String(s || "").toLowerCase().trim();

function detectColumn(headers, candidates) {
  const lowered = headers.map(normalize);
  for (const cand of candidates) {
    const idx = lowered.indexOf(cand);
    if (idx >= 0) return idx;
  }
  // partial matches
  for (let i = 0; i < lowered.length; i++) {
    for (const cand of candidates) {
      if (lowered[i].includes(cand) || cand.includes(lowered[i])) return i;
    }
  }
  return -1;
}

function parseSheetData(rows) {
  if (!rows || rows.length < 2) {
    throw new Error("File looks empty. Expected a header row and at least one data row.");
  }
  const headers = rows[0];
  const idIdx = detectColumn(headers, HEADER_MAP.id);
  const descIdx = detectColumn(headers, HEADER_MAP.description);
  const partIdx = detectColumn(headers, HEADER_MAP.partNumber);
  const qtyIdx = detectColumn(headers, HEADER_MAP.quantity);
  const priceIdx = detectColumn(headers, HEADER_MAP.unitPrice);

  if (descIdx === -1 || qtyIdx === -1 || priceIdx === -1) {
    throw new Error(
      "Could not find Description, Quantity and Unit Price columns. Check that your headers match common SAP/Ariba names."
    );
  }

  const items = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;
    const description = row[descIdx];
    if (!description) continue;
    const qty = Number(row[qtyIdx]) || 0;
    const price = Number(row[priceIdx]) || 0;
    if (qty <= 0 || price <= 0) continue;
    items.push({
      id: row[idIdx] !== undefined ? String(row[idIdx]) : String(r),
      description: String(description),
      partNumber: partIdx >= 0 ? String(row[partIdx] || "") : "",
      quantity: qty,
      unitPrice: price,
    });
  }
  if (items.length === 0) {
    throw new Error("No usable rows found. Check that quantity and unit price columns have numeric values.");
  }
  return items;
}

// Generate deterministic but realistic-looking comparison prices
function generateComparison(item) {
  const seed = (item.id + item.description).split("").reduce((a, c) => a + c.charCodeAt(0), 0);
  const rng = (n) => ((Math.sin(seed + n) + 1) / 2);

  const offers = SUPPLIERS.map((s, i) => {
    // Discount factor between -25% (cheaper) and +10% (more expensive)
    const factor = 0.75 + rng(i) * 0.35;
    const unitPrice = +(item.unitPrice * factor).toFixed(2);
    return {
      supplier: s.name,
      domain: s.domain,
      unitPrice,
      total: +(unitPrice * item.quantity).toFixed(2),
      deliveryDays: s.deliveryDays,
      rating: s.rating,
      searchUrl: `https://www.${s.domain}/search?q=${encodeURIComponent(
        item.partNumber || item.description
      )}`,
    };
  }).sort((a, b) => a.unitPrice - b.unitPrice);

  return offers;
}

// ---------- App ----------
export default function App() {
  const [tab, setTab] = useState("upload");
  const [items, setItems] = useState([]);
  const [selected, setSelected] = useState({});
  const [loadStatus, setLoadStatus] = useState({ state: "idle", message: "" });
  const [results, setResults] = useState(null);
  const [insight, setInsight] = useState("");
  const [insightLoading, setInsightLoading] = useState(false);
  const [showExport, setShowExport] = useState(false);
  const [showReport, setShowReport] = useState(false);
  const [rfq, setRfq] = useState("");
  const [rfqLoading, setRfqLoading] = useState(false);
  const fileRef = useRef(null);

  const apiKey = process.env.REACT_APP_ANTHROPIC_API_KEY;

  const totals = useMemo(() => {
    if (!results) return null;
    let currentSpend = 0;
    let bestSpend = 0;
    for (const r of results) {
      currentSpend += r.item.unitPrice * r.item.quantity;
      bestSpend += r.offers[0].total;
    }
    const savings = currentSpend - bestSpend;
    const rate = currentSpend > 0 ? (savings / currentSpend) * 100 : 0;
    return { currentSpend, bestSpend, savings, rate };
  }, [results]);

  const selectedCount = Object.values(selected).filter(Boolean).length;
  const itemsCurrentSpend = items.reduce((s, it) => s + it.unitPrice * it.quantity, 0);

  // ---------- File handling ----------
  const handleFile = async (file) => {
    if (!file) return;
    setLoadStatus({ state: "loading", message: `Reading ${file.name}...` });
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const parsed = parseSheetData(rows);
      setItems(parsed);
      const initSel = {};
      parsed.forEach((it) => (initSel[it.id] = true));
      setSelected(initSel);
      setLoadStatus({
        state: "success",
        message: `Loaded ${parsed.length} item${parsed.length === 1 ? "" : "s"} from ${file.name}.`,
      });
      setTab("items");
    } catch (err) {
      console.error(err);
      setLoadStatus({ state: "error", message: err.message || "Failed to read file." });
    }
  };

  const onDrop = (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };
  const onDragOver = (e) => e.preventDefault();

  const loadSample = () => {
    setItems(SAMPLE_ITEMS);
    const initSel = {};
    SAMPLE_ITEMS.forEach((it) => (initSel[it.id] = true));
    setSelected(initSel);
    setLoadStatus({ state: "success", message: `Loaded ${SAMPLE_ITEMS.length} sample items.` });
    setTab("items");
  };

  // ---------- Search ----------
  const runSearch = () => {
    const chosen = items.filter((it) => selected[it.id]);
    if (chosen.length === 0) return;
    const out = chosen.map((item) => ({
      item,
      offers: generateComparison(item),
    }));
    setResults(out);
    setTab("results");
    setInsight("");
    setShowExport(false);
    setShowReport(false);
    setRfq("");
    if (apiKey) generateInsight(out);
  };

  // ---------- AI insight ----------
  const generateInsight = async (data) => {
    if (!apiKey) return;
    setInsightLoading(true);
    try {
      const summary = data.slice(0, 30).map((r) => ({
        description: r.item.description,
        partNumber: r.item.partNumber,
        quantity: r.item.quantity,
        currentUnitPrice: r.item.unitPrice,
        bestSupplier: r.offers[0].supplier,
        bestUnitPrice: r.offers[0].unitPrice,
      }));
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-haiku-4-5-20251001",
          max_tokens: 600,
          messages: [
            {
              role: "user",
              content: `You are a procurement analyst. Write a concise (max 6 sentences) executive summary of savings opportunities from this MRO price comparison. Highlight the biggest wins and any anomalies. Data:\n${JSON.stringify(summary, null, 2)}`,
            },
          ],
        }),
      });
      const json = await res.json();
      const text = json?.content?.[0]?.text || "";
      setInsight(text);
    } catch (err) {
      console.error(err);
      setInsight("");
    } finally {
      setInsightLoading(false);
    }
  };

  // ---------- Bulk RFQ ----------
  const generateRFQ = async () => {
    if (!results) return;
    setRfqLoading(true);
    const top = [...results]
      .sort(
        (a, b) =>
          (b.item.unitPrice - b.offers[0].unitPrice) * b.item.quantity -
          (a.item.unitPrice - a.offers[0].unitPrice) * a.item.quantity
      )
      .slice(0, 20);

    if (apiKey) {
      try {
        const res = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "anthropic-dangerous-direct-browser-access": "true",
          },
          body: JSON.stringify({
            model: "claude-haiku-4-5-20251001",
            max_tokens: 1200,
            messages: [
              {
                role: "user",
                content: `Write a professional RFQ email to industrial suppliers asking for quotes on these MRO items. Include subject line, greeting, item table (description, part number, qty, target price), 2-week response deadline, and contact placeholders [Your Name], [Company], [Email], [Phone]. Items:\n${JSON.stringify(
                  top.map((r) => ({
                    description: r.item.description,
                    partNumber: r.item.partNumber,
                    quantity: r.item.quantity,
                    targetPrice: r.offers[0].unitPrice,
                  })),
                  null,
                  2
                )}`,
              },
            ],
          }),
        });
        const json = await res.json();
        setRfq(json?.content?.[0]?.text || buildFallbackRFQ(top));
      } catch (err) {
        setRfq(buildFallbackRFQ(top));
      }
    } else {
      setRfq(buildFallbackRFQ(top));
    }
    setRfqLoading(false);
  };

  const buildFallbackRFQ = (top) => {
    const today = new Date();
    const deadline = new Date(today.getTime() + 14 * 24 * 60 * 60 * 1000);
    const fmtDate = (d) => d.toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
    const lines = top
      .map(
        (r, i) =>
          `${i + 1}. ${r.item.description} — Part: ${r.item.partNumber || "N/A"} — Qty: ${r.item.quantity} — Target unit price: ${fmt(
            r.offers[0].unitPrice
          )}`
      )
      .join("\n");
    return `Subject: RFQ - MRO Items Bulk Quote Request - Response by ${fmtDate(deadline)}

Dear Supplier,

[Company] is requesting your best pricing on the MRO items listed below. We are benchmarking the market and would appreciate your response by ${fmtDate(
      deadline
    )}.

ITEMS:
${lines}

Please include unit price, total, lead time, and any volume discounts. Replies welcome by email.

Thank you,
[Your Name]
[Company]
[Email]
[Phone]
`;
  };

  // ---------- Render ----------
  return (
    <div style={S.app}>
      <header style={S.header}>
        <div style={S.headerInner}>
          <div style={S.logoWrap}>
            <div style={S.logoMark}>
              <Sparkles size={20} color="#fff" />
            </div>
            <div>
              <div style={S.logoTitle}>MRO Price Scout</div>
              <div style={S.logoSub}>AI-powered procurement price comparison</div>
            </div>
          </div>
          <div style={S.tabs}>
            {[
              { id: "upload", label: "1. Upload" },
              { id: "items", label: `2. Items${items.length ? ` (${items.length})` : ""}` },
              { id: "results", label: "3. Results" },
            ].map((t) => (
              <button
                key={t.id}
                style={{ ...S.tab, ...(tab === t.id ? S.tabActive : {}) }}
                onClick={() => setTab(t.id)}
              >
                {t.label}
              </button>
            ))}
          </div>
        </div>
      </header>

      <main style={S.main}>
        {tab === "upload" && (
          <UploadTab
            fileRef={fileRef}
            onDrop={onDrop}
            onDragOver={onDragOver}
            onPick={(e) => handleFile(e.target.files?.[0])}
            loadStatus={loadStatus}
            loadSample={loadSample}
          />
        )}

        {tab === "items" && (
          <ItemsTab
            items={items}
            selected={selected}
            setSelected={setSelected}
            selectedCount={selectedCount}
            currentSpend={itemsCurrentSpend}
            runSearch={runSearch}
          />
        )}

        {tab === "results" && results && (
          <ResultsTab
            results={results}
            totals={totals}
            insight={insight}
            insightLoading={insightLoading}
            apiKey={apiKey}
            showExport={showExport}
            setShowExport={setShowExport}
            showReport={showReport}
            setShowReport={setShowReport}
            rfq={rfq}
            setRfq={setRfq}
            rfqLoading={rfqLoading}
            generateRFQ={generateRFQ}
          />
        )}

        {tab === "results" && !results && (
          <div style={S.empty}>
            <Search size={36} color="#94a3b8" />
            <h3 style={{ margin: "12px 0 4px" }}>No results yet</h3>
            <p style={{ color: "#64748b", margin: 0 }}>
              Upload a file and run a price search to see comparisons here.
            </p>
          </div>
        )}
      </main>
    </div>
  );
}

// ---------- Upload tab ----------
function UploadTab({ fileRef, onDrop, onDragOver, onPick, loadStatus, loadSample }) {
  return (
    <div style={S.card}>
      <h2 style={S.h2}>Upload your SAP / Ariba export</h2>
      <p style={S.muted}>
        Drop a CSV or Excel file (.xlsx, .xls) with item descriptions, part numbers, quantities, and
        unit prices.
      </p>

      <div
        style={S.dropzone}
        onDrop={onDrop}
        onDragOver={onDragOver}
        onClick={() => fileRef.current?.click()}
      >
        <Upload size={36} color="#3b82f6" />
        <div style={{ marginTop: 12, fontWeight: 600 }}>Drop your file here</div>
        <div style={S.muted}>or click to browse — .xlsx, .xls, .csv</div>
        <input
          ref={fileRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          style={{ display: "none" }}
          onChange={onPick}
        />
      </div>

      {loadStatus.state === "loading" && (
        <div style={{ ...S.statusBar, background: "#eff6ff", color: "#1d4ed8" }}>
          <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} /> {loadStatus.message}
        </div>
      )}
      {loadStatus.state === "success" && (
        <div style={{ ...S.statusBar, background: "#ecfdf5", color: "#047857" }}>
          <CheckCircle2 size={16} /> {loadStatus.message}
        </div>
      )}
      {loadStatus.state === "error" && (
        <div style={{ ...S.statusBar, background: "#fef2f2", color: "#b91c1c" }}>
          <AlertCircle size={16} /> {loadStatus.message}
        </div>
      )}

      <div style={S.divider}>
        <span>or</span>
      </div>

      <button style={S.btnGhost} onClick={loadSample}>
        <FileSpreadsheet size={16} /> Load sample SAP export
      </button>

      <div style={S.helpCard}>
        <strong>Recognized column names:</strong>
        <ul style={{ margin: "8px 0 0 18px", color: "#475569", lineHeight: 1.7 }}>
          <li><b>Description:</b> Description, Item Description, Material Description, Short Text</li>
          <li><b>Part Number:</b> Part #, Part Number, Material, Material Number, SKU</li>
          <li><b>Quantity:</b> Qty, Quantity, Order Qty, PO Qty</li>
          <li><b>Unit Price:</b> Unit Price, Net Price, Price, Unit Cost</li>
        </ul>
      </div>

      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

// ---------- Items tab ----------
function ItemsTab({ items, selected, setSelected, selectedCount, currentSpend, runSearch }) {
  if (items.length === 0) {
    return (
      <div style={S.empty}>
        <Package size={36} color="#94a3b8" />
        <h3 style={{ margin: "12px 0 4px" }}>No items loaded yet</h3>
        <p style={{ color: "#64748b", margin: 0 }}>Go to step 1 to upload a file.</p>
      </div>
    );
  }

  const allSelected = items.every((it) => selected[it.id]);
  const toggleAll = () => {
    const next = {};
    items.forEach((it) => (next[it.id] = !allSelected));
    setSelected(next);
  };

  return (
    <div style={S.card}>
      <div style={S.itemsHeader}>
        <div>
          <h2 style={S.h2}>Review your line items</h2>
          <p style={S.muted}>
            {selectedCount} of {items.length} selected · current spend {fmt(currentSpend)}
          </p>
        </div>
        <button style={S.btnPrimary} disabled={selectedCount === 0} onClick={runSearch}>
          <Search size={16} /> Run price search
        </button>
      </div>

      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead>
            <tr>
              <th style={S.th}>
                <input type="checkbox" checked={allSelected} onChange={toggleAll} />
              </th>
              <th style={S.th}>Item</th>
              <th style={S.th}>Description</th>
              <th style={S.th}>Part Number</th>
              <th style={{ ...S.th, textAlign: "right" }}>Qty</th>
              <th style={{ ...S.th, textAlign: "right" }}>Unit Price</th>
              <th style={{ ...S.th, textAlign: "right" }}>Total</th>
            </tr>
          </thead>
          <tbody>
            {items.map((it) => (
              <tr key={it.id}>
                <td style={S.td}>
                  <input
                    type="checkbox"
                    checked={!!selected[it.id]}
                    onChange={() => setSelected({ ...selected, [it.id]: !selected[it.id] })}
                  />
                </td>
                <td style={S.td}>{it.id}</td>
                <td style={S.td}>{it.description}</td>
                <td style={{ ...S.td, fontFamily: "monospace" }}>{it.partNumber}</td>
                <td style={{ ...S.td, textAlign: "right" }}>{it.quantity}</td>
                <td style={{ ...S.td, textAlign: "right" }}>{fmt(it.unitPrice)}</td>
                <td style={{ ...S.td, textAlign: "right", fontWeight: 600 }}>
                  {fmt(it.unitPrice * it.quantity)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- Results tab ----------
function ResultsTab({
  results,
  totals,
  insight,
  insightLoading,
  apiKey,
  showExport,
  setShowExport,
  showReport,
  setShowReport,
  rfq,
  setRfq,
  rfqLoading,
  generateRFQ,
}) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      {/* KPIs */}
      <div style={S.kpiGrid}>
        <KPI label="Current spend" value={fmt(totals.currentSpend)} color="#0f172a" />
        <KPI label="Best-price spend" value={fmt(totals.bestSpend)} color="#0f172a" />
        <KPI label="Total savings" value={fmt(totals.savings)} color="#16a34a" />
        <KPI label="Savings rate" value={`${totals.rate.toFixed(1)}%`} color="#16a34a" />
      </div>

      {/* AI insight */}
      <div style={S.card}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
          <Sparkles size={18} color="#7c3aed" />
          <strong>AI Procurement Insight</strong>
        </div>
        {!apiKey && (
          <p style={S.muted}>
            Add <code>REACT_APP_ANTHROPIC_API_KEY</code> in Vercel to enable AI-generated insights and
            RFQs.
          </p>
        )}
        {apiKey && insightLoading && (
          <p style={S.muted}>
            <Loader2 size={14} style={{ animation: "spin 1s linear infinite" }} /> Generating
            insight...
          </p>
        )}
        {apiKey && !insightLoading && insight && (
          <p style={{ margin: 0, lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{insight}</p>
        )}
      </div>

      {/* Action buttons */}
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <button style={S.btnGhost} onClick={() => setShowExport(!showExport)}>
          <TableIcon size={16} /> {showExport ? "Hide" : "Export to Excel"}
        </button>
        <button style={S.btnGhost} onClick={() => setShowReport(!showReport)}>
          <FileText size={16} /> {showReport ? "Hide" : "Savings Report"}
        </button>
        <button style={S.btnPrimary} onClick={generateRFQ} disabled={rfqLoading}>
          {rfqLoading ? (
            <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} />
          ) : (
            <Mail size={16} />
          )}{" "}
          Generate Bulk RFQ
        </button>
      </div>

      {/* Export panel */}
      {showExport && <ExportPanel results={results} totals={totals} />}

      {/* Report panel */}
      {showReport && <ReportPanel results={results} totals={totals} insight={insight} />}

      {/* RFQ panel */}
      {rfq && (
        <div style={S.card}>
          <strong>Bulk RFQ Email</strong>
          <p style={{ ...S.muted, marginBottom: 8 }}>
            Edit the contact placeholders, then copy and paste into your email client.
          </p>
          <textarea
            value={rfq}
            onChange={(e) => setRfq(e.target.value)}
            style={{ ...S.textarea, minHeight: 280 }}
          />
        </div>
      )}

      {/* Result rows */}
      {results.map((r) => (
        <ResultRow key={r.item.id} row={r} />
      ))}
    </div>
  );
}

function KPI({ label, value, color }) {
  return (
    <div style={S.kpi}>
      <div style={S.kpiLabel}>{label}</div>
      <div style={{ ...S.kpiValue, color }}>{value}</div>
    </div>
  );
}

function ResultRow({ row }) {
  const { item, offers } = row;
  const best = offers[0];
  const savingsPerUnit = item.unitPrice - best.unitPrice;
  const totalSavings = savingsPerUnit * item.quantity;
  const savingsPct = item.unitPrice > 0 ? (savingsPerUnit / item.unitPrice) * 100 : 0;

  return (
    <div style={S.card}>
      <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
        <div>
          <div style={{ fontWeight: 600 }}>{item.description}</div>
          <div style={S.muted}>
            Part {item.partNumber || "—"} · Qty {item.quantity} · Current {fmt(item.unitPrice)}
          </div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ color: totalSavings > 0 ? "#16a34a" : "#dc2626", fontWeight: 700 }}>
            {totalSavings > 0 ? "Save " : ""}
            {fmt(totalSavings)}
          </div>
          <div style={S.muted}>{savingsPct.toFixed(1)}% per unit</div>
        </div>
      </div>

      <div style={S.offerGrid}>
        {offers.map((o, i) => (
          <div
            key={o.supplier}
            style={{
              ...S.offer,
              ...(i === 0 ? S.offerBest : {}),
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontWeight: 600 }}>{o.supplier}</div>
              {i === 0 && (
                <span style={S.bestTag}>
                  <TrendingDown size={12} /> Best
                </span>
              )}
            </div>
            <div style={{ fontSize: 22, fontWeight: 700, marginTop: 4 }}>{fmt(o.unitPrice)}</div>
            <div style={S.muted}>per unit · total {fmt(o.total)}</div>
            <div style={{ ...S.muted, fontSize: 12, marginTop: 4 }}>
              {o.deliveryDays} day delivery · {o.rating}★
            </div>
            <a href={o.searchUrl} target="_blank" rel="noreferrer" style={S.searchLink}>
              <ExternalLink size={12} /> Search on {o.supplier}
            </a>
          </div>
        ))}
      </div>
    </div>
  );
}

// ---------- Export panel (HTML table for copy-paste into Excel) ----------
function ExportPanel({ results, totals }) {
  const tableRef = useRef(null);
  const selectAll = () => {
    if (!tableRef.current) return;
    const range = document.createRange();
    range.selectNodeContents(tableRef.current);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  };

  return (
    <div style={S.card}>
      <strong>Excel-ready table</strong>
      <p style={S.muted}>
        Click the table to select all rows, then Ctrl+C / Cmd+C and paste directly into Excel.
      </p>
      <div onClick={selectAll} style={{ overflowX: "auto" }}>
        <table ref={tableRef} style={S.table}>
          <thead>
            <tr>
              <th style={S.th}>Item</th>
              <th style={S.th}>Description</th>
              <th style={S.th}>Part Number</th>
              <th style={S.th}>Qty</th>
              <th style={S.th}>Current Unit</th>
              <th style={S.th}>Best Supplier</th>
              <th style={S.th}>Best Unit</th>
              <th style={S.th}>Savings / unit</th>
              <th style={S.th}>Total Savings</th>
              <th style={S.th}>Delivery</th>
            </tr>
          </thead>
          <tbody>
            {results.map((r) => {
              const b = r.offers[0];
              const sper = r.item.unitPrice - b.unitPrice;
              return (
                <tr key={r.item.id}>
                  <td style={S.td}>{r.item.id}</td>
                  <td style={S.td}>{r.item.description}</td>
                  <td style={S.td}>{r.item.partNumber}</td>
                  <td style={S.td}>{r.item.quantity}</td>
                  <td style={S.td}>{fmt(r.item.unitPrice)}</td>
                  <td style={S.td}>{b.supplier}</td>
                  <td style={S.td}>{fmt(b.unitPrice)}</td>
                  <td style={S.td}>{fmt(sper)}</td>
                  <td style={S.td}>{fmt(sper * r.item.quantity)}</td>
                  <td style={S.td}>{b.deliveryDays} days</td>
                </tr>
              );
            })}
            <tr style={{ background: "#f1f5f9", fontWeight: 700 }}>
              <td style={S.td} colSpan={8}>
                Totals
              </td>
              <td style={S.td}>{fmt(totals.savings)}</td>
              <td style={S.td}></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- Inline savings report ----------
function ReportPanel({ results, totals, insight }) {
  return (
    <div style={S.card}>
      <h2 style={{ ...S.h2, marginBottom: 4 }}>Savings Report</h2>
      <div style={S.muted}>{new Date().toLocaleDateString("en-US")}</div>

      <div style={S.kpiGrid}>
        <KPI label="Current spend" value={fmt(totals.currentSpend)} color="#0f172a" />
        <KPI label="Best-price spend" value={fmt(totals.bestSpend)} color="#0f172a" />
        <KPI label="Total savings" value={fmt(totals.savings)} color="#16a34a" />
        <KPI label="Savings rate" value={`${totals.rate.toFixed(1)}%`} color="#16a34a" />
      </div>

      {insight && (
        <div style={{ ...S.card, marginTop: 12, background: "#faf5ff" }}>
          <strong>Procurement Insight</strong>
          <p style={{ marginTop: 6, lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{insight}</p>
        </div>
      )}
    </div>
  );
}

// ---------- Styles ----------
const S = {
  app: { minHeight: "100vh", background: "#f8fafc" },
  header: { background: "#fff", borderBottom: "1px solid #e2e8f0" },
  headerInner: {
    maxWidth: 1200,
    margin: "0 auto",
    padding: "16px 20px",
    display: "flex",
    flexWrap: "wrap",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 12,
  },
  logoWrap: { display: "flex", alignItems: "center", gap: 12 },
  logoMark: {
    width: 40,
    height: 40,
    borderRadius: 10,
    background: "linear-gradient(135deg, #3b82f6 0%, #7c3aed 100%)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  logoTitle: { fontSize: 18, fontWeight: 700 },
  logoSub: { fontSize: 12, color: "#64748b" },
  tabs: { display: "flex", gap: 4, background: "#f1f5f9", borderRadius: 10, padding: 4 },
  tab: {
    padding: "8px 14px",
    borderRadius: 8,
    border: "none",
    background: "transparent",
    cursor: "pointer",
    fontSize: 14,
    fontWeight: 500,
    color: "#475569",
  },
  tabActive: { background: "#fff", color: "#0f172a", boxShadow: "0 1px 2px rgba(0,0,0,0.05)" },
  main: { maxWidth: 1200, margin: "0 auto", padding: "24px 20px", display: "flex", flexDirection: "column", gap: 16 },
  card: {
    background: "#fff",
    border: "1px solid #e2e8f0",
    borderRadius: 12,
    padding: 20,
    boxShadow: "0 1px 2px rgba(0,0,0,0.03)",
  },
  h2: { margin: "0 0 4px", fontSize: 20 },
  muted: { color: "#64748b", margin: "4px 0", fontSize: 14 },
  dropzone: {
    border: "2px dashed #cbd5e1",
    borderRadius: 12,
    padding: 40,
    textAlign: "center",
    cursor: "pointer",
    background: "#f8fafc",
    marginTop: 12,
  },
  statusBar: {
    marginTop: 12,
    padding: "10px 14px",
    borderRadius: 8,
    display: "flex",
    alignItems: "center",
    gap: 8,
    fontSize: 14,
  },
  divider: {
    textAlign: "center",
    color: "#94a3b8",
    margin: "16px 0",
    fontSize: 12,
    textTransform: "uppercase",
    letterSpacing: 1,
  },
  helpCard: {
    marginTop: 16,
    padding: 14,
    borderRadius: 8,
    background: "#f1f5f9",
    fontSize: 14,
  },
  btnPrimary: {
    padding: "10px 16px",
    background: "#0f172a",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    cursor: "pointer",
    fontWeight: 600,
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
  },
  btnGhost: {
    padding: "10px 16px",
    background: "#fff",
    color: "#0f172a",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    cursor: "pointer",
    fontWeight: 600,
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
  },
  itemsHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 12,
    marginBottom: 12,
  },
  tableWrap: { overflowX: "auto" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 14 },
  th: {
    textAlign: "left",
    padding: "10px 8px",
    borderBottom: "1px solid #e2e8f0",
    color: "#475569",
    fontWeight: 600,
    fontSize: 12,
    textTransform: "uppercase",
    letterSpacing: 0.4,
    background: "#f8fafc",
  },
  td: { padding: "10px 8px", borderBottom: "1px solid #f1f5f9" },
  empty: {
    background: "#fff",
    border: "1px dashed #cbd5e1",
    borderRadius: 12,
    padding: 40,
    textAlign: "center",
  },
  kpiGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
    gap: 12,
  },
  kpi: { background: "#fff", border: "1px solid #e2e8f0", borderRadius: 12, padding: 16 },
  kpiLabel: { color: "#64748b", fontSize: 12, textTransform: "uppercase", letterSpacing: 0.4 },
  kpiValue: { fontSize: 24, fontWeight: 700, marginTop: 4 },
  offerGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
    gap: 10,
    marginTop: 14,
  },
  offer: {
    border: "1px solid #e2e8f0",
    borderRadius: 10,
    padding: 12,
    background: "#fff",
  },
  offerBest: { border: "2px solid #16a34a", background: "#f0fdf4" },
  bestTag: {
    background: "#16a34a",
    color: "#fff",
    fontSize: 11,
    padding: "2px 8px",
    borderRadius: 999,
    display: "inline-flex",
    alignItems: "center",
    gap: 4,
  },
  searchLink: {
    display: "inline-flex",
    alignItems: "center",
    gap: 4,
    marginTop: 8,
    fontSize: 12,
    color: "#3b82f6",
    textDecoration: "none",
  },
  textarea: {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    padding: 12,
    fontSize: 14,
    fontFamily: "inherit",
    resize: "vertical",
  },
};
