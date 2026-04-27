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

// Suppliers searched by the live price-lookup feature.
// Fastenal is listed first as the user's preferred negotiated-pricing supplier.
// `search(q)` builds the correct search URL per supplier so the manual fallback
// link actually works (Grainger uses ?searchQuery=, others differ).
const SUPPLIERS = [
  {
    name: "Fastenal",
    domain: "fastenal.com",
    deliveryDays: 3,
    rating: 4.5,
    preferred: true,
    search: (q) => `https://www.fastenal.com/products?term=${encodeURIComponent(q)}`,
  },
  {
    name: "Grainger",
    domain: "grainger.com",
    deliveryDays: 2,
    rating: 4.7,
    search: (q) => `https://www.grainger.com/search?searchQuery=${encodeURIComponent(q)}`,
  },
  {
    name: "MSC Industrial",
    domain: "mscdirect.com",
    deliveryDays: 3,
    rating: 4.6,
    search: (q) => `https://www.mscdirect.com/browse/tn?searchterm=${encodeURIComponent(q)}`,
  },
  {
    name: "Motion Industries",
    domain: "motionindustries.com",
    deliveryDays: 4,
    rating: 4.5,
    search: (q) => `https://www.motionindustries.com/search?searchTerm=${encodeURIComponent(q)}`,
  },
  {
    name: "Zoro",
    domain: "zoro.com",
    deliveryDays: 2,
    rating: 4.3,
    search: (q) => `https://www.zoro.com/search?q=${encodeURIComponent(q)}`,
  },
  {
    name: "Uline",
    domain: "uline.com",
    deliveryDays: 1,
    rating: 4.6,
    search: (q) => `https://www.uline.com/Browse_Search.aspx?keywords=${encodeURIComponent(q)}`,
  },
  {
    name: "Amazon Business",
    domain: "amazon.com",
    deliveryDays: 1,
    rating: 4.4,
    search: (q) => `https://www.amazon.com/s?k=${encodeURIComponent(q)}`,
  },
];

// How many items to look up in parallel. With Haiku as the search model we get
// a higher per-minute token budget, so 3 in parallel comfortably fits and gives
// us ~45s for a 10-item file.
const LOOKUP_CONCURRENCY = 3;
const LOOKUP_INTER_ITEM_DELAY_MS = 0;

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

// Call the serverless /api/ai endpoint with action=priceLookup.
// Returns an array of offers (some may have found:false) sourced from real
// supplier websites via the Anthropic web_search tool.
async function lookupItemPrices(item, password) {
  const res = await fetch("/api/ai", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      password,
      action: "priceLookup",
      item: {
        description: item.description,
        partNumber: item.partNumber,
        quantity: item.quantity,
      },
      suppliers: SUPPLIERS.map((s) => ({ name: s.name, domain: s.domain })),
    }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || `Lookup failed (${res.status})`);
  }
  const data = await res.json();
  return data.offers || [];
}

// Take the raw offers from the API and add UI-side fields (delivery, rating,
// computed total, manual-search fallback URL) and sort: real matches by
// ascending price first, no-match suppliers last.
function enrichOffers(rawOffers, item) {
  const supMap = Object.fromEntries(SUPPLIERS.map((s) => [s.name, s]));
  const query = item.partNumber || item.description;
  const enriched = (rawOffers || []).map((o) => {
    const sup = supMap[o.supplier] || {};
    const fallbackUrl = sup.search ? sup.search(query) : null;
    if (o.found) {
      return {
        ...o,
        unitPrice: Number(o.unitPrice) || 0,
        total: +((Number(o.unitPrice) || 0) * item.quantity).toFixed(2),
        deliveryDays: sup.deliveryDays,
        rating: sup.rating,
        preferred: !!sup.preferred,
        fallbackUrl,
      };
    }
    return {
      ...o,
      deliveryDays: sup.deliveryDays,
      rating: sup.rating,
      preferred: !!sup.preferred,
      fallbackUrl,
    };
  });
  enriched.sort((a, b) => {
    if (a.found && !b.found) return -1;
    if (!a.found && b.found) return 1;
    if (a.found && b.found) return a.unitPrice - b.unitPrice;
    return 0;
  });
  return enriched;
}

// ---------- App ----------
function MROApp({ password }) {
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

  // The shared password from the gate enables AI features. The actual
  // Anthropic API key never reaches the browser — it lives in /api/ai.js.
  const apiKey = password;

  const totals = useMemo(() => {
    if (!results) return null;
    let currentSpend = 0;
    let bestSpend = 0;
    let coveredItems = 0;
    let pendingItems = 0;
    let unmatchedItems = 0;
    let erroredItems = 0;
    for (const r of results) {
      currentSpend += r.item.unitPrice * r.item.quantity;
      if (r.status === "loading") {
        pendingItems++;
        bestSpend += r.item.unitPrice * r.item.quantity; // assume parity until done
        continue;
      }
      if (r.status === "error") {
        erroredItems++;
        bestSpend += r.item.unitPrice * r.item.quantity; // lookup failed → keep current
        continue;
      }
      const bestFound = (r.offers || []).find((o) => o.found);
      if (bestFound) {
        bestSpend += bestFound.total;
        coveredItems++;
      } else {
        bestSpend += r.item.unitPrice * r.item.quantity; // no match → keep current
        unmatchedItems++;
      }
    }
    const savings = currentSpend - bestSpend;
    const rate = currentSpend > 0 ? (savings / currentSpend) * 100 : 0;
    return {
      currentSpend,
      bestSpend,
      savings,
      rate,
      coveredItems,
      pendingItems,
      unmatchedItems,
      erroredItems,
    };
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

  // ---------- Search (real web-search lookup) ----------
  const runSearch = async () => {
    const chosen = items.filter((it) => selected[it.id]);
    if (chosen.length === 0) return;

    // Seed results with loading status so the UI renders progress immediately
    const initial = chosen.map((item) => ({
      item,
      offers: [],
      status: "loading",
      error: null,
    }));
    setResults(initial);
    setTab("results");
    setInsight("");
    setShowExport(false);
    setShowReport(false);
    setRfq("");

    // Process items with bounded concurrency
    const queue = chosen.map((item, idx) => ({ item, idx }));
    const updateRow = (idx, patch) => {
      setResults((cur) => {
        if (!cur) return cur;
        const next = [...cur];
        next[idx] = { ...next[idx], ...patch };
        return next;
      });
    };

    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

    const worker = async () => {
      while (queue.length > 0) {
        const job = queue.shift();
        if (!job) break;
        try {
          const raw = await lookupItemPrices(job.item, password);
          const offers = enrichOffers(raw, job.item);
          updateRow(job.idx, { offers, status: "ready", error: null });
        } catch (err) {
          updateRow(job.idx, {
            offers: [],
            status: "error",
            error: err.message || "Lookup failed",
          });
        }
        // Breather between items so we stay under Anthropic's per-minute limit
        if (queue.length > 0) await sleep(LOOKUP_INTER_ITEM_DELAY_MS);
      }
    };

    const workers = Array.from({ length: LOOKUP_CONCURRENCY }, worker);
    await Promise.all(workers);

    // After all lookups finish, generate the AI insight summary
    if (apiKey) {
      setResults((cur) => {
        if (cur) generateInsight(cur);
        return cur;
      });
    }
  };

  // ---------- AI insight ----------
  const generateInsight = async (data) => {
    if (!apiKey) return;
    setInsightLoading(true);
    try {
      const summary = data.slice(0, 30).map((r) => {
        const best = (r.offers || []).find((o) => o.found);
        let lookupStatus;
        if (r.status === "error") lookupStatus = "lookup errored (rate limit or transient failure)";
        else if (best) lookupStatus = "matched";
        else lookupStatus = "no supplier match found";
        return {
          description: r.item.description,
          partNumber: r.item.partNumber,
          quantity: r.item.quantity,
          currentUnitPrice: r.item.unitPrice,
          lookupStatus,
          bestSupplier: best ? best.supplier : null,
          bestUnitPrice: best ? best.unitPrice : null,
          matchConfidence: best ? best.confidence : null,
        };
      });
      const matchedCount = summary.filter((s) => s.lookupStatus === "matched").length;
      const noMatchCount = summary.filter((s) => s.lookupStatus === "no supplier match found").length;
      const erroredCount = summary.filter((s) => s.lookupStatus.startsWith("lookup errored")).length;
      const res = await fetch("/api/ai", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          password,
          maxTokens: 600,
          prompt: `You are a procurement analyst. Write a concise (max 6 sentences) executive summary of savings opportunities from this MRO price comparison.

Lookup outcome: ${matchedCount} matched, ${noMatchCount} no supplier match, ${erroredCount} errored during lookup. Treat 'errored' rows as 'unknown — needs re-run', NOT as 'no match'. Only flag the items genuinely marked 'no supplier match found' as needing manual sourcing.

Highlight the biggest wins and any anomalies (duplicate part numbers with different prices, extreme price gaps, etc.). Data:
${JSON.stringify(summary, null, 2)}`,
        }),
      });
      const json = await res.json();
      setInsight(json?.text || "");
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

    // Pick rows that have a real match, sorted by total savings opportunity desc.
    // Rows with no match still get included so the supplier is asked to quote them
    // (just no target price).
    const rowsWithBest = results.map((r) => ({
      r,
      best: (r.offers || []).find((o) => o.found),
    }));
    const ranked = rowsWithBest.sort((a, b) => {
      const sa = a.best ? (a.r.item.unitPrice - a.best.unitPrice) * a.r.item.quantity : 0;
      const sb = b.best ? (b.r.item.unitPrice - b.best.unitPrice) * b.r.item.quantity : 0;
      return sb - sa;
    });
    const top = ranked.slice(0, 20);

    if (apiKey) {
      try {
        const res = await fetch("/api/ai", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            password,
            maxTokens: 1500,
            prompt: `Write a professional RFQ email to an industrial supplier asking for quotes on these MRO items. Include subject line, greeting, item table (description, part number, qty, target price if available), 2-week response deadline, and contact placeholders [Your Name], [Company], [Email], [Phone]. Some items show "no target price" — for those, ask the supplier to quote their best price. Items:\n${JSON.stringify(
              top.map(({ r, best }) => ({
                description: r.item.description,
                partNumber: r.item.partNumber,
                quantity: r.item.quantity,
                targetPrice: best ? best.unitPrice : "no target price",
                referencedSupplier: best ? best.supplier : null,
              })),
              null,
              2
            )}`,
          }),
        });
        const json = await res.json();
        setRfq(json?.text || buildFallbackRFQ(top.map((t) => t.r)));
      } catch (err) {
        setRfq(buildFallbackRFQ(top.map((t) => t.r)));
      }
    } else {
      setRfq(buildFallbackRFQ(top.map((t) => t.r)));
    }
    setRfqLoading(false);
  };

  const buildFallbackRFQ = (top) => {
    const today = new Date();
    const deadline = new Date(today.getTime() + 14 * 24 * 60 * 60 * 1000);
    const fmtDate = (d) => d.toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
    const lines = top
      .map((r, i) => {
        const best = (r.offers || []).find((o) => o.found);
        const target = best ? `Target unit price: ${fmt(best.unitPrice)}` : "Please quote your best price";
        return `${i + 1}. ${r.item.description} — Part: ${r.item.partNumber || "N/A"} — Qty: ${r.item.quantity} — ${target}`;
      })
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
          <p style={{ ...S.muted, marginTop: 4, fontSize: 12 }}>
            Search uses real web lookups across Fastenal, Grainger, MSC, Motion, Zoro, Uline, and
            Amazon Business. Allow ~10–20 seconds per item.
          </p>
        </div>
        <button style={S.btnPrimary} disabled={selectedCount === 0} onClick={runSearch}>
          <Search size={16} /> Run live price search
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

      {/* Lookup status line */}
      {(totals.pendingItems > 0 ||
        totals.unmatchedItems > 0 ||
        totals.erroredItems > 0) && (
        <div
          style={{
            ...S.card,
            background: totals.pendingItems > 0 ? "#fef9c3" : "#f8fafc",
            display: "flex",
            alignItems: "center",
            gap: 8,
            padding: 12,
          }}
        >
          {totals.pendingItems > 0 ? (
            <>
              <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} />
              <span>
                Searching real prices… {totals.coveredItems} of{" "}
                {totals.coveredItems +
                  totals.pendingItems +
                  totals.unmatchedItems +
                  totals.erroredItems}{" "}
                items priced
                {totals.unmatchedItems > 0 ? ` · ${totals.unmatchedItems} no match` : ""}
                {totals.erroredItems > 0 ? ` · ${totals.erroredItems} errored` : ""}. Numbers
                update as searches complete.
              </span>
            </>
          ) : (
            <>
              <AlertCircle size={16} color="#b45309" />
              <span>
                {totals.unmatchedItems > 0 && (
                  <>
                    {totals.unmatchedItems} item
                    {totals.unmatchedItems === 1 ? "" : "s"} had no clear supplier match
                  </>
                )}
                {totals.unmatchedItems > 0 && totals.erroredItems > 0 ? " · " : ""}
                {totals.erroredItems > 0 && (
                  <>
                    {totals.erroredItems} item{totals.erroredItems === 1 ? "" : "s"} errored
                    during lookup (often a transient rate limit — re-run those rows)
                  </>
                )}
                . See the rows flagged below.
              </span>
            </>
          )}
        </div>
      )}

      {/* AI insight */}
      <div style={S.card}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
          <Sparkles size={18} color="#7c3aed" />
          <strong>AI Procurement Insight</strong>
        </div>
        {!apiKey && (
          <p style={S.muted}>
            AI features unavailable — server is missing required configuration.
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
        <button
          style={{
            ...S.btnPrimary,
            opacity: totals.pendingItems > 0 || rfqLoading ? 0.5 : 1,
            cursor: totals.pendingItems > 0 || rfqLoading ? "not-allowed" : "pointer",
          }}
          onClick={generateRFQ}
          disabled={rfqLoading || totals.pendingItems > 0}
          title={totals.pendingItems > 0 ? "Wait for all price lookups to finish" : ""}
        >
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

function ConfidenceBadge({ level }) {
  const palette = {
    high: { bg: "#dcfce7", fg: "#166534", label: "High match" },
    medium: { bg: "#fef9c3", fg: "#854d0e", label: "Medium match" },
    low: { bg: "#fee2e2", fg: "#991b1b", label: "Low match" },
  };
  const p = palette[level] || { bg: "#e2e8f0", fg: "#334155", label: level || "—" };
  return (
    <span
      style={{
        background: p.bg,
        color: p.fg,
        fontSize: 10,
        fontWeight: 700,
        padding: "2px 6px",
        borderRadius: 4,
        textTransform: "uppercase",
        letterSpacing: "0.03em",
      }}
    >
      {p.label}
    </span>
  );
}

function ResultRow({ row }) {
  const { item, offers, status, error } = row;

  // Loading state — search still running for this item
  if (status === "loading") {
    return (
      <div style={S.card}>
        <div style={{ fontWeight: 600 }}>{item.description}</div>
        <div style={S.muted}>
          Part {item.partNumber || "—"} · Qty {item.quantity} · Current {fmt(item.unitPrice)}
        </div>
        <div style={{ ...S.muted, marginTop: 10, display: "flex", alignItems: "center", gap: 6 }}>
          <Loader2 size={14} style={{ animation: "spin 1s linear infinite" }} />
          Searching suppliers for real prices...
        </div>
      </div>
    );
  }

  // Error state — the lookup itself failed
  if (status === "error") {
    return (
      <div style={S.card}>
        <div style={{ fontWeight: 600 }}>{item.description}</div>
        <div style={S.muted}>
          Part {item.partNumber || "—"} · Qty {item.quantity} · Current {fmt(item.unitPrice)}
        </div>
        <div style={{ color: "#dc2626", marginTop: 10, display: "flex", alignItems: "center", gap: 6 }}>
          <AlertCircle size={14} /> Lookup failed: {error || "unknown error"}
        </div>
      </div>
    );
  }

  const best = offers.find((o) => o.found);
  const savingsPerUnit = best ? item.unitPrice - best.unitPrice : 0;
  const totalSavings = savingsPerUnit * item.quantity;
  const savingsPct = best && item.unitPrice > 0 ? (savingsPerUnit / item.unitPrice) * 100 : 0;

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
          {best ? (
            <>
              <div style={{ color: totalSavings > 0 ? "#16a34a" : "#dc2626", fontWeight: 700 }}>
                {totalSavings > 0 ? "Save " : ""}
                {fmt(totalSavings)}
              </div>
              <div style={S.muted}>{savingsPct.toFixed(1)}% per unit</div>
            </>
          ) : (
            <div style={{ color: "#b45309", fontWeight: 600, fontSize: 13 }}>
              No matches found — investigate manually
            </div>
          )}
        </div>
      </div>

      <div style={S.offerGrid}>
        {offers.map((o, i) => {
          const isBest = best && o.supplier === best.supplier;
          return (
            <div
              key={o.supplier}
              style={{
                ...S.offer,
                ...(isBest ? S.offerBest : {}),
                opacity: o.found ? 1 : 0.7,
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 4 }}>
                <div style={{ fontWeight: 600, display: "flex", alignItems: "center", gap: 4 }}>
                  {o.supplier}
                  {o.preferred && (
                    <span title="Your preferred supplier" style={{ fontSize: 10, color: "#7c3aed" }}>
                      ★
                    </span>
                  )}
                </div>
                {isBest && (
                  <span style={S.bestTag}>
                    <TrendingDown size={12} /> Best
                  </span>
                )}
              </div>

              {o.found ? (
                <>
                  <div style={{ fontSize: 22, fontWeight: 700, marginTop: 4 }}>{fmt(o.unitPrice)}</div>
                  <div style={S.muted}>per unit · total {fmt(o.total)}</div>
                  <div style={{ marginTop: 6 }}>
                    <ConfidenceBadge level={o.confidence} />
                  </div>
                  <div style={{ ...S.muted, fontSize: 12, marginTop: 6, lineHeight: 1.35 }}>
                    {o.matchedDescription}
                    {o.matchedSku ? <> · SKU <span style={{ fontFamily: "monospace" }}>{o.matchedSku}</span></> : null}
                  </div>
                  {o.matchNotes && (
                    <div style={{ ...S.muted, fontSize: 11, fontStyle: "italic", marginTop: 4 }}>
                      {o.matchNotes}
                    </div>
                  )}
                  {o.url && (
                    <a href={o.url} target="_blank" rel="noreferrer" style={S.searchLink}>
                      <ExternalLink size={12} /> View on {o.supplier}
                    </a>
                  )}
                </>
              ) : (
                <>
                  <div style={{ fontSize: 13, marginTop: 8, color: "#64748b" }}>
                    No clear match{o.reason ? `: ${o.reason}` : ""}
                  </div>
                  {o.fallbackUrl && (
                    <a href={o.fallbackUrl} target="_blank" rel="noreferrer" style={S.searchLink}>
                      <Search size={12} /> Search {o.supplier} manually
                    </a>
                  )}
                </>
              )}
            </div>
          );
        })}
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
              const b = (r.offers || []).find((o) => o.found);
              const sper = b ? r.item.unitPrice - b.unitPrice : 0;
              return (
                <tr key={r.item.id}>
                  <td style={S.td}>{r.item.id}</td>
                  <td style={S.td}>{r.item.description}</td>
                  <td style={S.td}>{r.item.partNumber}</td>
                  <td style={S.td}>{r.item.quantity}</td>
                  <td style={S.td}>{fmt(r.item.unitPrice)}</td>
                  <td style={S.td}>{b ? b.supplier : "—"}</td>
                  <td style={S.td}>{b ? fmt(b.unitPrice) : "no match"}</td>
                  <td style={S.td}>{b ? fmt(sper) : "—"}</td>
                  <td style={S.td}>{b ? fmt(sper * r.item.quantity) : "—"}</td>
                  <td style={S.td}>{b ? `${b.deliveryDays} days` : "—"}</td>
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
  gateBg: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: "#f8fafc",
    padding: 16,
  },
  gateCard: {
    width: "100%",
    maxWidth: 380,
    background: "#fff",
    border: "1px solid #e2e8f0",
    borderRadius: 12,
    padding: 28,
    boxShadow: "0 4px 20px rgba(15,23,42,0.05)",
  },
  gateTitle: { margin: 0, fontSize: 20, fontWeight: 700, color: "#0f172a" },
  gateSub: { margin: "6px 0 20px", fontSize: 13, color: "#64748b" },
  gateInput: {
    width: "100%",
    padding: "10px 12px",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    fontSize: 15,
    fontFamily: "inherit",
    marginBottom: 12,
    boxSizing: "border-box",
  },
  gateBtn: {
    width: "100%",
    padding: "10px 12px",
    border: "none",
    borderRadius: 8,
    background: "#4f46e5",
    color: "#fff",
    fontWeight: 600,
    fontSize: 15,
    cursor: "pointer",
  },
  gateBtnDisabled: {
    width: "100%",
    padding: "10px 12px",
    border: "none",
    borderRadius: 8,
    background: "#cbd5e1",
    color: "#fff",
    fontWeight: 600,
    fontSize: 15,
    cursor: "not-allowed",
  },
  gateError: { color: "#b91c1c", fontSize: 13, marginTop: 10, marginBottom: 0 },
};

// ---------- Password gate (wraps the app) ----------
function PasswordGate() {
  const [pw, setPw] = useState("");
  const [authed, setAuthed] = useState(null); // null = unknown, false = bad, true = good
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");

  const submit = async (e) => {
    e.preventDefault();
    if (!pw) return;
    setSubmitting(true);
    setError("");
    try {
      const res = await fetch("/api/ai", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password: pw, action: "verify" }),
      });
      if (res.ok) {
        setAuthed(true);
      } else {
        const data = await res.json().catch(() => ({}));
        setAuthed(false);
        setError(data?.error || "Incorrect password");
      }
    } catch (err) {
      setError("Could not reach the server. Try again.");
    } finally {
      setSubmitting(false);
    }
  };

  if (authed === true) return <MROApp password={pw} />;

  return (
    <div style={S.gateBg}>
      <form style={S.gateCard} onSubmit={submit}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
          <div style={{ ...S.logoMark, width: 32, height: 32, borderRadius: 8 }}>
            <Sparkles size={18} color="#fff" />
          </div>
          <h1 style={S.gateTitle}>MRO Price Scout</h1>
        </div>
        <p style={S.gateSub}>Enter the access password to continue.</p>
        <input
          type="password"
          autoFocus
          value={pw}
          onChange={(e) => setPw(e.target.value)}
          placeholder="Password"
          style={S.gateInput}
          disabled={submitting}
        />
        <button
          type="submit"
          style={submitting || !pw ? S.gateBtnDisabled : S.gateBtn}
          disabled={submitting || !pw}
        >
          {submitting ? "Checking..." : "Unlock"}
        </button>
        {error && <p style={S.gateError}>{error}</p>}
      </form>
    </div>
  );
}

export default PasswordGate;
