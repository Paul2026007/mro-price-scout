// Vercel serverless function — proxies AI calls to Anthropic.
// The Anthropic API key lives ONLY in this server function's env vars
// (ANTHROPIC_API_KEY) and is never sent to the browser.
// Access is gated by a shared password (SCOUT_ACCESS_PASSWORD).
//
// Supported actions:
//   verify      — { password } → { ok: true } (no token spend)
//   (default)   — { password, prompt, maxTokens } → { text }
//   priceLookup — { password, item, suppliers } → { offers: [...] }
//                 Uses Anthropic web_search tool to find real listings.

const TEXT_MODEL = "claude-haiku-4-5-20251001";
// Use Haiku for lookups too — higher rate limit, faster per-request, lets us
// run multiple lookups in parallel without hitting Anthropic's per-minute cap.
const SEARCH_MODEL = "claude-haiku-4-5-20251001";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const expectedPassword = process.env.SCOUT_ACCESS_PASSWORD;
  const anthropicKey = process.env.ANTHROPIC_API_KEY;

  if (!expectedPassword || !anthropicKey) {
    return res.status(500).json({
      error:
        "Server is missing required env vars. Set ANTHROPIC_API_KEY and SCOUT_ACCESS_PASSWORD in Vercel.",
    });
  }

  const { password, action, prompt, maxTokens, item, suppliers } = req.body || {};

  if (password !== expectedPassword) {
    return res.status(401).json({ error: "Invalid password" });
  }

  if (action === "verify") {
    return res.status(200).json({ ok: true });
  }

  if (action === "priceLookup") {
    return handlePriceLookup(res, anthropicKey, item, suppliers);
  }

  // Default: plain text generation
  if (!prompt || typeof prompt !== "string") {
    return res.status(400).json({ error: "Missing prompt" });
  }

  try {
    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": anthropicKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: TEXT_MODEL,
        max_tokens: typeof maxTokens === "number" ? maxTokens : 600,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    const data = await upstream.json();
    if (!upstream.ok) {
      return res.status(upstream.status).json({
        error: data?.error?.message || "Anthropic upstream error",
      });
    }
    const text = data?.content?.[0]?.text || "";
    return res.status(200).json({ text });
  } catch (err) {
    return res.status(500).json({ error: err.message || "Request failed" });
  }
}

// ---------- Price lookup via web search ----------
async function handlePriceLookup(res, anthropicKey, item, suppliers) {
  if (!item || !item.description) {
    return res.status(400).json({ error: "Missing item.description" });
  }
  const list = (suppliers || []).map((s) => `- ${s.name} (search at ${s.domain})`);
  if (list.length === 0) {
    return res.status(400).json({ error: "No suppliers specified" });
  }

  const lookupPrompt = `You are a procurement researcher. Find current online list prices for this MRO part on the suppliers listed below. Use real web searches — do not guess prices or invent SKUs.

ITEM TO PRICE:
- Description: ${item.description}
- Part number provided by buyer: ${item.partNumber || "(none)"}
- Quantity needed: ${item.quantity}

SUPPLIERS TO CHECK:
${list.join("\n")}

For each supplier, search their website and find the closest matching product. Look at the actual product page when possible. Note the supplier's SKU, product title, and the listed unit price (single-unit price unless quantity tier pricing is shown for the buyer's quantity).

If you cannot find a credible match for a supplier (no product found, or only loosely related items), mark found: false for that supplier. Do not invent a price.

Respond with ONLY a JSON array — no preamble, no markdown fences, no explanation. Each element is one supplier:

[
  {
    "supplier": "<exact supplier name from the list above>",
    "found": true,
    "matchedDescription": "<product title from the listing>",
    "matchedSku": "<supplier SKU/item number, or empty string if not visible>",
    "unitPrice": <number, USD>,
    "url": "<direct URL to the product page>",
    "confidence": "high|medium|low",
    "matchNotes": "<one short phrase such as 'exact part number match', 'manufacturer cross-reference', 'compatible spec match', 'similar product, verify'>"
  }
]

For not-found use: { "supplier": "...", "found": false, "reason": "<short reason>" }

Confidence guide:
  high   — exact part number or SKU match on the supplier's site
  medium — same brand + same spec/size, different SKU
  low    — similar product but specs not fully verified

Return one entry per supplier in the SUPPLIERS TO CHECK list.`;

  try {
    const requestBody = JSON.stringify({
      model: SEARCH_MODEL,
      max_tokens: 2500,
      tools: [
        {
          type: "web_search_20250305",
          name: "web_search",
          max_uses: 8,
        },
      ],
      messages: [{ role: "user", content: lookupPrompt }],
    });

    const callAnthropic = () =>
      fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": anthropicKey,
          "anthropic-version": "2023-06-01",
        },
        body: requestBody,
      });

    // Retry with backoff on 429 (rate limit). Anthropic enforces 30k input
    // tokens/min on Sonnet, so a brief wait usually clears the budget.
    let upstream = await callAnthropic();
    let attempt = 0;
    while (upstream.status === 429 && attempt < 3) {
      attempt += 1;
      const retryAfterHeader = upstream.headers.get("retry-after");
      const retryAfterSec = retryAfterHeader ? parseInt(retryAfterHeader, 10) : NaN;
      const waitMs = Number.isFinite(retryAfterSec)
        ? Math.min(retryAfterSec * 1000, 45000)
        : Math.min(15000 * attempt, 45000);
      await new Promise((r) => setTimeout(r, waitMs));
      upstream = await callAnthropic();
    }

    const data = await upstream.json();
    if (!upstream.ok) {
      return res.status(upstream.status).json({
        error: data?.error?.message || "Anthropic upstream error",
      });
    }

    // Concatenate text blocks from the response (web_search adds search blocks too)
    const textBlocks = (data?.content || [])
      .filter((b) => b.type === "text")
      .map((b) => b.text);
    const fullText = textBlocks.join("\n").trim();

    // Extract JSON array (Claude sometimes wraps it in fences or prose)
    let offers = [];
    try {
      const start = fullText.indexOf("[");
      const end = fullText.lastIndexOf("]");
      if (start >= 0 && end > start) {
        offers = JSON.parse(fullText.slice(start, end + 1));
      } else {
        throw new Error("No JSON array found in response");
      }
    } catch (parseErr) {
      return res.status(500).json({
        error: "Could not parse supplier results",
        raw: fullText.slice(0, 500),
      });
    }

    return res.status(200).json({ offers });
  } catch (err) {
    return res.status(500).json({ error: err.message || "Request failed" });
  }
}
