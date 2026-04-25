// Vercel serverless function — proxies AI calls to Anthropic.
// The Anthropic API key lives ONLY in this server function's env vars
// (ANTHROPIC_API_KEY) and is never sent to the browser.
// Access is gated by a shared password (SCOUT_ACCESS_PASSWORD).

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

  const { password, action, prompt, maxTokens } = req.body || {};

  if (password !== expectedPassword) {
    return res.status(401).json({ error: "Invalid password" });
  }

  // Action "verify" lets the front-end check the password without spending tokens.
  if (action === "verify") {
    return res.status(200).json({ ok: true });
  }

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
        model: "claude-haiku-4-5-20251001",
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
