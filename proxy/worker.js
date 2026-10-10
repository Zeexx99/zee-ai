/**
 * Zee AI team proxy – Cloudflare Worker (free tier is plenty).
 * Keeps the Gemini API key on the server; add-in users only get a shared team token.
 *
 * Secrets (wrangler secret put …):
 *   GEMINI_API_KEY  – your Google AI Studio key
 *   TEAM_TOKEN      – any long random string you give to your team (Settings → Team proxy)
 * Optional variable:
 *   ALLOWED_ORIGIN  – e.g. https://<user>.github.io (defaults to *)
 */
const UPSTREAM = "https://generativelanguage.googleapis.com";

export default {
  async fetch(request, env) {
    const origin = env.ALLOWED_ORIGIN || "*";
    const cors = {
      "Access-Control-Allow-Origin": origin,
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Max-Age": "86400",
    };
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });

    const url = new URL(request.url);
    // Only allow model calls: /v1beta/models/<model>:<method>
    if (!/^\/v1beta\/models\/[\w.\-]+:(generateContent|streamGenerateContent|countTokens)$/.test(url.pathname)) {
      return new Response(JSON.stringify({ error: { code: 404, message: "Not allowed" } }), { status: 404, headers: { ...cors, "Content-Type": "application/json" } });
    }
    const auth = request.headers.get("Authorization") || "";
    if (!env.TEAM_TOKEN || auth !== `Bearer ${env.TEAM_TOKEN}`) {
      return new Response(JSON.stringify({ error: { code: 401, message: "Invalid team token" } }), { status: 401, headers: { ...cors, "Content-Type": "application/json" } });
    }

    const upstream = new URL(url.pathname + url.search, UPSTREAM);
    const res = await fetch(upstream, {
      method: request.method,
      headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
      body: request.method === "POST" ? request.body : undefined,
    });
    const headers = new Headers(res.headers);
    Object.entries(cors).forEach(([k, v]) => headers.set(k, v));
    return new Response(res.body, { status: res.status, headers });
  },
};
