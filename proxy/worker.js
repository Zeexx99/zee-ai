/**
 * Zee AI team proxy – Cloudflare Worker (free tier is plenty).
 * Keeps the Gemini API key on the server; add-in users only get a team token.
 *
 * Secrets (wrangler secret put …):
 *   GEMINI_API_KEY  – your Google AI Studio key
 *   TEAM_TOKEN      – one shared token, and/or
 *   TEAM_TOKENS     – one token per person: "ziad:longrandom1,omar:longrandom2" (lets you cap and revoke per person)
 * Variables (wrangler.toml [vars] or the dashboard):
 *   ALLOWED_ORIGIN  – your add-in origin, e.g. https://<user>.github.io (comma-separate several).
 *                     When set, requests from any other origin are refused. Strongly recommended.
 *   DAILY_LIMIT     – max requests per token per day (UTC). Needs the USAGE KV binding. 0/empty = no cap.
 *   MAX_BODY_MB     – largest request body accepted (default 20, Gemini's inline limit).
 *   ALLOWED_MODELS  – optional comma list of model IDs (e.g. gemini-3.8-flash,gemini-3.5-flash-lite).
 * KV binding (optional, for DAILY_LIMIT):
 *   USAGE           – a KV namespace (wrangler kv namespace create USAGE).
 */
const UPSTREAM = "https://generativelanguage.googleapis.com";
const PATH = /^\/v1beta\/models\/([\w.\-]+):(generateContent|streamGenerateContent|countTokens)$/;

const list = (v) =>
  String(v || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);

/** Map of token -> user name from TEAM_TOKEN / TEAM_TOKENS. */
export function tokenTable(env) {
  const table = new Map();
  if (env.TEAM_TOKEN) table.set(String(env.TEAM_TOKEN), "team");
  for (const entry of list(env.TEAM_TOKENS)) {
    const i = entry.indexOf(":");
    if (i > 0) table.set(entry.slice(i + 1), entry.slice(0, i));
    else table.set(entry, entry.slice(0, 6));
  }
  return table;
}

/** Constant-time string comparison (avoids leaking the token through timing). */
function safeEqual(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function findUser(table, auth) {
  const m = /^Bearer\s+(.+)$/.exec(auth || "");
  if (!m) return null;
  for (const [token, user] of table) if (safeEqual(token, m[1].trim())) return user;
  return null;
}

function json(status, message, headers) {
  return new Response(JSON.stringify({ error: { code: status, message } }), { status, headers: { ...headers, "Content-Type": "application/json" } });
}

export default {
  async fetch(request, env) {
    const origins = list(env.ALLOWED_ORIGIN);
    const origin = request.headers.get("Origin") || "";
    const originOk = !origins.length || origins.includes(origin);
    const cors = {
      "Access-Control-Allow-Origin": origins.length ? (originOk ? origin : origins[0]) : "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Max-Age": "86400",
      Vary: "Origin",
    };
    if (request.method === "OPTIONS") return new Response(null, { status: originOk ? 204 : 403, headers: cors });
    if (!originOk) return json(403, "Origin not allowed", cors);
    if (request.method !== "POST") return json(405, "Use POST", cors);

    const url = new URL(request.url);
    const route = PATH.exec(url.pathname);
    if (!route) return json(404, "Not allowed", cors);
    const allowedModels = list(env.ALLOWED_MODELS);
    if (allowedModels.length && !allowedModels.includes(route[1])) return json(403, `Model ${route[1]} is not enabled on this proxy`, cors);

    const table = tokenTable(env);
    const user = table.size ? findUser(table, request.headers.get("Authorization")) : null;
    if (!user) return json(401, "Invalid team token", cors);

    const maxBytes = (Number(env.MAX_BODY_MB) > 0 ? Number(env.MAX_BODY_MB) : 20) * 1024 * 1024;
    const length = Number(request.headers.get("Content-Length") || "0");
    if (length > maxBytes) return json(413, `Request too large (max ${Math.round(maxBytes / 1048576)} MB)`, cors);

    const body = await request.arrayBuffer();
    if (body.byteLength > maxBytes) return json(413, `Request too large (max ${Math.round(maxBytes / 1048576)} MB)`, cors);

    // Daily cap per token (soft: KV is eventually consistent, good enough to stop runaway use).
    const limit = Number(env.DAILY_LIMIT) || 0;
    if (limit > 0 && env.USAGE && route[2] !== "countTokens") {
      const key = `day:${new Date().toISOString().slice(0, 10)}:${user}`;
      const used = Number((await env.USAGE.get(key)) || "0");
      if (used >= limit) return json(429, `Daily limit of ${limit} requests reached for "${user}". It resets at 00:00 UTC.`, cors);
      await env.USAGE.put(key, String(used + 1), { expirationTtl: 60 * 60 * 48 });
    }


    const upstream = new URL(url.pathname + url.search, UPSTREAM);
    const res = await fetch(upstream, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
      body,
    });
    const headers = new Headers(res.headers);
    Object.entries(cors).forEach(([k, v]) => headers.set(k, v));
    return new Response(res.body, { status: res.status, headers });
  },
};
