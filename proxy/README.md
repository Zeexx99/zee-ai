# Zee AI team proxy (optional)

Use this when several people use Zee AI and you don't want a Gemini API key on every PC.

1. Install Wrangler: `npm i -g wrangler` and run `wrangler login`.
2. In this folder: `wrangler init zee-proxy --yes` (or create a Worker in the Cloudflare dashboard) and use `worker.js` as the code.
3. Add the secrets:
   ```
   wrangler secret put GEMINI_API_KEY
   wrangler secret put TEAM_TOKENS     # e.g.  ziad:3f9c…long-random…,omar:a71b…long-random…
   ```
   (`TEAM_TOKEN` with a single shared token still works, but per-person tokens let you cap and revoke one person.)
4. Set the variables (Cloudflare dashboard → Worker → Settings → Variables, or `[vars]` in `wrangler.toml`):
   - `ALLOWED_ORIGIN` = your add-in origin, e.g. `https://<user>.github.io` – **recommended**: other websites are then refused.
   - `DAILY_LIMIT` = e.g. `500` requests per person per day (needs step 5).
   - `MAX_BODY_MB` = largest request (default `20`; attachments count).
   - `ALLOWED_MODELS` = optional, e.g. `gemini-3.8-flash,gemini-3.5-flash-lite` to stop expensive models.
5. For `DAILY_LIMIT`, create a KV namespace and bind it as `USAGE`:
   ```
   wrangler kv namespace create USAGE
   ```
   and add the printed binding to `wrangler.toml` (`[[kv_namespaces]] binding = "USAGE" id = "…"`).
6. `wrangler deploy` → you get a URL like `https://zee-proxy.<account>.workers.dev`.
7. In Zee AI → Settings → **Team proxy**, paste that URL and the person's token.

The proxy only forwards `generateContent`, `streamGenerateContent` and `countTokens` POST calls, refuses requests without a valid token, from other origins (when `ALLOWED_ORIGIN` is set), over the size limit, or over the daily cap. To revoke someone, remove their entry from `TEAM_TOKENS`.
