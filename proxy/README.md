# Zee AI team proxy (optional)

Use this when several people use Zee AI and you don't want a Gemini API key on every PC.

1. Install Wrangler: `npm i -g wrangler` and run `wrangler login`.
2. In this folder: `wrangler init zee-proxy --yes` (or create a Worker in the Cloudflare dashboard) and use `worker.js` as the code.
3. Add the secrets:
   ```
   wrangler secret put GEMINI_API_KEY
   wrangler secret put TEAM_TOKEN
   ```
   Optionally set the variable `ALLOWED_ORIGIN` to your add-in origin, e.g. `https://<user>.github.io`.
4. `wrangler deploy` → you get a URL like `https://zee-proxy.<account>.workers.dev`.
5. In Zee AI → Settings → **Team proxy**, paste that URL and the team token.

The proxy only forwards `generateContent`, `streamGenerateContent` and `countTokens` calls, and it refuses requests without the token.
