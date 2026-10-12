import { afterEach, describe, expect, it } from "vitest";
import worker, { tokenTable } from "../../proxy/worker.js";

const URL_OK = "https://proxy.example/v1beta/models/gemini-3.8-flash:generateContent";
const realFetch = globalThis.fetch;
let upstreamCalls = 0;

function stubUpstream() {
  upstreamCalls = 0;
  globalThis.fetch = (async () => {
    upstreamCalls++;
    return new Response('{"candidates":[]}', { status: 200, headers: { "Content-Type": "application/json" } });
  }) as typeof fetch;
}

function req(opts: { url?: string; token?: string; origin?: string; method?: string; body?: string } = {}) {
  const headers: Record<string, string> = { "Content-Type": "application/json" };
  if (opts.token) headers.Authorization = `Bearer ${opts.token}`;
  if (opts.origin) headers.Origin = opts.origin;
  const method = opts.method ?? "POST";
  return new Request(opts.url ?? URL_OK, { method, headers, body: method === "POST" ? (opts.body ?? "{}") : undefined });
}

class MemoryKv {
  data = new Map<string, string>();
  async get(k: string) {
    return this.data.get(k) ?? null;
  }
  async put(k: string, v: string) {
    this.data.set(k, v);
  }
}

afterEach(() => {
  globalThis.fetch = realFetch;
});

describe("team proxy", () => {
  it("reads shared and per-person tokens", () => {
    const t = tokenTable({ TEAM_TOKEN: "shared", TEAM_TOKENS: "ziad:tok1, omar:tok2" });
    expect(t.get("shared")).toBe("team");
    expect(t.get("tok1")).toBe("ziad");
    expect(t.get("tok2")).toBe("omar");
  });

  it("forwards valid requests and refuses bad tokens", async () => {
    stubUpstream();
    const env = { GEMINI_API_KEY: "k", TEAM_TOKENS: "ziad:tok1" };
    expect((await worker.fetch(req({ token: "tok1" }), env)).status).toBe(200);
    expect((await worker.fetch(req({ token: "nope" }), env)).status).toBe(401);
    expect((await worker.fetch(req({}), env)).status).toBe(401);
    expect(upstreamCalls).toBe(1);
  });

  it("only allows model calls over POST", async () => {
    stubUpstream();
    const env = { GEMINI_API_KEY: "k", TEAM_TOKEN: "t" };
    expect((await worker.fetch(req({ token: "t", url: "https://proxy.example/v1beta/models" }), env)).status).toBe(404);
    expect((await worker.fetch(req({ token: "t", method: "GET" }), env)).status).toBe(405);
  });

  it("enforces the allowed origin", async () => {
    stubUpstream();
    const env = { GEMINI_API_KEY: "k", TEAM_TOKEN: "t", ALLOWED_ORIGIN: "https://zee.example" };
    expect((await worker.fetch(req({ token: "t", origin: "https://evil.example" }), env)).status).toBe(403);
    const ok = await worker.fetch(req({ token: "t", origin: "https://zee.example" }), env);
    expect(ok.status).toBe(200);
    expect(ok.headers.get("Access-Control-Allow-Origin")).toBe("https://zee.example");
  });

  it("caps request size, models and daily use", async () => {
    stubUpstream();
    const kv = new MemoryKv();
    const env = { GEMINI_API_KEY: "k", TEAM_TOKEN: "t", MAX_BODY_MB: "0.001", DAILY_LIMIT: "2", USAGE: kv, ALLOWED_MODELS: "gemini-3.8-flash" };
    expect((await worker.fetch(req({ token: "t", body: "x".repeat(5000) }), env)).status).toBe(413);
    expect((await worker.fetch(req({ token: "t", url: "https://proxy.example/v1beta/models/gemini-3.1-pro-preview:generateContent" }), env)).status).toBe(403);
    expect((await worker.fetch(req({ token: "t" }), env)).status).toBe(200);
    expect((await worker.fetch(req({ token: "t" }), env)).status).toBe(200);
    const third = await worker.fetch(req({ token: "t" }), env);
    expect(third.status).toBe(429);
    expect(upstreamCalls).toBe(2);
  });
});
