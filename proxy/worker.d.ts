/** Types for the Cloudflare Worker (used by the unit tests). */
export interface ProxyEnv {
  GEMINI_API_KEY?: string;
  TEAM_TOKEN?: string;
  TEAM_TOKENS?: string;
  ALLOWED_ORIGIN?: string;
  DAILY_LIMIT?: string;
  MAX_BODY_MB?: string;
  ALLOWED_MODELS?: string;
  USAGE?: { get(key: string): Promise<string | null>; put(key: string, value: string, opts?: { expirationTtl?: number }): Promise<void> };
}

export function tokenTable(env: ProxyEnv): Map<string, string>;

declare const worker: { fetch(request: Request, env: ProxyEnv): Promise<Response> };
export default worker;
