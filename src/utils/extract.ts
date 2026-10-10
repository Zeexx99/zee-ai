/** Pure, local (no-AI) extractors for the "Split / Extract" quick action. */

export type ExtractKind = "firstName" | "lastName" | "email" | "emailDomain" | "phone" | "number";

export const EXTRACT_LABELS: Record<ExtractKind, string> = {
  firstName: "First name",
  lastName: "Last name",
  email: "Email address",
  emailDomain: "Email domain",
  phone: "Phone number",
  number: "First number",
};

const TITLES = /^(mr|mrs|ms|miss|mx|dr|prof|sir|eng|engr)\.?$/i;
const SUFFIXES = /^(jr|sr|ii|iii|iv|phd|md|esq)\.?$/i;
const EMAIL = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
const PHONE = /(?:\+?\d[\d\s().\-]{5,}\d)/;
const NUMBER = /[-+]?\d[\d,]*(?:\.\d+)?/;

function nameTokens(s: string): string[] {
  return s
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .filter((t) => t && !TITLES.test(t) && !SUFFIXES.test(t.replace(/,$/, "")));
}

export function extractValue(input: unknown, kind: ExtractKind): string {
  if (input === null || input === undefined) return "";
  const s = String(input).trim();
  if (!s) return "";
  switch (kind) {
    case "firstName": {
      if (s.includes(",")) {
        const after = nameTokens(s.slice(s.indexOf(",") + 1));
        if (after.length) return after[0];
      }
      return nameTokens(s)[0] ?? "";
    }
    case "lastName": {
      if (s.includes(",")) return nameTokens(s.slice(0, s.indexOf(","))).join(" ");
      const t = nameTokens(s);
      return t.length > 1 ? t[t.length - 1] : "";
    }
    case "email":
      return EMAIL.exec(s)?.[0] ?? "";
    case "emailDomain": {
      const e = EMAIL.exec(s)?.[0];
      return e ? e.slice(e.indexOf("@") + 1).toLowerCase() : "";
    }
    case "phone":
      return PHONE.exec(s)?.[0].trim() ?? "";
    case "number":
      return NUMBER.exec(s)?.[0].replace(/,/g, "") ?? "";
  }
}
