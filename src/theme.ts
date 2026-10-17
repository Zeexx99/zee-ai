import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { createDarkTheme, createLightTheme, teamsHighContrastTheme, type BrandVariants, type Theme } from "@fluentui/react-components";
import type { AccentColor, Density } from "./services/settingsService";

type Shade = keyof BrandVariants;
const SHADES: Shade[] = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160];

/** Zee AI brand ramp (violet → used for primary buttons, focus, links). */
const zeeBrand: BrandVariants = {
  10: "#070417",
  20: "#160F3A",
  30: "#221663",
  40: "#2C1C82",
  50: "#3622A0",
  60: "#4129BC",
  70: "#4D33D3",
  80: "#5B42E0",
  90: "#6B55E8",
  100: "#7B68EE",
  110: "#8C7CF2",
  120: "#9D90F5",
  130: "#AEA4F8",
  140: "#C0B8FA",
  150: "#D2CDFC",
  160: "#E5E2FE",
};

function mix(hex: string, target: number, amount: number): string {
  const h = hex.replace("#", "");
  const ch = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
  return `#${ch.map((c) => Math.round(c + (target - c) * amount).toString(16).padStart(2, "0")).join("")}`.toUpperCase();
}

/**
 * Builds a 16-step Fluent brand ramp around one colour (shade 80 = the colour itself).
 * Darker shades mix towards black, lighter ones towards white. Pure – unit-tested.
 */
export function rampFrom(hex: string): BrandVariants {
  const out = {} as BrandVariants;
  SHADES.forEach((s, i) => {
    const idx = i - 7; // shade 80 sits at index 7
    out[s] = idx < 0 ? mix(hex, 0, (-idx / 7) * 0.92) : idx === 0 ? hex.toUpperCase() : mix(hex, 255, (idx / 8) * 0.86);
  });
  return out;
}

const RAMPS: Record<AccentColor, BrandVariants> = {
  violet: zeeBrand,
  teal: rampFrom("#00897B"),
  blue: rampFrom("#1F6FD1"),
  orange: rampFrom("#D9640F"),
};

export function makeTheme(dark: boolean, accent: AccentColor = "violet"): Theme {
  const brand = RAMPS[accent] ?? zeeBrand;
  return dark ? { ...createDarkTheme(brand), colorBrandForeground1: brand[120], colorBrandForeground2: brand[130] } : createLightTheme(brand);
}

export const lightTheme: Theme = makeTheme(false);
export const darkTheme: Theme = makeTheme(true);
export const highContrastTheme: Theme = teamsHighContrastTheme;

/** Follow the Office theme (falls back to the OS colour scheme). */
export function prefersDark(): boolean {
  try {
    const bg = Office?.context?.officeTheme?.bodyBackgroundColor;
    if (bg && /^#?[0-9a-f]{6}$/i.test(bg)) {
      const hex = bg.replace("#", "");
      const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
      return 0.299 * r + 0.587 * g + 0.114 * b < 128;
    }
  } catch {
    /* not in Office */
  }
  return typeof window !== "undefined" && window.matchMedia?.("(prefers-color-scheme: dark)").matches;
}

/**
 * Tracks the "automatic" colour scheme: Office theme first, OS setting as fallback.
 * Excel doesn't reliably fire a theme-changed event, so we also re-check when the pane regains
 * focus or the OS scheme flips.
 */
export function useAutoDark(): boolean {
  const [dark, setDark] = useState(prefersDark);
  useEffect(() => {
    const check = () => setDark(prefersDark());
    const mq = window.matchMedia?.("(prefers-color-scheme: dark)");
    mq?.addEventListener?.("change", check);
    window.addEventListener("focus", check);
    document.addEventListener("visibilitychange", check);
    let officeHandler = false;
    try {
      // Supported by some Office hosts/versions only; harmless elsewhere.
      Office.context.document.addHandlerAsync(Office.EventType.OfficeThemeChanged, check, (r) => {
        officeHandler = r.status === Office.AsyncResultStatus.Succeeded;
      });
    } catch {
      /* not supported */
    }
    return () => {
      mq?.removeEventListener?.("change", check);
      window.removeEventListener("focus", check);
      document.removeEventListener("visibilitychange", check);
      if (officeHandler) {
        try {
          Office.context.document.removeHandlerAsync(Office.EventType.OfficeThemeChanged, { handler: check });
        } catch {
          /* ignore */
        }
      }
    };
  }, []);
  return dark;
}

/** True while Windows high-contrast mode (forced colours) is on. */
export function useHighContrast(): boolean {
  const query = "(forced-colors: active)";
  const [hc, setHc] = useState(() => typeof window !== "undefined" && !!window.matchMedia?.(query).matches);
  useEffect(() => {
    const mq = window.matchMedia?.(query);
    const check = () => setHc(!!mq?.matches);
    mq?.addEventListener?.("change", check);
    return () => mq?.removeEventListener?.("change", check);
  }, []);
  return hc;
}

export interface BodyLook {
  dark: boolean;
  accent: AccentColor;
  density: Density;
  highContrast: boolean;
}

/**
 * Applies theme / accent / density classes to <body> synchronously so portals (dialogs, drawer,
 * toasts) match and nothing flashes. Theme switches after the first paint fade briefly.
 */
export function useBodyThemeClass(look: BodyLook | boolean) {
  const l: BodyLook = typeof look === "boolean" ? { dark: look, accent: "violet", density: "comfortable", highContrast: false } : look;
  const first = useRef(true);
  useLayoutEffect(() => {
    const b = document.body.classList;
    if (!first.current) {
      b.add("zee-theme-anim");
      window.setTimeout(() => b.remove("zee-theme-anim"), 350);
    }
    first.current = false;
    b.toggle("zee-dark", l.dark);
    b.toggle("zee-light", !l.dark);
    for (const a of ["violet", "teal", "blue", "orange"]) b.toggle(`zee-accent-${a}`, l.accent === a);
    b.toggle("zee-compact", l.density === "compact");
    b.toggle("zee-hc", l.highContrast);
  }, [l.dark, l.accent, l.density, l.highContrast]);
}
