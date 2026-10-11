import { useEffect, useLayoutEffect, useState } from "react";
import { createDarkTheme, createLightTheme, type BrandVariants, type Theme } from "@fluentui/react-components";

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

export const lightTheme: Theme = createLightTheme(zeeBrand);
export const darkTheme: Theme = { ...createDarkTheme(zeeBrand), colorBrandForeground1: zeeBrand[120], colorBrandForeground2: zeeBrand[130] };

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

/** Applies the theme class to <body> synchronously so portals (dialogs, drawer, toasts) match and nothing flashes. */
export function useBodyThemeClass(dark: boolean) {
  useLayoutEffect(() => {
    document.body.classList.toggle("zee-dark", dark);
    document.body.classList.toggle("zee-light", !dark);
  }, [dark]);
}
