/**
 * Free-form unit conversion for =ZEE.UNITS(value, "from", "to").
 * Pure, local, instant – no AI call. Factors convert to an SI-ish base unit per dimension.
 */

type Dim = "length" | "pressure" | "volume" | "rate" | "power" | "mass" | "density" | "viscosity" | "time" | "temperature";

interface UnitDef {
  dim: Dim;
  /** Multiply by this to get the base unit (ignored for temperature / API). */
  factor: number;
}

const BBL = 0.158987294928; // m³
const FT3 = 0.028316846592; // m³
const GAL = 0.003785411784; // m³

const UNITS: Record<string, UnitDef> = {
  // length → m
  m: { dim: "length", factor: 1 },
  cm: { dim: "length", factor: 0.01 },
  mm: { dim: "length", factor: 0.001 },
  km: { dim: "length", factor: 1000 },
  ft: { dim: "length", factor: 0.3048 },
  in: { dim: "length", factor: 0.0254 },
  mi: { dim: "length", factor: 1609.344 },
  // pressure → Pa (gauge/absolute suffixes share the factor)
  pa: { dim: "pressure", factor: 1 },
  kpa: { dim: "pressure", factor: 1000 },
  mpa: { dim: "pressure", factor: 1e6 },
  bar: { dim: "pressure", factor: 1e5 },
  psi: { dim: "pressure", factor: 6894.757293168 },
  atm: { dim: "pressure", factor: 101325 },
  "kg/cm2": { dim: "pressure", factor: 98066.5 },
  // volume → m³ (gas standard volumes use the same factor)
  m3: { dim: "volume", factor: 1 },
  l: { dim: "volume", factor: 0.001 },
  bbl: { dim: "volume", factor: BBL },
  gal: { dim: "volume", factor: GAL },
  ft3: { dim: "volume", factor: FT3 },
  mscf: { dim: "volume", factor: FT3 * 1e3 },
  mmscf: { dim: "volume", factor: FT3 * 1e6 },
  // rate → m³/d
  "m3/d": { dim: "rate", factor: 1 },
  "m3/h": { dim: "rate", factor: 24 },
  "l/s": { dim: "rate", factor: 86.4 },
  "bbl/d": { dim: "rate", factor: BBL },
  gpm: { dim: "rate", factor: GAL * 1440 },
  "ft3/d": { dim: "rate", factor: FT3 },
  "mscf/d": { dim: "rate", factor: FT3 * 1e3 },
  "mmscf/d": { dim: "rate", factor: FT3 * 1e6 },
  "1000m3/d": { dim: "rate", factor: 1000 },
  // power → W
  w: { dim: "power", factor: 1 },
  kw: { dim: "power", factor: 1000 },
  mw: { dim: "power", factor: 1e6 },
  hp: { dim: "power", factor: 745.699872 },
  // mass → kg
  kg: { dim: "mass", factor: 1 },
  g: { dim: "mass", factor: 0.001 },
  t: { dim: "mass", factor: 1000 },
  lb: { dim: "mass", factor: 0.45359237 },
  // density → kg/m³ (sg = relative to water at 1000 kg/m³; api handled specially)
  "kg/m3": { dim: "density", factor: 1 },
  "g/cc": { dim: "density", factor: 1000 },
  "lb/ft3": { dim: "density", factor: 16.018463373960138 },
  ppg: { dim: "density", factor: 119.82642731689663 },
  sg: { dim: "density", factor: 1000 },
  api: { dim: "density", factor: NaN },
  "psi/ft": { dim: "density", factor: 1000 / 0.433 }, // oilfield convention: fresh water = 0.433 psi/ft
  // viscosity → Pa·s
  cp: { dim: "viscosity", factor: 0.001 },
  "pa.s": { dim: "viscosity", factor: 1 },
  // time → s
  s: { dim: "time", factor: 1 },
  min: { dim: "time", factor: 60 },
  h: { dim: "time", factor: 3600 },
  d: { dim: "time", factor: 86400 },
  // temperature (special)
  c: { dim: "temperature", factor: NaN },
  f: { dim: "temperature", factor: NaN },
  k: { dim: "temperature", factor: NaN },
  r: { dim: "temperature", factor: NaN },
};

const ALIASES: Record<string, string> = {
  meter: "m", meters: "m", metre: "m", feet: "ft", foot: "ft", inch: "in", inches: "in", mile: "mi", miles: "mi",
  psia: "psi", psig: "psi", bara: "bar", barg: "bar", "kgf/cm2": "kg/cm2", "kg/cm²": "kg/cm2", "kpaa": "kpa", "kpag": "kpa",
  "m³": "m3", "cu.m": "m3", "sm3": "m3", liter: "l", liters: "l", litre: "l", litres: "l", bbls: "bbl", stb: "bbl", barrel: "bbl", barrels: "bbl",
  "us gal": "gal", gallon: "gal", gallons: "gal", "cu.ft": "ft3", "ft³": "ft3", scf: "ft3", cf: "ft3", mcf: "mscf", mmcf: "mmscf",
  bpd: "bbl/d", "b/d": "bbl/d", "stb/d": "bbl/d", "bbl/day": "bbl/d", blpd: "bbl/d", bopd: "bbl/d", bwpd: "bbl/d", "m3/day": "m3/d", "m³/d": "m3/d",
  "sm3/d": "m3/d", "m³/h": "m3/h", "m3/hr": "m3/h", "scf/d": "ft3/d", "mcf/d": "mscf/d", mscfd: "mscf/d", "mmcf/d": "mmscf/d", mmscfd: "mmscf/d",
  "10^3m3/d": "1000m3/d", "e3m3/d": "1000m3/d", "10³m³/d": "1000m3/d",
  watt: "w", watts: "w", horsepower: "hp", kilowatt: "kw",
  ton: "t", tonne: "t", tonnes: "t", lbs: "lb", lbm: "lb", pound: "lb", pounds: "lb",
  "kg/m³": "kg/m3", "g/cm3": "g/cc", "g/cm³": "g/cc", "lb/gal": "ppg", "lbm/ft3": "lb/ft3", "°api": "api", gravity: "sg",
  cps: "cp", centipoise: "cp", "pa·s": "pa.s", "pas": "pa.s",
  sec: "s", second: "s", seconds: "s", mins: "min", minute: "min", minutes: "min", hr: "h", hrs: "h", hour: "h", hours: "h", day: "d", days: "d",
  "°c": "c", degc: "c", celsius: "c", "°f": "f", degf: "f", fahrenheit: "f", kelvin: "k", "°r": "r", rankine: "r",
};

export class UnitError extends Error {}

export function normalizeUnit(u: string): string {
  const s = String(u ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/per/g, "/");
  if (UNITS[s]) return s;
  if (ALIASES[s]) return ALIASES[s];
  const noDeg = s.replace(/°/g, "");
  if (UNITS[noDeg]) return noDeg;
  if (ALIASES[noDeg]) return ALIASES[noDeg];
  throw new UnitError(`Unknown unit "${u}". Try e.g. psi, bar, kPa, bbl/d, m3/d, Mscf/d, ft, m, F, C, API, SG, hp, kW, cp.`);
}

function toKelvin(v: number, u: string): number {
  switch (u) {
    case "c":
      return v + 273.15;
    case "f":
      return ((v - 32) * 5) / 9 + 273.15;
    case "r":
      return (v * 5) / 9;
    default:
      return v;
  }
}

function fromKelvin(k: number, u: string): number {
  switch (u) {
    case "c":
      return k - 273.15;
    case "f":
      return ((k - 273.15) * 9) / 5 + 32;
    case "r":
      return (k * 9) / 5;
    default:
      return k;
  }
}

/** Convert value between two unit strings of the same dimension. */
export function convertUnit(value: number, from: string, to: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) throw new UnitError("Value must be a number.");
  const a = normalizeUnit(from);
  const b = normalizeUnit(to);
  const da = UNITS[a].dim;
  const db = UNITS[b].dim;
  if (da !== db) throw new UnitError(`Can't convert ${from} (${da}) to ${to} (${db}).`);
  if (a === b) return value;
  if (da === "temperature") return fromKelvin(toKelvin(value, a), b);
  if (da === "density") {
    // API ↔ anything goes through SG.
    let base: number;
    if (a === "api") {
      if (value <= -131.5) throw new UnitError("API gravity must be greater than -131.5.");
      base = (141.5 / (value + 131.5)) * 1000;
    } else base = value * UNITS[a].factor;
    if (b === "api") {
      const sg = base / 1000;
      if (sg <= 0) throw new UnitError("Density must be positive to convert to API.");
      return 141.5 / sg - 131.5;
    }
    return base / UNITS[b].factor;
  }
  return (value * UNITS[a].factor) / UNITS[b].factor;
}

/** Short list for help text. */
export const UNIT_HELP = "length (ft, m, in, km), pressure (psi, bar, kPa, MPa, atm), volume (bbl, m3, gal, scf, Mscf), rate (bbl/d, m3/d, gpm, Mscf/d, MMscf/d), temperature (F, C, K, R), density (API, SG, kg/m3, ppg, psi/ft), power (hp, kW), mass, viscosity (cp), time";
