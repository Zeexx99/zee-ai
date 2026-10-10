/** Pure A1-notation helpers (no Office dependency, unit-testable). */

/** 0 -> "A", 25 -> "Z", 26 -> "AA" */
export function columnLetter(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Zero-based row/column -> "B3" */
export function a1(row: number, col: number, absolute = false): string {
  return absolute ? `$${columnLetter(col)}$${row + 1}` : `${columnLetter(col)}${row + 1}`;
}

/** Zero-based block -> "B3:D10" (or a single cell address). */
export function a1Range(row: number, col: number, rows: number, cols: number, absolute = false): string {
  const tl = a1(row, col, absolute);
  if (rows <= 1 && cols <= 1) return tl;
  return `${tl}:${a1(row + rows - 1, col + cols - 1, absolute)}`;
}

/** "'My Sheet'!A1:B2" -> "A1:B2" */
export function stripSheet(address: string): string {
  const i = address.lastIndexOf("!");
  return i >= 0 ? address.slice(i + 1) : address;
}

/** Quote a sheet name for use in a formula reference when needed. */
export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

export const EXCEL_MAX_ROWS = 1048576;
export const EXCEL_MAX_COLS = 16384;
