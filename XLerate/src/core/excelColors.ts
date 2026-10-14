/**
 * Excel's color picker, as pure data: theme colors with their tint/shade
 * rows, the ten standard colors, and a parser for a workbook theme
 * (xl/theme/theme1.xml). No DOM, no Office.js.
 */

export type ThemeSlot =
  | "lt1"
  | "dk1"
  | "lt2"
  | "dk2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6";

/** Column order of Excel's Theme Colors grid. */
export const THEME_SLOT_ORDER: ThemeSlot[] = [
  "lt1",
  "dk1",
  "lt2",
  "dk2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
];

export const THEME_SLOT_LABELS: Record<ThemeSlot, string> = {
  lt1: "Background 1",
  dk1: "Text 1",
  lt2: "Background 2",
  dk2: "Text 2",
  accent1: "Accent 1",
  accent2: "Accent 2",
  accent3: "Accent 3",
  accent4: "Accent 4",
  accent5: "Accent 5",
  accent6: "Accent 6",
};

export type ThemeColors = {
  name: string;
  /** Hex colors (#RRGGBB) in THEME_SLOT_ORDER. */
  colors: string[];
};

/** Default theme of current Microsoft 365 workbooks ("Office", 2023+). */
export const OFFICE_DEFAULT_THEME: ThemeColors = {
  name: "Office",
  colors: [
    "#FFFFFF",
    "#000000",
    "#E8E8E8",
    "#0E2841",
    "#156082",
    "#E97132",
    "#196B24",
    "#0F9ED5",
    "#A02B93",
    "#4EA72E",
  ],
};

export type NamedColor = { hex: string; label: string };

export const EXCEL_STANDARD_COLORS: NamedColor[] = [
  { hex: "#C00000", label: "Dark Red" },
  { hex: "#FF0000", label: "Red" },
  { hex: "#FFC000", label: "Orange" },
  { hex: "#FFFF00", label: "Yellow" },
  { hex: "#92D050", label: "Light Green" },
  { hex: "#00B050", label: "Green" },
  { hex: "#00B0F0", label: "Light Blue" },
  { hex: "#0070C0", label: "Blue" },
  { hex: "#002060", label: "Dark Blue" },
  { hex: "#7030A0", label: "Purple" },
];

export function normalizeHex(value: string): string | null {
  const trimmed = value.trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(trimmed)) return `#${trimmed.toUpperCase()}`;
  if (/^[0-9a-fA-F]{3}$/.test(trimmed)) {
    const [r, g, b] = trimmed.toUpperCase().split("");
    return `#${r}${r}${g}${g}${b}${b}`;
  }
  return null;
}

type Hsl = { h: number; s: number; l: number };

function hexToRgb(hex: string): [number, number, number] {
  const n = normalizeHex(hex) ?? "#000000";
  return [1, 3, 5].map((i) => parseInt(n.slice(i, i + 2), 16) / 255) as [number, number, number];
}

function rgbToHex(r: number, g: number, b: number): string {
  const part = (v: number) =>
    Math.round(Math.min(1, Math.max(0, v)) * 255)
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  return `#${part(r)}${part(g)}${part(b)}`;
}

export function hexToHsl(hex: string): Hsl {
  const [r, g, b] = hexToRgb(hex);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return { h: h / 6, s, l };
}

export function hslToHex({ h, s, l }: Hsl): string {
  if (s === 0) return rgbToHex(l, l, l);
  const hue = (p: number, q: number, t: number) => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return rgbToHex(hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3));
}

/** Positive = "Lighter x%" (lumMod/lumOff), negative = "Darker x%" (lumMod). */
export type TintStep = number;

/** The five variant rows Excel shows under a theme color, chosen by its luminance. */
export function excelTintSteps(hex: string): TintStep[] {
  const { l } = hexToHsl(hex);
  if (l <= 0.0001) return [0.5, 0.35, 0.25, 0.15, 0.05];
  if (l >= 0.9999) return [-0.05, -0.15, -0.25, -0.35, -0.5];
  if (l < 0.2) return [0.9, 0.75, 0.5, 0.25, 0.1];
  if (l > 0.8) return [-0.1, -0.25, -0.5, -0.75, -0.9];
  return [0.8, 0.6, 0.4, -0.25, -0.5];
}

export function applyTint(hex: string, step: TintStep): string {
  const hsl = hexToHsl(hex);
  const l = step >= 0 ? hsl.l + (1 - hsl.l) * step : hsl.l * (1 + step);
  return hslToHex({ ...hsl, l });
}

export function tintLabel(step: TintStep): string {
  const pct = Math.round(Math.abs(step) * 100);
  return step >= 0 ? `Lighter ${pct}%` : `Darker ${pct}%`;
}

/** Excel's Theme Colors grid: row 0 = base colors, rows 1–5 = tints/shades; one column per slot. */
export function buildThemeGrid(theme: ThemeColors): NamedColor[][] {
  const base = THEME_SLOT_ORDER.map((slot, i) => ({
    hex: normalizeHex(theme.colors[i] ?? "") ?? OFFICE_DEFAULT_THEME.colors[i],
    label: THEME_SLOT_LABELS[slot],
  }));
  const rows: NamedColor[][] = [base];
  for (let r = 0; r < 5; r++) {
    rows.push(
      base.map((c) => {
        const step = excelTintSteps(c.hex)[r];
        return { hex: applyTint(c.hex, step), label: `${c.label}, ${tintLabel(step)}` };
      })
    );
  }
  return rows;
}

/**
 * Reads <a:clrScheme> from a theme part. Accepts a:srgbClr@val and
 * a:sysClr@lastClr. Returns null when any of the 12 scheme slots we need is
 * missing (hlink/folHlink are ignored).
 */
export function parseThemeXml(xml: string): ThemeColors | null {
  const scheme = /<(?:\w+:)?clrScheme\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?clrScheme>/.exec(xml);
  if (!scheme) return null;
  const nameMatch = /\bname="([^"]*)"/.exec(scheme[1]);
  const body = scheme[2];
  const colors: string[] = [];
  for (const slot of THEME_SLOT_ORDER) {
    const block = new RegExp(`<(?:\\w+:)?${slot}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${slot}>`).exec(body);
    if (!block) return null;
    const srgb = /<(?:\w+:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(block[1]);
    const sys = /<(?:\w+:)?sysClr\b[^>]*\blastClr="([0-9A-Fa-f]{6})"/.exec(block[1]);
    const hex = normalizeHex((srgb ?? sys)?.[1] ?? "");
    if (!hex) return null;
    colors.push(hex);
  }
  return { name: nameMatch?.[1] || "Workbook theme", colors };
}

/** Most-recent-first list without duplicates, capped like Excel's Recent Colors row. */
export function pushRecentColor(recent: string[], hex: string, max = 10): string[] {
  const n = normalizeHex(hex);
  if (!n) return recent.slice(0, max);
  return [n, ...recent.filter((c) => normalizeHex(c) !== n)].slice(0, max);
}
