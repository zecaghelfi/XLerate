/**
 * Excel border vocabulary shared by the cell-format and text-style presets.
 *
 * Mirrors what Excel exposes in Format Cells > Border: the full set of line
 * styles (Office.js `BorderLineStyle`), the four weights (`BorderWeight`) and
 * the eight edges (outline, inside, diagonals). Pure data: no Office.js here.
 */

export const BORDER_LINE_STYLES = [
  "None",
  "Continuous",
  "Dash",
  "DashDot",
  "DashDotDot",
  "Dot",
  "Double",
  "SlantDashDot",
] as const;

export type BorderLineStyle = (typeof BORDER_LINE_STYLES)[number];

export const BORDER_WEIGHTS = ["Hairline", "Thin", "Medium", "Thick"] as const;

export type BorderWeight = (typeof BORDER_WEIGHTS)[number];

/** Edge flags as stored on presets (same names the text-style preset already used). */
export const BORDER_EDGE_FIELDS = [
  "borderTop",
  "borderBottom",
  "borderLeft",
  "borderRight",
  "borderInsideHorizontal",
  "borderInsideVertical",
  "borderDiagonalDown",
  "borderDiagonalUp",
] as const;

export type BorderEdgeField = (typeof BORDER_EDGE_FIELDS)[number];

export type BorderEdgeFlags = Record<BorderEdgeField, boolean>;

export type BorderLinePreset = {
  id: string;
  label: string;
  style: BorderLineStyle;
  weight: BorderWeight;
};

/**
 * The line-style gallery of Excel's Format Cells > Border tab, in Excel's
 * order (left column top-to-bottom, then right column).
 */
export const EXCEL_BORDER_LINE_PRESETS: BorderLinePreset[] = [
  { id: "none", label: "None", style: "None", weight: "Thin" },
  { id: "hairline", label: "Hairline", style: "Continuous", weight: "Hairline" },
  { id: "dot-thin", label: "Dotted", style: "Dot", weight: "Thin" },
  { id: "dashdotdot-thin", label: "Dash dot dot", style: "DashDotDot", weight: "Thin" },
  { id: "dashdot-thin", label: "Dash dot", style: "DashDot", weight: "Thin" },
  { id: "dash-thin", label: "Dashed", style: "Dash", weight: "Thin" },
  { id: "continuous-thin", label: "Thin", style: "Continuous", weight: "Thin" },
  { id: "dashdotdot-medium", label: "Medium dash dot dot", style: "DashDotDot", weight: "Medium" },
  { id: "slantdashdot-medium", label: "Slanted dash dot", style: "SlantDashDot", weight: "Medium" },
  { id: "dashdot-medium", label: "Medium dash dot", style: "DashDot", weight: "Medium" },
  { id: "dash-medium", label: "Medium dashed", style: "Dash", weight: "Medium" },
  { id: "continuous-medium", label: "Medium", style: "Continuous", weight: "Medium" },
  { id: "continuous-thick", label: "Thick", style: "Continuous", weight: "Thick" },
  { id: "double", label: "Double", style: "Double", weight: "Thick" },
];

export function isBorderLineStyle(value: unknown): value is BorderLineStyle {
  return typeof value === "string" && (BORDER_LINE_STYLES as readonly string[]).includes(value);
}

export function isBorderWeight(value: unknown): value is BorderWeight {
  return typeof value === "string" && (BORDER_WEIGHTS as readonly string[]).includes(value);
}

/** Closest gallery entry for a style/weight pair (used to highlight the current choice). */
export function findBorderLinePreset(style: BorderLineStyle, weight: BorderWeight): BorderLinePreset {
  if (style === "None") return EXCEL_BORDER_LINE_PRESETS[0];
  const exact = EXCEL_BORDER_LINE_PRESETS.find((p) => p.style === style && p.weight === weight);
  if (exact) return exact;
  const sameStyle = EXCEL_BORDER_LINE_PRESETS.find((p) => p.style === style);
  return sameStyle ?? EXCEL_BORDER_LINE_PRESETS[6];
}

/**
 * Reads edge flags from a preset, filling missing (legacy) flags with the
 * supplied defaults. Lets old saved settings keep their behavior.
 */
export function resolveBorderEdges(
  source: Partial<Record<BorderEdgeField, unknown>>,
  defaults: BorderEdgeFlags
): BorderEdgeFlags {
  const out = { ...defaults };
  for (const field of BORDER_EDGE_FIELDS) {
    const value = source[field];
    if (typeof value === "boolean") out[field] = value;
  }
  return out;
}

/** Cell-format presets historically bordered every edge plus the inside grid. */
export const CELL_FORMAT_DEFAULT_EDGES: BorderEdgeFlags = {
  borderTop: true,
  borderBottom: true,
  borderLeft: true,
  borderRight: true,
  borderInsideHorizontal: true,
  borderInsideVertical: true,
  borderDiagonalDown: false,
  borderDiagonalUp: false,
};

/** Text-style presets historically had only the four outline flags. */
export const TEXT_STYLE_DEFAULT_EDGES: BorderEdgeFlags = {
  borderTop: false,
  borderBottom: false,
  borderLeft: false,
  borderRight: false,
  borderInsideHorizontal: false,
  borderInsideVertical: false,
  borderDiagonalDown: false,
  borderDiagonalUp: false,
};

/** Excel's Presets row in Format Cells > Border. */
export function applyBorderEdgePreset(
  current: BorderEdgeFlags,
  preset: "none" | "outline" | "inside"
): BorderEdgeFlags {
  if (preset === "none") {
    return { ...TEXT_STYLE_DEFAULT_EDGES };
  }
  if (preset === "outline") {
    return { ...current, borderTop: true, borderBottom: true, borderLeft: true, borderRight: true };
  }
  return { ...current, borderInsideHorizontal: true, borderInsideVertical: true };
}

/** CSS approximation of an Excel border for previews. */
export function borderCss(style: BorderLineStyle, weight: BorderWeight, color: string): string {
  if (style === "None") return "none";
  if (style === "Double") return `3px double ${color}`;
  const width = weight === "Thick" ? 3 : weight === "Medium" ? 2 : 1;
  const css =
    style === "Dot" ? "dotted" : style === "Continuous" ? "solid" : "dashed";
  return `${width}px ${css} ${color}`;
}
