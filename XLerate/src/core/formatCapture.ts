/**
 * "Copy from selected cell" for the settings panel: turns the formatting read
 * from one Excel cell into a preset for any of the format cycles.
 *
 * Pure data in, pure data out. The adapter reads the cell (excelPortLive's
 * getSelectionFormatCapture); this module only decides how that formatting
 * maps onto a preset, which is narrower than what Excel can express (one line
 * style per preset, solid fill or none). Anything that does not fit is kept
 * as close as possible and reported in `notes`.
 */

import {
  BORDER_EDGE_FIELDS,
  isBorderLineStyle,
  isBorderWeight,
  type BorderEdgeField,
  type BorderEdgeFlags,
  type BorderLineStyle,
  type BorderWeight,
} from "./borders";
import type { CellFormatDefinition } from "./cellFormatCycle";
import type { NumberFormatDefinition } from "./numberFormatCycle";
import type { TextStyleDefinition } from "./textStyleCycle";

export type CapturedBorderEdge = {
  style: string | null;
  weight: string | null;
  color: string | null;
};

/** Formatting of the sample cell, as read from Excel. `null` = Excel did not report it. */
export type CellFormatCapture = {
  numberFormat: string | null;
  fontName: string | null;
  fontSize: number | null;
  fontColor: string | null;
  bold: boolean | null;
  italic: boolean | null;
  underline: boolean | null;
  strikethrough: boolean | null;
  fillPattern: string | null;
  fillColor: string | null;
  /** Inside edges are only read for a multi-cell selection; null otherwise. */
  edges: Record<BorderEdgeField, CapturedBorderEdge | null>;
};

export type CaptureResult<T> = { preset: T; notes: string[] };

/** Order used to pick the line style when edges disagree: the edge most presets care about first. */
const EDGE_PRIORITY: BorderEdgeField[] = [
  "borderBottom",
  "borderTop",
  "borderLeft",
  "borderRight",
  "borderInsideHorizontal",
  "borderInsideVertical",
  "borderDiagonalDown",
  "borderDiagonalUp",
];

const EDGE_NAMES: Record<BorderEdgeField, string> = {
  borderTop: "top",
  borderBottom: "bottom",
  borderLeft: "left",
  borderRight: "right",
  borderInsideHorizontal: "inside horizontal",
  borderInsideVertical: "inside vertical",
  borderDiagonalDown: "diagonal down",
  borderDiagonalUp: "diagonal up",
};

export function emptyCapture(): CellFormatCapture {
  const edges = {} as Record<BorderEdgeField, CapturedBorderEdge | null>;
  for (const field of BORDER_EDGE_FIELDS) edges[field] = null;
  return {
    numberFormat: null,
    fontName: null,
    fontSize: null,
    fontColor: null,
    bold: null,
    italic: null,
    underline: null,
    strikethrough: null,
    fillPattern: null,
    fillColor: null,
    edges,
  };
}

/** "#rrggbb" / "rrggbb" -> "#RRGGBB"; anything else (theme names, null, "") -> null. */
export function normalizeHexColor(value: string | null | undefined): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!/^#?[0-9a-fA-F]{6}$/.test(trimmed)) return null;
  return `#${trimmed.replace("#", "").toUpperCase()}`;
}

function isEdgeOn(
  edge: CapturedBorderEdge | null
): edge is CapturedBorderEdge & { style: BorderLineStyle } {
  return edge !== null && isBorderLineStyle(edge.style) && edge.style !== "None";
}

function sameLine(a: CapturedBorderEdge, b: CapturedBorderEdge): boolean {
  const weight = (e: CapturedBorderEdge) => (isBorderWeight(e.weight) ? e.weight : "Thin");
  return (
    a.style === b.style &&
    (a.style === "Double" || weight(a) === weight(b)) &&
    normalizeHexColor(a.color) === normalizeHexColor(b.color)
  );
}

type CapturedFill = { pattern: "Solid" | "None"; color: string | null };

/**
 * Excel reports an untouched cell as white with no pattern (or pattern null on
 * desktop even for real solid fills), so: explicit "None", no color, or
 * white-with-unknown-pattern all read as no fill.
 */
function captureFill(capture: CellFormatCapture): CapturedFill {
  const pattern = capture.fillPattern?.trim().toLowerCase() ?? null;
  const color = normalizeHexColor(capture.fillColor);
  if (pattern === "none" || color === null) return { pattern: "None", color: null };
  if (pattern === null && color === "#FFFFFF") return { pattern: "None", color: null };
  return { pattern: "Solid", color };
}

type CapturedBorders = {
  style: BorderLineStyle;
  weight: BorderWeight;
  color: string | null;
  edges: BorderEdgeFlags;
  notes: string[];
};

/**
 * Presets carry one line style/weight/color for all their edges. The edge
 * flags come straight from the cell; the line comes from the first bordered
 * edge in EDGE_PRIORITY, with a note when other edges used a different line.
 */
function captureBorders(capture: CellFormatCapture, deriveInside: boolean): CapturedBorders {
  const edges = {} as BorderEdgeFlags;
  for (const field of BORDER_EDGE_FIELDS) edges[field] = isEdgeOn(capture.edges[field]);

  // A single sample cell has no inside edges. When it is boxed, applying the
  // preset to a block should repeat that box on every cell, so inside edges
  // follow the opposite outline pair (top+bottom -> horizontal, left+right -> vertical).
  if (deriveInside) {
    if (capture.edges.borderInsideHorizontal === null) {
      edges.borderInsideHorizontal = edges.borderTop && edges.borderBottom;
    }
    if (capture.edges.borderInsideVertical === null) {
      edges.borderInsideVertical = edges.borderLeft && edges.borderRight;
    }
  }

  const primaryField = EDGE_PRIORITY.find((field) => isEdgeOn(capture.edges[field]));
  if (!primaryField) {
    return { style: "None", weight: "Thin", color: null, edges, notes: [] };
  }

  const primary = capture.edges[primaryField] as CapturedBorderEdge & { style: BorderLineStyle };
  const notes: string[] = [];
  const differing = EDGE_PRIORITY.filter((field) => {
    const edge = capture.edges[field];
    return field !== primaryField && isEdgeOn(edge) && !sameLine(primary, edge);
  });
  if (differing.length > 0) {
    notes.push(
      `The cell's borders use more than one line; the preset uses the ${EDGE_NAMES[primaryField]} border's line on every edge (${differing.map((f) => EDGE_NAMES[f]).join(", ")} differed).`
    );
  }

  return {
    style: primary.style,
    weight: isBorderWeight(primary.weight) ? primary.weight : "Thin",
    color: normalizeHexColor(primary.color),
    edges,
    notes,
  };
}

function bool(value: boolean | null, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

/** Number / percent / date presets: copies the cell's number format code, keeps the preset name. */
export function captureToFormatCode<T extends NumberFormatDefinition>(
  capture: CellFormatCapture,
  base: T
): CaptureResult<T> {
  const code = capture.numberFormat?.trim();
  if (!code) {
    return {
      preset: { ...base },
      notes: ["Excel did not report a number format for the selected cell."],
    };
  }
  return { preset: { ...base, formatCode: code }, notes: [] };
}

/** Cell Format Cycle preset from the cell's fill, font color/emphasis and borders. */
export function captureToCellFormat(
  capture: CellFormatCapture,
  base: CellFormatDefinition
): CaptureResult<CellFormatDefinition> {
  const fill = captureFill(capture);
  const borders = captureBorders(capture, true);
  const preset: CellFormatDefinition = {
    ...base,
    fillPattern: fill.pattern,
    fillColor: fill.color ?? base.fillColor,
    fontColor: normalizeHexColor(capture.fontColor) ?? base.fontColor,
    fontBold: bool(capture.bold, base.fontBold),
    fontItalic: bool(capture.italic, base.fontItalic),
    fontUnderline: bool(capture.underline, base.fontUnderline),
    fontStrikethrough: bool(capture.strikethrough, base.fontStrikethrough),
    borderStyle: borders.style,
    borderWeight: borders.weight,
    borderColor: borders.color ?? base.borderColor,
    ...borders.edges,
  };
  return { preset, notes: borders.notes };
}

/** Text Style Cycle preset from the cell's font, fill and borders. */
export function captureToTextStyle(
  capture: CellFormatCapture,
  base: TextStyleDefinition
): CaptureResult<TextStyleDefinition> {
  const fill = captureFill(capture);
  const borders = captureBorders(capture, false);
  const notes = [...borders.notes];
  const fontName = capture.fontName?.trim();
  const fontSize =
    typeof capture.fontSize === "number" &&
    Number.isFinite(capture.fontSize) &&
    capture.fontSize > 0
      ? capture.fontSize
      : null;
  if (!fontName || fontSize === null) {
    notes.push("Excel did not report the font name or size; the preset kept its own.");
  }
  if (capture.strikethrough === true) {
    notes.push("Text styles have no strikethrough option; it was not copied.");
  }

  const preset: TextStyleDefinition = {
    ...base,
    fontName: fontName || base.fontName,
    fontSize: fontSize ?? base.fontSize,
    fontColor: normalizeHexColor(capture.fontColor) ?? base.fontColor,
    bold: bool(capture.bold, base.bold),
    italic: bool(capture.italic, base.italic),
    underline: bool(capture.underline, base.underline),
    fillPattern: fill.pattern,
    backColor: fill.color ?? base.backColor,
    borderStyle: borders.style,
    borderWeight: borders.weight,
    borderColor: borders.color ?? base.borderColor ?? base.fontColor,
    ...borders.edges,
  };
  return { preset, notes };
}
