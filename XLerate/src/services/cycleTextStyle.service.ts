import { CellFormatMutation, ExcelPort } from "../adapters/excelPort";
import {
  computeNextTextStyle,
  DEFAULT_TEXT_STYLES,
  TextStyleDefinition,
  textStyleBorderWeight,
  textStyleEdges,
} from "../core/textStyleCycle";

function styleToMutation(style: TextStyleDefinition): CellFormatMutation {
  const weight = textStyleBorderWeight(style);
  const edge =
    style.borderStyle === "None"
      ? undefined
      : { style: style.borderStyle, color: style.borderColor ?? style.fontColor, weight };
  const edges = textStyleEdges(style);

  return {
    font: {
      name: style.fontName,
      size: style.fontSize,
      color: style.fontColor,
      bold: style.bold,
      italic: style.italic,
      underline: style.underline,
    },
    fill: {
      pattern: style.fillPattern,
      color: style.backColor,
    },
    borders: {
      clearAll: true,
      top: edge && edges.borderTop ? edge : undefined,
      bottom: edge && edges.borderBottom ? edge : undefined,
      left: edge && edges.borderLeft ? edge : undefined,
      right: edge && edges.borderRight ? edge : undefined,
      insideHorizontal: edge && edges.borderInsideHorizontal ? edge : undefined,
      insideVertical: edge && edges.borderInsideVertical ? edge : undefined,
      diagonalDown: edge && edges.borderDiagonalDown ? edge : undefined,
      diagonalUp: edge && edges.borderDiagonalUp ? edge : undefined,
    },
  };
}

/**
 * Implements spec §3.10. Advances the cycle index and applies the new
 * text style to the selection. Returns the new index so the caller can
 * persist it for the next invocation (session-scoped per spec §4.2).
 */
export async function runCycleTextStyle(
  port: ExcelPort,
  currentIndex: number,
  configuredStyles: TextStyleDefinition[] = DEFAULT_TEXT_STYLES
): Promise<{ index: number; style: TextStyleDefinition }> {
  // Text styles cycle by a session index, so nothing needs to be read from
  // the sheet: one write call per press.
  const { index, style } = computeNextTextStyle(currentIndex, configuredStyles);
  const mutation = styleToMutation(style);
  await port.applySelectionFormatBundle(mutation);

  return { index, style };
}
