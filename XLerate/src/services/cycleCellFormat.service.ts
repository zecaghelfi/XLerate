import { CellFormatMutation, CellFormattingSnapshot, ExcelPort } from "../adapters/excelPort";
import {
  CellFormatDefinition,
  cellFormatEdges,
  computeNextCellFormat,
  DEFAULT_CELL_FORMATS,
  SelectionCellFormatState,
} from "../core/cellFormatCycle";

function snapshotToState(snap: CellFormattingSnapshot): SelectionCellFormatState {
  return {
    fillPattern: snap.fillPattern,
    fillColor: snap.fillColor,
    fontColor: snap.fontColor,
    fontBold: snap.fontBold,
    fontItalic: snap.fontItalic,
    fontUnderline: snap.fontUnderline,
    fontStrikethrough: snap.fontStrikethrough,
    edgeLeftStyle: snap.edgeLeftStyle,
    edgeTopStyle: snap.edgeTopStyle,
    edgeBottomStyle: snap.edgeBottomStyle,
    edgeRightStyle: snap.edgeRightStyle,
    edgeLeftColor: snap.edgeLeftColor,
    edgeTopColor: snap.edgeTopColor,
    edgeBottomColor: snap.edgeBottomColor,
    edgeRightColor: snap.edgeRightColor,
  };
}

function definitionToMutation(def: CellFormatDefinition): CellFormatMutation {
  const edge =
    def.borderStyle === "None"
      ? undefined
      : { style: def.borderStyle, color: def.borderColor, weight: def.borderWeight };
  const edges = cellFormatEdges(def);
  const pick = (on: boolean) => (edge && on ? edge : undefined);
  return {
    fill: { pattern: def.fillPattern, color: def.fillColor },
    font: {
      color: def.fontColor,
      bold: def.fontBold,
      italic: def.fontItalic,
      underline: def.fontUnderline,
      strikethrough: def.fontStrikethrough,
    },
    borders: {
      clearAll: true,
      left: pick(edges.borderLeft),
      top: pick(edges.borderTop),
      bottom: pick(edges.borderBottom),
      right: pick(edges.borderRight),
      insideHorizontal: pick(edges.borderInsideHorizontal),
      insideVertical: pick(edges.borderInsideVertical),
      diagonalDown: pick(edges.borderDiagonalDown),
      diagonalUp: pick(edges.borderDiagonalUp),
    },
  };
}

/**
 * Implements spec §3.8. Reads the first selected cell's formatting, asks
 * the core which preset is next, and applies that preset to every selected
 * cell as one batch.
 */
export async function runCycleCellFormat(
  port: ExcelPort,
  configuredFormats: CellFormatDefinition[] = DEFAULT_CELL_FORMATS
): Promise<void> {
  const snaps = await port.getSelectionFormatting();
  if (snaps.length === 0) return;

  const state = snapshotToState(snaps[0]);
  const next = computeNextCellFormat(state, configuredFormats);
  const mutation = definitionToMutation(next);
  await port.applySelectionFormatBundle(mutation);
}
