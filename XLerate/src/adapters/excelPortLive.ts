import {
  ActiveCellLeftRowSnapshot,
  AutoColorCellSnapshot,
  BorderEdgeMutation,
  BordersMutation,
  CellAddress,
  CellFormatMutation,
  CellFormattingSnapshot,
  CellMutation,
  CellSnapshot,
  ExcelPort,
  FillMutation,
  FontMutation,
} from "./excelPort";
import { autoColorProbe } from "./autoColorProbe";
import { coalesceCells, rectToA1 } from "../core/rangeCoalesce";
import {
  emptyCapture,
  type CapturedBorderEdge,
  type CellFormatCapture,
} from "../core/formatCapture";

/**
 * The only place in the codebase that imports Office.js from domain code.
 * Services depend on the ExcelPort interface; the taskpane constructs this
 * concrete implementation once per feature invocation.
 *
 * Array formula mutations are still deferred (throw). Regular formulas,
 * values, number format, font color, formatBundle, and clearSheetFill are
 * all implemented.
 */
export class ExcelPortLive implements ExcelPort {
  private static readonly MAX_AUTO_COLOR_CELL_COUNT = 50000;
  private static readonly AUTO_COLOR_READ_CHUNK_CELL_BUDGET = 15000;
  /** Max length of one comma-separated address list passed to getRanges. */
  private static readonly ADDRESS_LIST_MAX_CHARS = 6000;
  /** Above this many selected cells the number cycles read only the first cell. */
  private static readonly NUMBER_FORMAT_SCAN_CELL_LIMIT = 20000;

  async getSelectionCells(): Promise<CellSnapshot[]> {
    return Excel.run(async (context) => {
      const areas = await getSelectedAreas(context);

      for (const area of areas) {
        area.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "values", "formulas"]);
        area.worksheet.load("name");
      }
      await context.sync();

      const formulaArrayCells: Excel.Range[][][] = [];
      for (const area of areas) {
        const areaCells: Excel.Range[][] = [];
        for (let r = 0; r < area.rowCount; r++) {
          const rowCells: Excel.Range[] = [];
          for (let c = 0; c < area.columnCount; c++) {
            const cell = area.getCell(r, c);
            // eslint-disable-next-line office-addins/no-navigational-load
            cell.load("formulaArray");
            rowCells.push(cell);
          }
          areaCells.push(rowCells);
        }
        formulaArrayCells.push(areaCells);
      }
      await context.sync();

      const snapshots: CellSnapshot[] = [];
      for (let areaIndex = 0; areaIndex < areas.length; areaIndex++) {
        const area = areas[areaIndex];
        const worksheet = area.worksheet;
        const areaCells = formulaArrayCells[areaIndex];

        for (let r = 0; r < area.rowCount; r++) {
          for (let c = 0; c < area.columnCount; c++) {
            const formula = area.formulas[r][c];
            const value = area.values[r][c];
            const address: CellAddress = {
              sheet: worksheet.name,
              row: area.rowIndex + r,
              col: area.columnIndex + c,
            };
            const formulaText = typeof formula === "string" ? formula : "";
            const formulaArrayText = areaCells[r][c].formulaArray;
            const isArrayFormula =
              typeof formulaArrayText === "string" && formulaArrayText.length > 0;
            const arrayFormula = isArrayFormula
              ? formatArrayFormulaForSnapshot(formulaArrayText)
              : "";
            const isFormula = isArrayFormula || formulaText.startsWith("=");

            snapshots.push({
              address,
              isFormula,
              isArrayFormula,
              formula: isArrayFormula ? arrayFormula : isFormula ? formulaText : "",
              value: isFormula ? undefined : value,
            });
          }
        }
      }
      return snapshots;
    });
  }

  async getSelectionFormatting(): Promise<CellFormattingSnapshot[]> {
    return Excel.run(async (context) => {
      const areas = await getSelectedAreas(context);
      const cellProperties = areas.map((area) => {
        area.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "numberFormat"]);
        area.worksheet.load("name");
        return area.getCellProperties(FORMATTING_PROPERTY_SPEC);
      });
      await context.sync();

      const snapshots: CellFormattingSnapshot[] = [];
      for (let areaIndex = 0; areaIndex < areas.length; areaIndex++) {
        const area = areas[areaIndex];
        const worksheet = area.worksheet;
        const props = cellProperties[areaIndex].value;

        for (let r = 0; r < area.rowCount; r++) {
          for (let c = 0; c < area.columnCount; c++) {
            const address: CellAddress = {
              sheet: worksheet.name,
              row: area.rowIndex + r,
              col: area.columnIndex + c,
            };
            const numberFormatCell =
              Array.isArray(area.numberFormat) && Array.isArray(area.numberFormat[r])
                ? String(area.numberFormat[r][c] ?? "General")
                : "General";
            snapshots.push(toFormattingSnapshot(address, numberFormatCell, props[r][c]));
          }
        }
      }
      return snapshots;
    });
  }

  async getSelectionNumberFormats(): Promise<{ formats: string[]; scanned: boolean } | null> {
    return Excel.run(async (context) => {
      const areas = context.workbook.getSelectedRanges().areas;
      areas.load(["rowCount", "columnCount"]);
      await context.sync();
      if (areas.items.length === 0) return null;

      const totalCells = areas.items.reduce((n, a) => n + a.rowCount * a.columnCount, 0);
      const scanned = totalCells <= ExcelPortLive.NUMBER_FORMAT_SCAN_CELL_LIMIT;
      const first = areas.items[0].getCell(0, 0);
      first.load("numberFormat");
      if (scanned && totalCells > 1) {
        for (const area of areas.items) area.load("numberFormat");
      }
      await context.sync();

      const firstFormat = String(first.numberFormat?.[0]?.[0] ?? "General");
      const formats = [firstFormat];
      if (scanned && totalCells > 1) {
        for (const area of areas.items) {
          for (const row of area.numberFormat) {
            for (const value of row) {
              const fmt = String(value ?? "General");
              if (!formats.includes(fmt)) {
                formats.push(fmt);
                // Two distinct codes already mean "mixed"; stop early.
                if (formats.length > 1) return { formats, scanned };
              }
            }
          }
        }
      }
      return { formats, scanned };
    });
  }

  async applySelectionNumberFormat(format: string): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const areas = context.workbook.getSelectedRanges().areas;
        areas.load("items");
        await context.sync();
        // A scalar is broadcast to every cell of the area: one tiny payload
        // regardless of selection size.
        for (const area of areas.items) area.numberFormat = format as unknown as string[][];
        await context.sync();
      });
    } catch (error) {
      // Fallback for hosts that reject the scalar form: full-size grid per area.
      // eslint-disable-next-line no-console
      console.warn("[XLerate] scalar numberFormat write rejected, using grid write", error);
      await Excel.run(async (context) => {
        const areas = context.workbook.getSelectedRanges().areas;
        areas.load(["rowCount", "columnCount"]);
        await context.sync();
        for (const area of areas.items) {
          area.numberFormat = fillGrid(format, area.rowCount, area.columnCount);
        }
        await context.sync();
      });
    }
  }

  async getSelectionFirstCellFormatting(): Promise<CellFormattingSnapshot | null> {
    return Excel.run(async (context) => {
      const areas = context.workbook.getSelectedRanges().areas;
      areas.load("items");
      await context.sync();
      if (areas.items.length === 0) return null;
      const cell = areas.items[0].getCell(0, 0);
      cell.load(["rowIndex", "columnIndex", "numberFormat"]);
      cell.worksheet.load("name");
      const props = cell.getCellProperties(FORMATTING_PROPERTY_SPEC);
      await context.sync();
      return toFormattingSnapshot(
        { sheet: cell.worksheet.name, row: cell.rowIndex, col: cell.columnIndex },
        String(cell.numberFormat?.[0]?.[0] ?? "General"),
        props.value[0][0]
      );
    });
  }

  async getSelectionFormatCapture(): Promise<CellFormatCapture | null> {
    return Excel.run(async (context) => {
      const areas = context.workbook.getSelectedRanges().areas;
      areas.load("items");
      await context.sync();
      if (areas.items.length === 0) return null;
      const area = areas.items[0];
      area.load(["rowCount", "columnCount"]);
      const cell = area.getCell(0, 0);
      cell.load("numberFormat");
      const props = cell.getCellProperties(CAPTURE_PROPERTY_SPEC);
      await context.sync();

      const format = props.value[0][0]?.format;
      const borders = format?.borders;
      const font = format?.font;
      const capture = emptyCapture();
      capture.numberFormat = String(cell.numberFormat?.[0]?.[0] ?? "General");
      capture.fontName = font?.name ?? null;
      capture.fontSize = typeof font?.size === "number" ? font.size : null;
      capture.fontColor = font?.color ?? null;
      capture.bold = typeof font?.bold === "boolean" ? font.bold : null;
      capture.italic = typeof font?.italic === "boolean" ? font.italic : null;
      capture.underline = typeof font?.underline === "string" ? font.underline !== "None" : null;
      capture.strikethrough = typeof font?.strikethrough === "boolean" ? font.strikethrough : null;
      capture.fillPattern = format?.fill?.pattern ?? null;
      capture.fillColor = format?.fill?.color ?? null;
      capture.edges.borderTop = toCapturedEdge(borders?.top);
      capture.edges.borderBottom = toCapturedEdge(borders?.bottom);
      capture.edges.borderLeft = toCapturedEdge(borders?.left);
      capture.edges.borderRight = toCapturedEdge(borders?.right);
      capture.edges.borderDiagonalDown = toCapturedEdge(borders?.diagonalDown);
      capture.edges.borderDiagonalUp = toCapturedEdge(borders?.diagonalUp);

      // Inside edges exist only on a multi-cell range; reading them on a single
      // cell is meaningless, so they stay null and the core derives them.
      const insideH =
        area.rowCount > 1 ? area.format.borders.getItem(Excel.BorderIndex.insideHorizontal) : null;
      const insideV =
        area.columnCount > 1 ? area.format.borders.getItem(Excel.BorderIndex.insideVertical) : null;
      if (insideH || insideV) {
        insideH?.load(["style", "weight", "color"]);
        insideV?.load(["style", "weight", "color"]);
        try {
          await context.sync();
          if (insideH) capture.edges.borderInsideHorizontal = toCapturedEdge(insideH);
          if (insideV) capture.edges.borderInsideVertical = toCapturedEdge(insideV);
        } catch (error) {
          console.warn("[XLerate] inside borders not readable; deriving from outline", error);
        }
      }
      return capture;
    });
  }

  async getActiveCellLeftRowSnapshot(): Promise<ActiveCellLeftRowSnapshot> {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      const activeCell = workbook.getActiveCell();
      const worksheet = workbook.worksheets.getActiveWorksheet();
      const usedRange = worksheet.getUsedRangeOrNullObject(true);

      activeCell.load(["rowIndex", "columnIndex"]);
      activeCell.worksheet.load("name");
      // eslint-disable-next-line office-addins/no-navigational-load
      usedRange.load(["isNullObject", "columnIndex"]);
      await context.sync();

      const snapshot: ActiveCellLeftRowSnapshot = {
        activeCell: {
          sheet: activeCell.worksheet.name,
          row: activeCell.rowIndex,
          col: activeCell.columnIndex,
        },
        leftCells: [],
      };

      const startCol = usedRange.isNullObject
        ? 0
        : Math.max(0, Math.min(activeCell.columnIndex, usedRange.columnIndex));
      const cellCount = Math.max(0, activeCell.columnIndex - startCol);

      if (cellCount === 0) {
        return snapshot;
      }

      const rowRange = activeCell.worksheet.getRangeByIndexes(
        activeCell.rowIndex,
        startCol,
        1,
        cellCount
      );
      rowRange.load("values");
      await context.sync();

      snapshot.leftCells = Array.from({ length: cellCount }, (_, index) => ({
        address: {
          sheet: activeCell.worksheet.name,
          row: activeCell.rowIndex,
          col: startCol + index,
        },
        value: rowRange.values[0][index],
      }));

      return snapshot;
    });
  }

  async getSelectionAutoColorCells(): Promise<AutoColorCellSnapshot[]> {
    return Excel.run(async (context) => {
      autoColorProbe("20 getSelectionAutoColorCells.enter");
      const snapshots: AutoColorCellSnapshot[] = [];
      const snapshotByKey = new Map<string, AutoColorCellSnapshot>();
      const worksheetCache = new Map<string, Excel.Worksheet>();
      autoColorProbe("21 before-resolve-target-ranges");
      const targetRanges = await resolveAutoColorTargetRanges(context, worksheetCache);
      autoColorProbe("22 after-resolve-target-ranges", {
        count: targetRanges.length,
        sample: targetRanges.slice(0, 10).map((target) => ({
          sheet: target.worksheetName,
          rowIndex: target.rowIndex,
          columnIndex: target.columnIndex,
          rowCount: target.rowCount,
          columnCount: target.columnCount,
        })),
      });
      let totalTargetCellCount = 0;

      for (let targetIndex = 0; targetIndex < targetRanges.length; targetIndex += 1) {
        const target = targetRanges[targetIndex];
        const boundedRowCount = target.rowCount;
        const boundedColumnCount = target.columnCount;
        const chunkRowCount = Math.max(
          1,
          Math.floor(ExcelPortLive.AUTO_COLOR_READ_CHUNK_CELL_BUDGET / boundedColumnCount)
        );
        autoColorProbe("23 target-range-enter", {
          targetIndex,
          sheet: target.worksheetName,
          rowIndex: target.rowIndex,
          columnIndex: target.columnIndex,
          rowCount: boundedRowCount,
          columnCount: boundedColumnCount,
          chunkRowCount,
        });

        for (let offset = 0; offset < boundedRowCount; offset += chunkRowCount) {
          const rowsInChunk = Math.min(chunkRowCount, boundedRowCount - offset);
          const chunkRange = target.worksheet.getRangeByIndexes(
            target.rowIndex + offset,
            target.columnIndex,
            rowsInChunk,
            boundedColumnCount
          );

          const formulaCells = chunkRange.getSpecialCellsOrNullObject(
            Excel.SpecialCellType.formulas
          );
          const constantCells = chunkRange.getSpecialCellsOrNullObject(
            Excel.SpecialCellType.constants
          );
          const textConstantCells = chunkRange.getSpecialCellsOrNullObject(
            Excel.SpecialCellType.constants,
            Excel.SpecialCellValueType.text
          );

          // eslint-disable-next-line office-addins/no-navigational-load
          formulaCells.load(["isNullObject", "cellCount"]);
          // eslint-disable-next-line office-addins/no-navigational-load
          constantCells.load(["isNullObject", "cellCount"]);
          // eslint-disable-next-line office-addins/no-navigational-load
          textConstantCells.load("isNullObject");
          autoColorProbe("24 before-sync-special-cells", {
            targetIndex,
            offset,
            rowsInChunk,
          });
          await context.sync();
          autoColorProbe("25 after-sync-special-cells", {
            targetIndex,
            offset,
            rowsInChunk,
            formulaCellCount: formulaCells.isNullObject ? 0 : formulaCells.cellCount,
            constantCellCount: constantCells.isNullObject ? 0 : constantCells.cellCount,
            textConstantsExist: !textConstantCells.isNullObject,
          });

          totalTargetCellCount += formulaCells.isNullObject ? 0 : formulaCells.cellCount;
          totalTargetCellCount += constantCells.isNullObject ? 0 : constantCells.cellCount;

          if (totalTargetCellCount > ExcelPortLive.MAX_AUTO_COLOR_CELL_COUNT) {
            autoColorProbe("26 abort-too-many-target-cells", {
              totalTargetCellCount,
              limit: ExcelPortLive.MAX_AUTO_COLOR_CELL_COUNT,
            });
            throw new Error(
              `Auto-color supports up to ${ExcelPortLive.MAX_AUTO_COLOR_CELL_COUNT.toLocaleString()} non-empty cells at a time. Narrow the selection and try again.`
            );
          }

          // Load the per-area payload together with the area list (collection
          // load applies to every item): one round trip instead of two.
          const PAYLOAD_PROPS = [
            "rowCount",
            "columnCount",
            "rowIndex",
            "columnIndex",
            "values",
            "formulas",
            "numberFormat",
          ];
          if (!formulaCells.isNullObject) formulaCells.areas.load(PAYLOAD_PROPS);
          if (!constantCells.isNullObject) constantCells.areas.load(PAYLOAD_PROPS);
          if (!textConstantCells.isNullObject) {
            textConstantCells.areas.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
          }
          autoColorProbe("27 before-sync-area-payloads", { targetIndex, offset });
          await context.sync();

          const formulaPayloads = formulaCells.isNullObject ? [] : formulaCells.areas.items;
          const constantPayloads = constantCells.isNullObject ? [] : constantCells.areas.items;
          const textAreas = textConstantCells.isNullObject ? [] : textConstantCells.areas.items;
          autoColorProbe("28 after-sync-area-payloads", {
            targetIndex,
            offset,
            formulaAreaCount: formulaPayloads.length,
            constantAreaCount: constantPayloads.length,
            textAreaCount: textAreas.length,
          });

          // Hyperlinks only live on text cells and need per-area cell properties,
          // so this extra round trip happens only when the chunk has text.
          const textHyperlinkPayloads = textAreas.map((range) =>
            range.getCellProperties({ hyperlink: true })
          );
          if (textAreas.length > 0) {
            autoColorProbe("29 before-sync-hyperlinks", { targetIndex, offset });
            await context.sync();
            autoColorProbe("30 after-sync-hyperlinks", { targetIndex, offset });
          }

          for (const area of formulaPayloads) {
            pushAutoColorSnapshots(area, target.worksheetName, false, snapshots, snapshotByKey);
          }

          for (const area of constantPayloads) {
            pushAutoColorSnapshots(area, target.worksheetName, true, snapshots, snapshotByKey);
          }

          for (let textIndex = 0; textIndex < textAreas.length; textIndex += 1) {
            const area = textAreas[textIndex];
            const props = textHyperlinkPayloads[textIndex].value;
            markHyperlinks(area, target.worksheetName, props, snapshotByKey);
          }
          autoColorProbe("31 chunk-complete", {
            targetIndex,
            offset,
            snapshotCount: snapshots.length,
          });
        }
      }

      autoColorProbe("32 getSelectionAutoColorCells.return", {
        snapshotCount: snapshots.length,
      });
      return snapshots;
    });
  }

  async applySelectionFormatBundle(format: CellFormatMutation): Promise<void> {
    await Excel.run(async (context) => {
      const areas = await getSelectedAreas(context);
      for (const area of areas) {
        applyFormatBundle(area, format);
      }
      await context.sync();
    });
  }

  async applyMutations(mutations: CellMutation[]): Promise<void> {
    if (mutations.length === 0) return;

    await Excel.run(async (context) => {
      autoColorProbe("40 applyMutations.enter", {
        mutationCount: mutations.length,
      });
      const sheetCache = new Map<string, Excel.Worksheet>();
      const sheetFor = (name: string): Excel.Worksheet => {
        let s = sheetCache.get(name);
        if (!s) {
          s = context.workbook.worksheets.getItem(name);
          sheetCache.set(name, s);
        }
        return s;
      };

      const fontColorGroups = new Map<string, CellAddress[]>();

      for (const m of mutations) {
        if (m.kind === "fontColor") {
          const groupKey = `${m.address.sheet}|${m.color}`;
          const group = fontColorGroups.get(groupKey) ?? [];
          group.push(m.address);
          fontColorGroups.set(groupKey, group);
          continue;
        }

        const cell = sheetFor(m.address.sheet).getRangeByIndexes(
          m.address.row,
          m.address.col,
          1,
          1
        );
        if (m.kind === "value") {
          cell.values = [[m.value as string | number | boolean | null]];
        } else if (m.kind === "formula") {
          cell.formulas = [[m.formula]];
        } else if (m.kind === "arrayFormula") {
          cell.formulaArray = normalizeFormulaArrayForWrite(m.formula);
        } else if (m.kind === "numberFormat") {
          cell.numberFormat = [[m.format]];
        } else if (m.kind === "formatBundle") {
          applyFormatBundle(cell, m.format);
        }
      }

      for (const [groupKey, addresses] of fontColorGroups) {
        const [sheetName, color] = groupKey.split("|");
        const worksheet = sheetFor(sheetName);
        // Collapse cells into rectangles first: a column of 500 formulas
        // becomes one "B2:B501" instead of 500 addresses for Excel to parse.
        const refs = coalesceCells(addresses).map(rectToA1);
        for (const chunk of chunkByLength(refs, ExcelPortLive.ADDRESS_LIST_MAX_CHARS)) {
          worksheet.getRanges(chunk.join(",")).format.font.color = color;
        }
      }

      autoColorProbe("41 before-sync-applyMutations", {
        groupCount: fontColorGroups.size,
        groupedTargetCount: Array.from(fontColorGroups.values()).reduce(
          (sum, addresses) => sum + addresses.length,
          0
        ),
      });
      await context.sync();
      autoColorProbe("42 after-sync-applyMutations");
    });
  }

  async clearSheetFill(sheetName: string): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const used = sheet.getUsedRange(true);
      used.load(["rowCount", "columnCount"]);
      await context.sync();
      if (used.rowCount === 0 || used.columnCount === 0) return;
      used.format.fill.clear();
      await context.sync();
    });
  }
}

const FORMATTING_PROPERTY_SPEC: Excel.CellPropertiesLoadOptions = {
  hyperlink: true,
  format: {
    fill: {
      color: true,
      pattern: true,
    },
    font: {
      name: true,
      size: true,
      color: true,
      bold: true,
      italic: true,
      underline: true,
      strikethrough: true,
    },
    borders: {
      color: true,
      style: true,
    },
  },
};

/** Same axes as FORMATTING_PROPERTY_SPEC plus border weight (the cycles never need it). */
const CAPTURE_PROPERTY_SPEC: Excel.CellPropertiesLoadOptions = {
  format: {
    fill: { color: true, pattern: true },
    font: {
      name: true,
      size: true,
      color: true,
      bold: true,
      italic: true,
      underline: true,
      strikethrough: true,
    },
    borders: { color: true, style: true, weight: true },
  },
};

function toCapturedEdge(
  border:
    | { style?: string | null; weight?: string | null; color?: string | null }
    | null
    | undefined
): CapturedBorderEdge | null {
  if (!border || typeof border.style !== "string") return null;
  return {
    style: border.style,
    weight: typeof border.weight === "string" ? border.weight : null,
    color: typeof border.color === "string" ? border.color : null,
  };
}

function toFormattingSnapshot(
  address: CellAddress,
  numberFormat: string,
  cellProp: Excel.CellProperties
): CellFormattingSnapshot {
  const format = cellProp.format;
  const borders = format?.borders;
  return {
    address,
    numberFormat,
    hasHyperlink: hasCellHyperlink(cellProp.hyperlink),
    fillPattern: format?.fill?.pattern ?? null,
    fillColor: format?.fill?.color ?? null,
    fontName: format?.font?.name ?? null,
    fontSize: typeof format?.font?.size === "number" ? format.font.size : null,
    fontColor: format?.font?.color ?? null,
    fontBold: typeof format?.font?.bold === "boolean" ? format.font.bold : null,
    fontItalic: typeof format?.font?.italic === "boolean" ? format.font.italic : null,
    fontUnderline:
      typeof format?.font?.underline === "string" ? format.font.underline !== "None" : null,
    fontStrikethrough:
      typeof format?.font?.strikethrough === "boolean" ? format.font.strikethrough : null,
    edgeLeftStyle: borders?.left?.style ?? null,
    edgeTopStyle: borders?.top?.style ?? null,
    edgeBottomStyle: borders?.bottom?.style ?? null,
    edgeRightStyle: borders?.right?.style ?? null,
    edgeLeftColor: borders?.left?.color ?? null,
    edgeTopColor: borders?.top?.color ?? null,
    edgeBottomColor: borders?.bottom?.color ?? null,
    edgeRightColor: borders?.right?.color ?? null,
  };
}

/** Repeats one value over a rows x cols grid (fallback when a scalar write is rejected). */
function fillGrid<T>(value: T, rows: number, cols: number): T[][] {
  const row = new Array<T>(cols).fill(value);
  return Array.from({ length: rows }, () => row);
}

async function getSelectedAreas(context: Excel.RequestContext): Promise<Excel.Range[]> {
  const selectedRanges = context.workbook.getSelectedRanges();
  const areas = selectedRanges.areas;
  areas.load("items");
  await context.sync();
  return areas.items;
}

function hasCellHyperlink(hyperlink: Excel.RangeHyperlink | undefined): boolean {
  if (!hyperlink) return false;
  return Boolean(
    hyperlink.address ||
    hyperlink.documentReference ||
    hyperlink.screenTip ||
    hyperlink.textToDisplay
  );
}

function cellKey(address: CellAddress): string {
  return `${address.sheet}!${address.row},${address.col}`;
}

function chunkByLength(refs: string[], maxChars: number): string[][] {
  const chunks: string[][] = [];
  let current: string[] = [];
  let length = 0;
  for (const ref of refs) {
    if (current.length > 0 && length + ref.length + 1 > maxChars) {
      chunks.push(current);
      current = [];
      length = 0;
    }
    current.push(ref);
    length += ref.length + 1;
  }
  if (current.length > 0) chunks.push(current);
  return chunks;
}

type AutoColorTargetRange = {
  worksheet: Excel.Worksheet;
  worksheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
};

async function resolveAutoColorTargetRanges(
  context: Excel.RequestContext,
  worksheetCache: Map<string, Excel.Worksheet>
): Promise<AutoColorTargetRange[]> {
  autoColorProbe("50 resolveTargetRanges.enter");
  // The selection always lives on the active sheet, so the areas, the sheet
  // name and its used range come back in a single round trip. (This also
  // handles multi-area selections, where getSelectedRange() would throw, and
  // whole-sheet selections, whose intersection with the used range is the
  // used range itself.)
  const areas = context.workbook.getSelectedRanges().areas;
  areas.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
  const worksheet = context.workbook.worksheets.getActiveWorksheet();
  worksheet.load("name");
  const usedRange = worksheet.getUsedRangeOrNullObject(true);
  usedRange.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();
  autoColorProbe("51 after-sync-selection-and-used-range", {
    areaCount: areas.items.length,
    usedRangeIsNull: usedRange.isNullObject,
  });

  if (usedRange.isNullObject) {
    return [];
  }
  worksheetCache.set(worksheet.name, worksheet);

  const targetRanges: AutoColorTargetRange[] = [];
  for (const area of areas.items) {
    const rowStart = Math.max(area.rowIndex, usedRange.rowIndex);
    const columnStart = Math.max(area.columnIndex, usedRange.columnIndex);
    const rowEnd = Math.min(area.rowIndex + area.rowCount, usedRange.rowIndex + usedRange.rowCount);
    const columnEnd = Math.min(
      area.columnIndex + area.columnCount,
      usedRange.columnIndex + usedRange.columnCount
    );
    if (rowStart >= rowEnd || columnStart >= columnEnd) {
      continue;
    }
    targetRanges.push({
      worksheet,
      worksheetName: worksheet.name,
      rowIndex: rowStart,
      columnIndex: columnStart,
      rowCount: rowEnd - rowStart,
      columnCount: columnEnd - columnStart,
    });
  }

  autoColorProbe("62 resolveTargetRanges.return", { count: targetRanges.length });
  return targetRanges;
}

function pushAutoColorSnapshots(
  area: Excel.Range,
  sheetName: string,
  includeConstants: boolean,
  snapshots: AutoColorCellSnapshot[],
  snapshotByKey: Map<string, AutoColorCellSnapshot>
): void {
  for (let r = 0; r < area.rowCount; r += 1) {
    for (let c = 0; c < area.columnCount; c += 1) {
      const formula = area.formulas[r][c];
      const value = area.values[r][c];
      const formulaText = typeof formula === "string" ? formula : "";
      const isFormula = formulaText.startsWith("=");
      const hasMeaningfulValue =
        isFormula || (value !== null && value !== undefined && value !== "");

      if (!hasMeaningfulValue) {
        continue;
      }

      if (!isFormula && !includeConstants) {
        continue;
      }

      const address: CellAddress = {
        sheet: sheetName,
        row: area.rowIndex + r,
        col: area.columnIndex + c,
      };
      const snapshot: AutoColorCellSnapshot = {
        address,
        isFormula,
        formula: isFormula ? formulaText : "",
        value: isFormula ? undefined : value,
        numberFormat:
          Array.isArray(area.numberFormat) && Array.isArray(area.numberFormat[r])
            ? String(area.numberFormat[r][c] ?? "General")
            : "General",
        hasHyperlink: false,
      };
      snapshots.push(snapshot);
      snapshotByKey.set(cellKey(address), snapshot);
    }
  }
}

function markHyperlinks(
  area: Excel.Range,
  sheetName: string,
  props: Excel.CellProperties[][],
  snapshotByKey: Map<string, AutoColorCellSnapshot>
): void {
  for (let r = 0; r < area.rowCount; r += 1) {
    for (let c = 0; c < area.columnCount; c += 1) {
      const address: CellAddress = {
        sheet: sheetName,
        row: area.rowIndex + r,
        col: area.columnIndex + c,
      };
      const snapshot = snapshotByKey.get(cellKey(address));
      if (snapshot && hasCellHyperlink(props[r][c].hyperlink)) {
        snapshot.hasHyperlink = true;
      }
    }
  }
}

function formatArrayFormulaForSnapshot(formulaArray: string): string {
  const trimmed = formulaArray.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return trimmed;
  }
  if (trimmed.startsWith("=")) {
    return `{${trimmed}}`;
  }
  return `{=${trimmed}}`;
}

function normalizeFormulaArrayForWrite(formula: string): string {
  const trimmed = formula.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return trimmed.slice(1, -1);
  }
  if (trimmed.startsWith("=")) {
    return trimmed;
  }
  return `=${trimmed}`;
}

function applyFontMutation(cell: Excel.Range, font: FontMutation): void {
  if (font.name !== undefined) cell.format.font.name = font.name;
  if (font.size !== undefined) cell.format.font.size = font.size;
  if (font.color !== undefined) cell.format.font.color = font.color;
  if (font.bold !== undefined) cell.format.font.bold = font.bold;
  if (font.italic !== undefined) cell.format.font.italic = font.italic;
  if (font.underline !== undefined) cell.format.font.underline = font.underline ? "Single" : "None";
  if (font.strikethrough !== undefined) cell.format.font.strikethrough = font.strikethrough;
}

function applyFillMutation(cell: Excel.Range, fill: FillMutation): void {
  if (fill.pattern === "None") {
    cell.format.fill.clear();
    return;
  }
  // Order matters: Office.js requires pattern to be set for a color to render
  // as a solid fill on a previously-unfilled cell. Setting color alone does
  // not work reliably. Always set pattern first (when provided), then color.
  if (fill.pattern !== undefined) cell.format.fill.pattern = fill.pattern;
  if (fill.color !== undefined) cell.format.fill.color = fill.color;
}

function applyBorderEdge(cell: Excel.Range, edge: Excel.BorderIndex, m: BorderEdgeMutation): void {
  const border = cell.format.borders.getItem(edge);
  border.style = m.style;
  // Office.js quirk: assigning border.color on a "None"-style border silently
  // upgrades the style to "Continuous". Only set color when the style is
  // visible (not "None"). Mirrors the pre-migration setRangeBorder helper in
  // the old taskpane.ts which had the same guard.
  if (m.style !== "None" && m.color !== undefined) {
    border.color = m.color;
  }
  if (m.style !== "None" && m.style !== "Double" && m.weight !== undefined) {
    border.weight = m.weight as Excel.BorderWeight;
  }
}

function applyBordersMutation(cell: Excel.Range, borders: BordersMutation): void {
  if (borders.clearAll) {
    const edges: Excel.BorderIndex[] = [
      Excel.BorderIndex.edgeLeft,
      Excel.BorderIndex.edgeTop,
      Excel.BorderIndex.edgeBottom,
      Excel.BorderIndex.edgeRight,
      Excel.BorderIndex.insideHorizontal,
      Excel.BorderIndex.insideVertical,
      Excel.BorderIndex.diagonalDown,
      Excel.BorderIndex.diagonalUp,
    ];
    for (const e of edges) cell.format.borders.getItem(e).style = "None";
  }
  if (borders.left) applyBorderEdge(cell, Excel.BorderIndex.edgeLeft, borders.left);
  if (borders.top) applyBorderEdge(cell, Excel.BorderIndex.edgeTop, borders.top);
  if (borders.bottom) applyBorderEdge(cell, Excel.BorderIndex.edgeBottom, borders.bottom);
  if (borders.right) applyBorderEdge(cell, Excel.BorderIndex.edgeRight, borders.right);
  if (borders.insideHorizontal)
    applyBorderEdge(cell, Excel.BorderIndex.insideHorizontal, borders.insideHorizontal);
  if (borders.insideVertical)
    applyBorderEdge(cell, Excel.BorderIndex.insideVertical, borders.insideVertical);
  if (borders.diagonalDown)
    applyBorderEdge(cell, Excel.BorderIndex.diagonalDown, borders.diagonalDown);
  if (borders.diagonalUp) applyBorderEdge(cell, Excel.BorderIndex.diagonalUp, borders.diagonalUp);
}

function applyFormatBundle(cell: Excel.Range, format: CellFormatMutation): void {
  if (format.numberFormat !== undefined) cell.numberFormat = [[format.numberFormat]];
  if (format.font) applyFontMutation(cell, format.font);
  if (format.fill) applyFillMutation(cell, format.fill);
  if (format.borders) applyBordersMutation(cell, format.borders);
}
