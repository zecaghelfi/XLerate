import { describe, expect, it } from "vitest";
import { DEFAULT_CELL_FORMATS } from "../src/core/cellFormatCycle";
import {
  captureToCellFormat,
  captureToFormatCode,
  captureToTextStyle,
  emptyCapture,
  normalizeHexColor,
  type CellFormatCapture,
} from "../src/core/formatCapture";
import { resolveFormatSettings } from "../src/core/formatSettings";
import { DEFAULT_NUMBER_FORMATS } from "../src/core/numberFormatCycle";
import { DEFAULT_TEXT_STYLES } from "../src/core/textStyleCycle";

const line = (style: string, weight = "Thin", color = "#000000") => ({ style, weight, color });

function capture(overrides: Partial<CellFormatCapture> = {}): CellFormatCapture {
  const base = emptyCapture();
  return { ...base, ...overrides, edges: { ...base.edges, ...(overrides.edges ?? {}) } };
}

/** The saved-settings validator keeps a preset only when it is valid. */
const survivesValidation = (key: "cellFormats" | "textStyles", preset: unknown) =>
  JSON.stringify(resolveFormatSettings({ [key]: [preset] })[key]) === JSON.stringify([preset]);

const cellBase = { ...DEFAULT_CELL_FORMATS[0], name: "Mine" };
const textBase = { ...DEFAULT_TEXT_STYLES[3], name: "Mine" };

describe("normalizeHexColor", () => {
  it("normalizes hex and rejects anything else", () => {
    expect(normalizeHexColor("#0b2545")).toBe("#0B2545");
    expect(normalizeHexColor("0b2545")).toBe("#0B2545");
    expect(normalizeHexColor("Automatic")).toBeNull();
    expect(normalizeHexColor(null)).toBeNull();
  });
});

describe("captureToFormatCode", () => {
  it("copies the number format and keeps the name", () => {
    const base = { ...DEFAULT_NUMBER_FORMATS[0], name: "Millions" };
    const { preset, notes } = captureToFormatCode(capture({ numberFormat: '#,##0.0,,;(#,##0.0,,);"-"' }), base);
    expect(preset).toEqual({ name: "Millions", formatCode: '#,##0.0,,;(#,##0.0,,);"-"' });
    expect(notes).toEqual([]);
  });

  it("keeps the preset and explains when Excel reports nothing", () => {
    const base = { ...DEFAULT_NUMBER_FORMATS[0] };
    const { preset, notes } = captureToFormatCode(capture({ numberFormat: "  " }), base);
    expect(preset).toEqual(base);
    expect(notes).toHaveLength(1);
  });
});

describe("captureToCellFormat", () => {
  it("copies fill, font emphasis and a uniform box border", () => {
    const { preset, notes } = captureToCellFormat(
      capture({
        fillPattern: "Solid",
        fillColor: "#eef2f7",
        fontColor: "#0000ff",
        bold: true,
        italic: false,
        underline: true,
        strikethrough: false,
        edges: {
          borderTop: line("Continuous", "Medium", "#0B2545"),
          borderBottom: line("Continuous", "Medium", "#0B2545"),
          borderLeft: line("Continuous", "Medium", "#0B2545"),
          borderRight: line("Continuous", "Medium", "#0B2545"),
          borderDiagonalDown: line("None"),
          borderDiagonalUp: line("None"),
        } as CellFormatCapture["edges"],
      }),
      cellBase
    );
    expect(notes).toEqual([]);
    expect(preset).toMatchObject({
      name: "Mine",
      fillPattern: "Solid",
      fillColor: "#EEF2F7",
      fontColor: "#0000FF",
      fontBold: true,
      fontUnderline: true,
      borderStyle: "Continuous",
      borderWeight: "Medium",
      borderColor: "#0B2545",
      borderTop: true,
      borderBottom: true,
      borderLeft: true,
      borderRight: true,
      // a boxed single cell repeats its box over a block
      borderInsideHorizontal: true,
      borderInsideVertical: true,
      borderDiagonalDown: false,
      borderDiagonalUp: false,
    });
    expect(survivesValidation("cellFormats", preset)).toBe(true);
  });

  it("reads a bottom-only double border as a bottom-only preset", () => {
    const { preset } = captureToCellFormat(
      capture({ edges: { borderBottom: line("Double", "Thick", "#002060") } as CellFormatCapture["edges"] }),
      cellBase
    );
    expect(preset).toMatchObject({
      borderStyle: "Double",
      borderTop: false,
      borderBottom: true,
      borderLeft: false,
      borderRight: false,
      borderInsideHorizontal: false,
      borderInsideVertical: false,
    });
  });

  it("uses inside edges read from a multi-cell selection instead of deriving them", () => {
    const { preset } = captureToCellFormat(
      capture({
        edges: {
          borderTop: line("Continuous"),
          borderBottom: line("Continuous"),
          borderInsideHorizontal: line("None"),
        } as CellFormatCapture["edges"],
      }),
      cellBase
    );
    expect(preset.borderInsideHorizontal).toBe(false);
  });

  it("treats white with unknown pattern as no fill", () => {
    const { preset } = captureToCellFormat(capture({ fillPattern: null, fillColor: "#FFFFFF" }), cellBase);
    expect(preset.fillPattern).toBe("None");
    expect(preset.fillColor).toBe(cellBase.fillColor);
  });

  it("notes when edges use different lines and keeps the bottom one", () => {
    const { preset, notes } = captureToCellFormat(
      capture({
        edges: {
          borderTop: line("Continuous", "Thin"),
          borderBottom: line("Double", "Thick"),
        } as CellFormatCapture["edges"],
      }),
      cellBase
    );
    expect(preset.borderStyle).toBe("Double");
    expect(notes).toHaveLength(1);
    expect(notes[0]).toContain("bottom");
  });

  it("clears borders when the cell has none", () => {
    const boxed = { ...cellBase, borderStyle: "Continuous" as const };
    const { preset } = captureToCellFormat(capture(), boxed);
    expect(preset.borderStyle).toBe("None");
    expect(preset.borderTop).toBe(false);
  });
});

describe("captureToTextStyle", () => {
  it("copies font, size, fill and top/bottom borders", () => {
    const { preset, notes } = captureToTextStyle(
      capture({
        fontName: "Arial",
        fontSize: 14,
        fontColor: "#0B2545",
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        fillPattern: "Solid",
        fillColor: "#EEF2F7",
        edges: {
          borderTop: line("Continuous", "Thin", "#0B2545"),
          borderBottom: line("Continuous", "Thin", "#0B2545"),
        } as CellFormatCapture["edges"],
      }),
      textBase
    );
    expect(notes).toEqual([]);
    expect(preset).toMatchObject({
      name: "Mine",
      fontName: "Arial",
      fontSize: 14,
      fontColor: "#0B2545",
      bold: true,
      fillPattern: "Solid",
      backColor: "#EEF2F7",
      borderStyle: "Continuous",
      borderWeight: "Thin",
      borderColor: "#0B2545",
      borderTop: true,
      borderBottom: true,
      borderLeft: false,
      borderRight: false,
      borderInsideHorizontal: false,
      borderInsideVertical: false,
    });
    expect(survivesValidation("textStyles", preset)).toBe(true);
  });

  it("keeps font settings Excel did not report and says so", () => {
    const { preset, notes } = captureToTextStyle(capture({ strikethrough: true }), textBase);
    expect(preset.fontName).toBe(textBase.fontName);
    expect(preset.fontSize).toBe(textBase.fontSize);
    expect(notes).toHaveLength(2);
  });
});
