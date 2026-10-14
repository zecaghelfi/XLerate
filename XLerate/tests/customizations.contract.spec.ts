import { describe, expect, it } from "vitest";
import { CellAddress, CellFormatMutation } from "../src/adapters/excelPort";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { CellFormatDefinition, computeNextCellFormat } from "../src/core/cellFormatCycle";
import {
  buildDefaultFormatSettings,
  getFormatSettingsValidationError,
  resolveFormatSettings,
} from "../src/core/formatSettings";
import { DEFAULT_PERCENT_FORMATS } from "../src/core/numberFormatCycle";
import { DEFAULT_TEXT_STYLES } from "../src/core/textStyleCycle";
import { runCycleCellFormat } from "../src/services/cycleCellFormat.service";
import { runCycleNumberFormat } from "../src/services/cycleNumberFormat.service";
import { runCycleTextStyle } from "../src/services/cycleTextStyle.service";

const addr = (row: number, col: number, sheet = "Sheet1"): CellAddress => ({ sheet, row, col });

class CapturingFake extends ExcelPortFake {
  bundles: CellFormatMutation[] = [];
  override async applySelectionFormatBundle(format: CellFormatMutation): Promise<void> {
    this.bundles.push(format);
    await super.applySelectionFormatBundle(format);
  }
}

const bottomOnly: CellFormatDefinition = {
  name: "Total",
  fillPattern: "None",
  fillColor: "#FFFFFF",
  borderStyle: "Double",
  borderWeight: "Thick",
  borderColor: "#002060",
  fontColor: "#000000",
  fontBold: true,
  fontItalic: false,
  fontUnderline: false,
  fontStrikethrough: false,
  borderTop: false,
  borderBottom: true,
  borderLeft: false,
  borderRight: false,
  borderInsideHorizontal: false,
  borderInsideVertical: false,
  borderDiagonalDown: false,
  borderDiagonalUp: false,
};

describe("percent format cycle (Ctrl+Shift+5)", () => {
  it("cycles through the percent presets and wraps", async () => {
    const port = new ExcelPortFake();
    port.setSelection([addr(0, 0)]);
    await runCycleNumberFormat(port, DEFAULT_PERCENT_FORMATS);
    let [snap] = await port.getSelectionFormatting();
    expect(snap.numberFormat).toBe(DEFAULT_PERCENT_FORMATS[0].formatCode);
    await runCycleNumberFormat(port, DEFAULT_PERCENT_FORMATS);
    [snap] = await port.getSelectionFormatting();
    expect(snap.numberFormat).toBe(DEFAULT_PERCENT_FORMATS[1].formatCode);
    await runCycleNumberFormat(port, DEFAULT_PERCENT_FORMATS);
    await runCycleNumberFormat(port, DEFAULT_PERCENT_FORMATS);
    [snap] = await port.getSelectionFormatting();
    expect(snap.numberFormat).toBe(DEFAULT_PERCENT_FORMATS[0].formatCode);
  });

  it("is part of the workbook settings with defaults", () => {
    expect(buildDefaultFormatSettings().percentFormats).toEqual(DEFAULT_PERCENT_FORMATS);
    const custom = resolveFormatSettings(
      JSON.stringify({ percentFormats: [{ name: "bp", formatCode: "0.00%" }] })
    );
    expect(custom.percentFormats).toEqual([{ name: "bp", formatCode: "0.00%" }]);
    expect(resolveFormatSettings(undefined).percentFormats).toEqual(DEFAULT_PERCENT_FORMATS);
  });
});

describe("full Excel border options", () => {
  it("applies only the selected edges with style, weight and color", async () => {
    const port = new CapturingFake();
    port.setSelection([addr(0, 0)]);
    await runCycleCellFormat(port, [bottomOnly]);
    const borders = port.bundles[0].borders!;
    expect(borders.clearAll).toBe(true);
    expect(borders.bottom).toEqual({ style: "Double", weight: "Thick", color: "#002060" });
    expect(borders.top).toBeUndefined();
    expect(borders.insideHorizontal).toBeUndefined();
    expect(borders.diagonalDown).toBeUndefined();
  });

  it("recognises a bottom-only preset when cycling", () => {
    const next = computeNextCellFormat(
      {
        fillPattern: null,
        fillColor: null,
        fontColor: "#000000",
        fontBold: true,
        fontItalic: false,
        fontUnderline: false,
        fontStrikethrough: false,
        edgeLeftStyle: "None",
        edgeTopStyle: "None",
        edgeBottomStyle: "Double",
        edgeRightStyle: "None",
        edgeLeftColor: null,
        edgeTopColor: null,
        edgeBottomColor: "#002060",
        edgeRightColor: null,
      },
      [bottomOnly, { ...bottomOnly, name: "Other", fontBold: false }]
    );
    expect(next.name).toBe("Other");
  });

  it("keeps legacy cell-format presets bordering every edge", async () => {
    const port = new CapturingFake();
    port.setSelection([addr(0, 0)]);
    const legacy = { ...bottomOnly, borderStyle: "Continuous" as const, borderWeight: undefined };
    for (const key of [
      "borderTop",
      "borderBottom",
      "borderLeft",
      "borderRight",
      "borderInsideHorizontal",
      "borderInsideVertical",
      "borderDiagonalDown",
      "borderDiagonalUp",
    ] as const) {
      delete (legacy as Partial<CellFormatDefinition>)[key];
    }
    await runCycleCellFormat(port, [legacy]);
    const b = port.bundles[0].borders!;
    expect([b.top, b.bottom, b.left, b.right, b.insideHorizontal, b.insideVertical].every(Boolean)).toBe(true);
    expect(b.diagonalUp).toBeUndefined();
  });

  it("text styles support diagonals, explicit weight and border color", async () => {
    const port = new CapturingFake();
    port.setSelection([addr(0, 0)]);
    const style = {
      ...DEFAULT_TEXT_STYLES[0],
      borderStyle: "DashDot" as const,
      borderWeight: "Medium" as const,
      borderColor: "#C00000",
      borderDiagonalUp: true,
    };
    await runCycleTextStyle(port, -1, [style]);
    const b = port.bundles[0].borders!;
    expect(b.diagonalUp).toEqual({ style: "DashDot", weight: "Medium", color: "#C00000" });
    expect(b.top).toEqual({ style: "DashDot", weight: "Medium", color: "#C00000" });
  });

  it("validates every Excel line style and weight in saved settings", () => {
    const settings = buildDefaultFormatSettings();
    settings.cellFormats = [{ ...bottomOnly, borderStyle: "SlantDashDot", borderWeight: "Medium" }];
    settings.textStyles = [{ ...DEFAULT_TEXT_STYLES[0], borderStyle: "DashDotDot", borderWeight: "Hairline" }];
    expect(getFormatSettingsValidationError(settings)).toBeNull();
    const roundTrip = resolveFormatSettings(JSON.stringify(settings));
    expect(roundTrip.cellFormats[0].borderStyle).toBe("SlantDashDot");
    expect(roundTrip.textStyles[0].borderWeight).toBe("Hairline");

    const bad = resolveFormatSettings(
      JSON.stringify({ cellFormats: [{ ...bottomOnly, borderWeight: "Huge" }] })
    );
    expect(bad.cellFormats[0].name).not.toBe("Total");
  });
});
