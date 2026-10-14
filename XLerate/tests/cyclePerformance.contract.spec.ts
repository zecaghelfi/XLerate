import { describe, expect, it, vi } from "vitest";
import { CellAddress } from "../src/adapters/excelPort";
import { ExcelPortFake } from "../src/adapters/excelPortFake";
import { DEFAULT_NUMBER_FORMATS } from "../src/core/numberFormatCycle";
import { runCycleCellFormat } from "../src/services/cycleCellFormat.service";
import { runCycleDateFormat } from "../src/services/cycleDateFormat.service";
import { runCycleNumberFormat } from "../src/services/cycleNumberFormat.service";
import { runCycleTextStyle } from "../src/services/cycleTextStyle.service";

const addr = (row: number, col: number): CellAddress => ({ sheet: "Sheet1", row, col });
const block = (rows: number, cols: number) =>
  Array.from({ length: rows * cols }, (_, i) => addr(Math.floor(i / cols), i % cols));

describe("cycles touch only what they need (lag fix)", () => {
  it("number, date, cell and text cycles never load full-selection formatting", async () => {
    const port = new ExcelPortFake();
    port.setSelection(block(50, 20));
    const full = vi.spyOn(port, "getSelectionFormatting");
    const perCell = vi.spyOn(port, "applyMutations");
    const oneShot = vi.spyOn(port, "applySelectionNumberFormat");

    await runCycleNumberFormat(port);
    await runCycleDateFormat(port);
    await runCycleCellFormat(port);
    await runCycleTextStyle(port, -1);

    expect(full).not.toHaveBeenCalled();
    expect(oneShot).toHaveBeenCalledTimes(2);
    // the fake implements selection-wide writes via applyMutations internally;
    // services themselves must not build per-cell mutation lists
    expect(perCell.mock.calls.every((call) => call[0].length === 1000)).toBe(true);
  });

  it("restarts at the first preset when the selection mixes formats", async () => {
    const port = new ExcelPortFake();
    port.setCellFormatting(addr(0, 0), { numberFormat: DEFAULT_NUMBER_FORMATS[0].formatCode });
    port.setCellFormatting(addr(1, 0), { numberFormat: "0.000" });
    port.setSelection([addr(0, 0), addr(1, 0)]);
    await runCycleNumberFormat(port);
    const snaps = await port.getSelectionFormatting();
    expect(snaps.map((s) => s.numberFormat)).toEqual([
      DEFAULT_NUMBER_FORMATS[0].formatCode,
      DEFAULT_NUMBER_FORMATS[0].formatCode,
    ]);
  });
});
