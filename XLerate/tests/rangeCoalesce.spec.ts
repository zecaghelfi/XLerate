import { describe, expect, it } from "vitest";
import { coalesceCells, rectToA1 } from "../src/core/rangeCoalesce";

const toA1 = (cells: Array<{ row: number; col: number }>) => coalesceCells(cells).map(rectToA1);

describe("coalesceCells", () => {
  it("returns nothing for no cells", () => {
    expect(coalesceCells([])).toEqual([]);
  });

  it("merges a column of formulas into one range", () => {
    const cells = Array.from({ length: 40 }, (_, i) => ({ row: i + 1, col: 1 }));
    expect(toA1(cells)).toEqual(["B2:B41"]);
  });

  it("merges adjacent columns with the same rows into a block", () => {
    const cells = [];
    for (let col = 3; col <= 5; col++) for (let row = 1; row <= 39; row++) cells.push({ row, col });
    expect(toA1(cells)).toEqual(["D2:F40"]);
  });

  it("keeps gaps and covers every cell exactly once", () => {
    const cells = [
      { row: 0, col: 0 },
      { row: 1, col: 0 },
      { row: 3, col: 0 },
      { row: 0, col: 1 },
      { row: 1, col: 1 },
      { row: 0, col: 1 },
      { row: 10, col: 30 },
    ];
    const rects = coalesceCells(cells);
    expect(rects.map(rectToA1).sort()).toEqual(["A1:B2", "A4", "AE11"].sort());
    const covered = rects.reduce((n, r) => n + r.rowCount * r.colCount, 0);
    expect(covered).toBe(6);
  });
});
