/**
 * Collapses a set of cells into few rectangles so Excel receives
 * "B2:B40,D2:F40" instead of hundreds of single-cell addresses. Pure.
 */

export type CellRect = { row: number; col: number; rowCount: number; colCount: number };

export function coalesceCells(cells: Array<{ row: number; col: number }>): CellRect[] {
  if (cells.length === 0) return [];

  // 1. Vertical runs per column.
  const byCol = new Map<number, number[]>();
  for (const { row, col } of cells) {
    const rows = byCol.get(col);
    if (rows) rows.push(row);
    else byCol.set(col, [row]);
  }
  const runs: Array<{ col: number; r0: number; r1: number }> = [];
  for (const [col, rows] of byCol) {
    rows.sort((a, b) => a - b);
    let r0 = rows[0];
    let prev = rows[0];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (r === prev) continue;
      if (r === prev + 1) {
        prev = r;
        continue;
      }
      runs.push({ col, r0, r1: prev });
      r0 = r;
      prev = r;
    }
    runs.push({ col, r0, r1: prev });
  }

  // 2. Merge identical row spans on adjacent columns.
  runs.sort((a, b) => a.r0 - b.r0 || a.r1 - b.r1 || a.col - b.col);
  const rects: CellRect[] = [];
  for (const run of runs) {
    const last = rects[rects.length - 1];
    if (
      last &&
      last.row === run.r0 &&
      last.row + last.rowCount - 1 === run.r1 &&
      last.col + last.colCount === run.col
    ) {
      last.colCount += 1;
    } else {
      rects.push({ row: run.r0, col: run.col, rowCount: run.r1 - run.r0 + 1, colCount: 1 });
    }
  }
  return rects;
}

export function columnLetters(columnIndex: number): string {
  let remainder = columnIndex + 1;
  let letters = "";
  while (remainder > 0) {
    const zeroBased = (remainder - 1) % 26;
    letters = String.fromCharCode(65 + zeroBased) + letters;
    remainder = Math.floor((remainder - 1) / 26);
  }
  return letters;
}

/** A1-style address of a rectangle (zero-based indexes). */
export function rectToA1(rect: CellRect): string {
  const tl = `${columnLetters(rect.col)}${rect.row + 1}`;
  if (rect.rowCount === 1 && rect.colCount === 1) return tl;
  return `${tl}:${columnLetters(rect.col + rect.colCount - 1)}${rect.row + rect.rowCount}`;
}
