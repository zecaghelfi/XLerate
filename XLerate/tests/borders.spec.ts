import { describe, expect, it } from "vitest";
import {
  applyBorderEdgePreset,
  BORDER_LINE_STYLES,
  CELL_FORMAT_DEFAULT_EDGES,
  EXCEL_BORDER_LINE_PRESETS,
  findBorderLinePreset,
  resolveBorderEdges,
  TEXT_STYLE_DEFAULT_EDGES,
} from "../src/core/borders";

describe("Excel border vocabulary", () => {
  it("offers Excel's 14-entry line-style gallery covering every line style", () => {
    expect(EXCEL_BORDER_LINE_PRESETS).toHaveLength(14);
    const styles = new Set(EXCEL_BORDER_LINE_PRESETS.map((p) => p.style));
    for (const style of BORDER_LINE_STYLES) expect(styles.has(style)).toBe(true);
    const weights = new Set(EXCEL_BORDER_LINE_PRESETS.map((p) => p.weight));
    expect([...weights].sort()).toEqual(["Hairline", "Medium", "Thick", "Thin"]);
  });

  it("finds the gallery entry for a style/weight pair, falling back by style", () => {
    expect(findBorderLinePreset("Continuous", "Medium").id).toBe("continuous-medium");
    expect(findBorderLinePreset("None", "Thick").id).toBe("none");
    expect(findBorderLinePreset("Dot", "Thick").style).toBe("Dot");
  });

  it("fills missing edge flags with the legacy defaults", () => {
    expect(resolveBorderEdges({}, CELL_FORMAT_DEFAULT_EDGES)).toEqual(CELL_FORMAT_DEFAULT_EDGES);
    const edges = resolveBorderEdges(
      { borderTop: true, borderDiagonalUp: true },
      TEXT_STYLE_DEFAULT_EDGES
    );
    expect(edges.borderTop).toBe(true);
    expect(edges.borderDiagonalUp).toBe(true);
    expect(edges.borderBottom).toBe(false);
  });

  it("applies the None / Outline / Inside presets like Excel", () => {
    const outline = applyBorderEdgePreset(TEXT_STYLE_DEFAULT_EDGES, "outline");
    expect([outline.borderTop, outline.borderBottom, outline.borderLeft, outline.borderRight]).toEqual([
      true,
      true,
      true,
      true,
    ]);
    expect(outline.borderInsideHorizontal).toBe(false);
    const inside = applyBorderEdgePreset(outline, "inside");
    expect(inside.borderInsideHorizontal && inside.borderInsideVertical && inside.borderTop).toBe(true);
    expect(applyBorderEdgePreset(inside, "none")).toEqual(TEXT_STYLE_DEFAULT_EDGES);
  });
});
