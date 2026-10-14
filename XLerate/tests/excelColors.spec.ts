import { describe, expect, it } from "vitest";
import {
  applyTint,
  buildThemeGrid,
  EXCEL_STANDARD_COLORS,
  excelTintSteps,
  normalizeHex,
  OFFICE_DEFAULT_THEME,
  parseThemeXml,
  pushRecentColor,
} from "../src/core/excelColors";

describe("Excel color picker data", () => {
  it("has Excel's ten standard colors", () => {
    expect(EXCEL_STANDARD_COLORS.map((c) => c.hex)).toEqual([
      "#C00000",
      "#FF0000",
      "#FFC000",
      "#FFFF00",
      "#92D050",
      "#00B050",
      "#00B0F0",
      "#0070C0",
      "#002060",
      "#7030A0",
    ]);
  });

  it("picks Excel's tint rows by luminance", () => {
    expect(excelTintSteps("#FFFFFF")).toEqual([-0.05, -0.15, -0.25, -0.35, -0.5]);
    expect(excelTintSteps("#000000")).toEqual([0.5, 0.35, 0.25, 0.15, 0.05]);
    expect(excelTintSteps("#4472C4")).toEqual([0.8, 0.6, 0.4, -0.25, -0.5]);
    expect(excelTintSteps("#E7E6E6")).toEqual([-0.1, -0.25, -0.5, -0.75, -0.9]);
  });

  it("reproduces Excel's tints for the classic Office theme (Accent 1 #4472C4)", () => {
    // Values shown by Excel's picker for Blue, Accent 1.
    expect(applyTint("#4472C4", 0.8)).toBe("#DAE3F3");
    expect(applyTint("#4472C4", 0.6)).toBe("#B4C7E7");
    expect(applyTint("#4472C4", 0.4)).toBe("#8FAADC");
    expect(applyTint("#4472C4", -0.25)).toBe("#2F5597");
    expect(applyTint("#4472C4", -0.5)).toBe("#203864");
    expect(applyTint("#FFFFFF", -0.05)).toBe("#F2F2F2");
    expect(applyTint("#000000", 0.5)).toBe("#808080");
  });

  it("builds a 6 x 10 theme grid", () => {
    const grid = buildThemeGrid(OFFICE_DEFAULT_THEME);
    expect(grid).toHaveLength(6);
    for (const row of grid) expect(row).toHaveLength(10);
    expect(grid[0][4].label).toBe("Accent 1");
    expect(grid[1][4].label).toBe("Accent 1, Lighter 80%");
  });

  it("parses a workbook theme part in Excel's column order", () => {
    const xml = `<?xml version="1.0"?><a:theme xmlns:a="x" name="Office Theme"><a:themeElements>
      <a:clrScheme name="Minha Marca">
        <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
        <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="1F3864"/></a:dk2>
        <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
        <a:accent1><a:srgbClr val="002D5A"/></a:accent1>
        <a:accent2><a:srgbClr val="c8102e"/></a:accent2>
        <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
        <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
        <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
        <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      </a:clrScheme></a:themeElements></a:theme>`;
    const theme = parseThemeXml(xml);
    expect(theme?.name).toBe("Minha Marca");
    expect(theme?.colors).toEqual([
      "#FFFFFF",
      "#000000",
      "#EEECE1",
      "#1F3864",
      "#002D5A",
      "#C8102E",
      "#A5A5A5",
      "#FFC000",
      "#5B9BD5",
      "#70AD47",
    ]);
    expect(parseThemeXml("<a:theme/>")).toBeNull();
  });

  it("keeps recent colors unique, newest first, capped at ten", () => {
    let recent: string[] = [];
    for (let i = 0; i < 12; i++) recent = pushRecentColor(recent, `#0000${String(i).padStart(2, "0")}`);
    expect(recent).toHaveLength(10);
    recent = pushRecentColor(recent, "#000005");
    expect(recent[0]).toBe("#000005");
    expect(recent.filter((c) => c === "#000005")).toHaveLength(1);
    expect(normalizeHex("abc")).toBe("#AABBCC");
    expect(normalizeHex("nope")).toBeNull();
  });
});
