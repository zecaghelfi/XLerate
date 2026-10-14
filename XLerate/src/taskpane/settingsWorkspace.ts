import { DEFAULT_AUTO_COLOR_PALETTE, type AutoColorPalette } from "../core/autoColor";
import {
  applyBorderEdgePreset,
  BORDER_EDGE_FIELDS,
  borderCss,
  EXCEL_BORDER_LINE_PRESETS,
  findBorderLinePreset,
  isBorderLineStyle,
  isBorderWeight,
  type BorderEdgeField,
  type BorderEdgeFlags,
  type BorderLineStyle,
  type BorderWeight,
} from "../core/borders";
import {
  cellFormatEdges,
  DEFAULT_CELL_FORMATS,
  type CellFormatDefinition,
} from "../core/cellFormatCycle";
import { DEFAULT_DATE_FORMATS, type DateFormatDefinition } from "../core/dateFormatCycle";
import {
  buildDefaultFormatSettings,
  cloneResolvedFormatSettings,
  getFormatSettingsValidationError,
  type ResolvedFormatSettings,
} from "../core/formatSettings";
import {
  DEFAULT_NUMBER_FORMATS,
  DEFAULT_PERCENT_FORMATS,
  type NumberFormatDefinition,
} from "../core/numberFormatCycle";
import {
  DEFAULT_TEXT_STYLES,
  textStyleBorderWeight,
  textStyleEdges,
  type TextStyleDefinition,
} from "../core/textStyleCycle";
import { initColorPicker, renderColorControl } from "./colorPicker";
import { MAX_TRACE_MAX_DEPTH, MAX_TRACE_SAFETY_LIMIT } from "../core/traceUtils";

type SettingsTabKey =
  | "numberFormats"
  | "percentFormats"
  | "dateFormats"
  | "cellFormats"
  | "textStyles"
  | "autoColorPalette"
  | "trace";

type ListTabKey = Exclude<SettingsTabKey, "autoColorPalette" | "trace">;

type SettingsView = "index" | SettingsTabKey;

type SettingsWorkspaceOptions = {
  initialSettings: ResolvedFormatSettings;
  loadSavedSettings: () => Promise<ResolvedFormatSettings>;
  exportSettings: (settings: ResolvedFormatSettings) => Promise<void>;
  saveSettings: (settings: ResolvedFormatSettings) => Promise<void>;
  onStatus: (message: string) => void;
};

type SettingsWorkspaceState = {
  draft: ResolvedFormatSettings;
  savedSnapshot: ResolvedFormatSettings;
  currentView: SettingsView;
  selectedIndex: Record<ListTabKey, number | null>;
};

type TabMeta = {
  title: string;
  description: string;
  group: "Formatting" | "Behavior";
};

const LIST_TABS: ListTabKey[] = [
  "numberFormats",
  "percentFormats",
  "dateFormats",
  "cellFormats",
  "textStyles",
];

const TAB_META: Record<SettingsTabKey, TabMeta> = {
  numberFormats: {
    title: "Number Format Cycle",
    description: "Ordered presets for Cycle Number on the XLerate ribbon.",
    group: "Formatting",
  },
  percentFormats: {
    title: "Percent Format Cycle",
    description: "Ordered presets for the percent cycle (Ctrl+Shift+5).",
    group: "Formatting",
  },
  dateFormats: {
    title: "Date Format Cycle",
    description: "Ordered presets for Cycle Date on the XLerate ribbon.",
    group: "Formatting",
  },
  cellFormats: {
    title: "Cell Format Cycle",
    description: "Fill, font, and border combinations for Cycle Cell.",
    group: "Formatting",
  },
  textStyles: {
    title: "Text Style Cycle",
    description: "Session-cycled presets for headings, sums, and normal rows.",
    group: "Formatting",
  },
  autoColorPalette: {
    title: "Auto-color Palette",
    description: "Semantic font colors used by Auto-color on the ribbon.",
    group: "Behavior",
  },
  trace: {
    title: "Trace Settings",
    description: "Default depth and safety limits for the trace dialog.",
    group: "Behavior",
  },
};

const AUTO_COLOR_FIELDS: Array<{
  key: keyof AutoColorPalette;
  label: string;
  description: string;
}> = [
  {
    key: "input",
    label: "Inputs",
    description: "Typed numeric values that are not formulas.",
  },
  {
    key: "formula",
    label: "Formula",
    description: "Formulas with no external references.",
  },
  {
    key: "worksheetLink",
    label: "Same-sheet link",
    description: "Formulas referencing cells on the active sheet.",
  },
  {
    key: "workbookLink",
    label: "Cross-sheet link",
    description: "Formulas referencing other sheets in the workbook.",
  },
  {
    key: "external",
    label: "External link",
    description: "Formulas referencing another workbook or external source.",
  },
  {
    key: "hyperlink",
    label: "Hyperlink",
    description: "Cells with a hyperlink attached.",
  },
  {
    key: "partialInput",
    label: "Partial input",
    description: "Constants mixed with small references, such as =100+A1.",
  },
];

const COMMON_EXCEL_FONTS = [
  "Aptos",
  "Aptos Display",
  "Aptos Narrow",
  "Arial",
  "Arial Narrow",
  "Bahnschrift",
  "Calibri",
  "Cambria",
  "Candara",
  "Century Gothic",
  "Consolas",
  "Constantia",
  "Corbel",
  "Courier New",
  "Franklin Gothic Book",
  "Garamond",
  "Georgia",
  "Gill Sans MT",
  "Helvetica",
  "Lucida Sans Unicode",
  "Segoe UI",
  "Tahoma",
  "Times New Roman",
  "Trebuchet MS",
  "Verdana",
];

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function isListTab(tab: SettingsTabKey): tab is ListTabKey {
  return LIST_TABS.includes(tab as ListTabKey);
}

function isDetailView(view: SettingsView): view is SettingsTabKey {
  return view !== "index";
}

function serializeSettings(settings: ResolvedFormatSettings): string {
  return JSON.stringify(settings);
}

function safeColor(value: string, fallback = "#F97316"): string {
  const trimmed = value.trim();
  if (/^#?[0-9a-fA-F]{6}$/.test(trimmed)) {
    return trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
  }
  return fallback;
}

function clampIndex(index: number, length: number): number {
  if (length <= 0) {
    return 0;
  }
  if (!Number.isInteger(index) || index < 0) {
    return 0;
  }
  if (index >= length) {
    return length - 1;
  }
  return index;
}

function normalizeSelection(state: SettingsWorkspaceState): void {
  state.selectedIndex.numberFormats =
    state.selectedIndex.numberFormats === null
      ? null
      : clampIndex(state.selectedIndex.numberFormats, state.draft.numberFormats.length);
  state.selectedIndex.percentFormats =
    state.selectedIndex.percentFormats === null
      ? null
      : clampIndex(state.selectedIndex.percentFormats, state.draft.percentFormats.length);
  state.selectedIndex.dateFormats =
    state.selectedIndex.dateFormats === null
      ? null
      : clampIndex(state.selectedIndex.dateFormats, state.draft.dateFormats.length);
  state.selectedIndex.cellFormats =
    state.selectedIndex.cellFormats === null
      ? null
      : clampIndex(state.selectedIndex.cellFormats, state.draft.cellFormats.length);
  state.selectedIndex.textStyles =
    state.selectedIndex.textStyles === null
      ? null
      : clampIndex(state.selectedIndex.textStyles, state.draft.textStyles.length);
}

function getListItems(state: SettingsWorkspaceState, tab: ListTabKey) {
  switch (tab) {
    case "numberFormats":
      return state.draft.numberFormats;
    case "percentFormats":
      return state.draft.percentFormats;
    case "dateFormats":
      return state.draft.dateFormats;
    case "cellFormats":
      return state.draft.cellFormats;
    case "textStyles":
      return state.draft.textStyles;
  }
}

function createNumberFormatPreset(count: number): NumberFormatDefinition {
  return {
    ...DEFAULT_NUMBER_FORMATS[0],
    name: `Number Format ${count + 1}`,
  };
}

function createPercentFormatPreset(count: number): NumberFormatDefinition {
  return {
    ...DEFAULT_PERCENT_FORMATS[0],
    name: `Percent Format ${count + 1}`,
  };
}

function createDateFormatPreset(count: number): DateFormatDefinition {
  return {
    ...DEFAULT_DATE_FORMATS[0],
    name: `Date Format ${count + 1}`,
  };
}

function createCellFormatPreset(count: number): CellFormatDefinition {
  return {
    ...DEFAULT_CELL_FORMATS[0],
    name: `Cell Format ${count + 1}`,
  };
}

function createTextStylePreset(count: number): TextStyleDefinition {
  return {
    ...DEFAULT_TEXT_STYLES[0],
    name: `Text Style ${count + 1}`,
  };
}

function addPreset(state: SettingsWorkspaceState, tab: ListTabKey): void {
  switch (tab) {
    case "numberFormats":
      state.draft.numberFormats.push(createNumberFormatPreset(state.draft.numberFormats.length));
      state.selectedIndex.numberFormats = state.draft.numberFormats.length - 1;
      return;
    case "percentFormats":
      state.draft.percentFormats.push(createPercentFormatPreset(state.draft.percentFormats.length));
      state.selectedIndex.percentFormats = state.draft.percentFormats.length - 1;
      return;
    case "dateFormats":
      state.draft.dateFormats.push(createDateFormatPreset(state.draft.dateFormats.length));
      state.selectedIndex.dateFormats = state.draft.dateFormats.length - 1;
      return;
    case "cellFormats":
      state.draft.cellFormats.push(createCellFormatPreset(state.draft.cellFormats.length));
      state.selectedIndex.cellFormats = state.draft.cellFormats.length - 1;
      return;
    case "textStyles":
      state.draft.textStyles.push(createTextStylePreset(state.draft.textStyles.length));
      state.selectedIndex.textStyles = state.draft.textStyles.length - 1;
      return;
  }
}

function movePreset(
  state: SettingsWorkspaceState,
  tab: ListTabKey,
  index: number,
  direction: number
): void {
  const items = getListItems(state, tab);
  const targetIndex = index + direction;
  if (targetIndex < 0 || targetIndex >= items.length) {
    return;
  }

  const [moved] = items.splice(index, 1);
  items.splice(targetIndex, 0, moved);
  state.selectedIndex[tab] = targetIndex;
}

function deletePreset(state: SettingsWorkspaceState, tab: ListTabKey, index: number): void {
  const items = getListItems(state, tab);
  if (items.length <= 1) {
    return;
  }
  items.splice(index, 1);
  state.selectedIndex[tab] = clampIndex(index - 1, items.length);
}

function updateSelectedListField(
  state: SettingsWorkspaceState,
  field: string,
  value: string | number | boolean
): void {
  if (!isDetailView(state.currentView) || !isListTab(state.currentView)) {
    return;
  }

  const index = state.selectedIndex[state.currentView];
  if (index === null) {
    return;
  }
  const items = getListItems(state, state.currentView);
  const target = items[index] as Record<string, unknown> | undefined;
  if (!target) {
    return;
  }

  target[field] = value;
}

function getSectionSummary(state: SettingsWorkspaceState, tab: SettingsTabKey): string {
  switch (tab) {
    case "numberFormats":
      return `${state.draft.numberFormats.length} presets`;
    case "percentFormats":
      return `${state.draft.percentFormats.length} presets`;
    case "dateFormats":
      return `${state.draft.dateFormats.length} presets`;
    case "cellFormats":
      return `${state.draft.cellFormats.length} presets`;
    case "textStyles":
      return `${state.draft.textStyles.length} presets`;
    case "autoColorPalette":
      return `${AUTO_COLOR_FIELDS.length} colors`;
    case "trace":
      return `Depth ${state.draft.trace.maxDepth} / Limit ${state.draft.trace.safetyLimit}`;
  }
}

function renderNumberDatePreview(item: NumberFormatDefinition | DateFormatDefinition): string {
  return `<span class="accordion-item-subtitle accordion-code">${escapeHtml(item.formatCode)}</span>`;
}

function renderCellFormatPreview(item: CellFormatDefinition): string {
  const background =
    item.fillPattern === "Solid" ? safeColor(item.fillColor, "#FFFFFF") : "#FFFFFF";
  const edges = cellFormatEdges(item);
  const line = borderCss(
    item.borderStyle,
    item.borderWeight ?? "Thin",
    safeColor(item.borderColor, "#D6D3D1")
  );
  const side = (on: boolean) => (on && line !== "none" ? line : "1px dashed #D6D3D1");
  const border = `border-top:${side(edges.borderTop)}; border-bottom:${side(edges.borderBottom)}; border-left:${side(edges.borderLeft)}; border-right:${side(edges.borderRight)}`;
  const textDecoration = [
    item.fontUnderline ? " underline" : "",
    item.fontStrikethrough ? " line-through" : "",
  ].join("");

  return `
    <div class="mini-preview" style="background:${background}; ${border}; color:${safeColor(item.fontColor, "#1C1917")}; font-weight:${item.fontBold ? 700 : 500}; font-style:${item.fontItalic ? "italic" : "normal"}; text-decoration:${textDecoration.trim() || "none"};">
      Aa
    </div>
  `;
}

function renderTextStylePreview(item: TextStyleDefinition): string {
  const edges = textStyleEdges(item);
  const border = borderCss(
    item.borderStyle,
    textStyleBorderWeight(item),
    safeColor(item.borderColor ?? item.fontColor, "#1C1917")
  );
  const style = [
    `background:${item.fillPattern === "None" ? "transparent" : safeColor(item.backColor, "#FFFFFF")};`,
    `color:${safeColor(item.fontColor, "#1C1917")};`,
    `font-size:${Number.isFinite(item.fontSize) ? item.fontSize : 11}px;`,
    item.bold ? "font-weight:700;" : "",
    item.italic ? "font-style:italic;" : "",
    item.underline ? "text-decoration:underline;" : "",
    `border-top:${edges.borderTop ? border : "none"};`,
    `border-bottom:${edges.borderBottom ? border : "none"};`,
    `border-left:${edges.borderLeft ? border : "none"};`,
    `border-right:${edges.borderRight ? border : "none"};`,
  ].join(" ");

  return `<div class="mini-preview text-style-preview" style="${style}">Sample</div>`;
}

function renderField(label: string, control: string, helpText?: string, wide = false): string {
  return `
    <label class="field${wide ? " field-wide" : ""}">
      <span class="field-label">${escapeHtml(label)}</span>
      ${control}
      ${helpText ? `<span class="field-help">${escapeHtml(helpText)}</span>` : ""}
    </label>
  `;
}

function parseEditorValue(
  target: HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement
): string | number | boolean {
  if (target instanceof HTMLInputElement && target.type === "checkbox") {
    return target.checked;
  }

  if (target instanceof HTMLInputElement && target.type === "number") {
    return target.value === "" ? NaN : Number(target.value);
  }

  return target.value;
}

function syncDraftFromDom(root: HTMLDivElement, state: SettingsWorkspaceState): void {
  const controls = root.querySelectorAll<
    HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement
  >(
    "[data-action='update-list-field'], [data-action='update-palette'], [data-action='update-trace']"
  );

  controls.forEach((target) => {
    const action = target.dataset.action;
    if (!action) {
      return;
    }

    if (action === "update-list-field") {
      updateSelectedListField(state, String(target.dataset.field ?? ""), parseEditorValue(target));
      return;
    }

    if (action === "update-palette") {
      const key = target.dataset.key as keyof AutoColorPalette | undefined;
      if (key) {
        state.draft.autoColorPalette[key] = String(parseEditorValue(target));
      }
      return;
    }

    if (action === "update-trace") {
      const key = target.dataset.key as keyof ResolvedFormatSettings["trace"] | undefined;
      if (key) {
        state.draft.trace[key] = Number(parseEditorValue(target));
      }
    }
  });
}

function renderFontOptions(currentFont: string): string {
  const options = COMMON_EXCEL_FONTS.includes(currentFont)
    ? COMMON_EXCEL_FONTS
    : [currentFont, ...COMMON_EXCEL_FONTS];

  return options
    .map(
      (fontName) =>
        `<option value="${escapeHtml(fontName)}" ${fontName === currentFont ? "selected" : ""}>${escapeHtml(fontName)}</option>`
    )
    .join("");
}

function colorControl(field: string, value: string, label: string, disabled = false): string {
  return renderColorControl({
    dataAttrs: `data-action="update-list-field" data-field="${field}"`,
    value,
    label,
    disabled,
  });
}

const LINE_DASH: Record<BorderLineStyle, string> = {
  None: "",
  Continuous: "",
  Dash: "6 3",
  DashDot: "6 3 1.5 3",
  DashDotDot: "6 2 1.5 2 1.5 2",
  Dot: "1.5 2",
  Double: "",
  SlantDashDot: "8 2 2 2",
};

const LINE_WIDTH: Record<BorderWeight, number> = {
  Hairline: 0.6,
  Thin: 1,
  Medium: 2,
  Thick: 3,
};

/** SVG segment for one border line between two points (used in gallery and preview). */
function svgLine(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  style: BorderLineStyle,
  weight: BorderWeight,
  color: string
): string {
  if (style === "None") return "";
  if (style === "Double") {
    const horizontal = Math.abs(y2 - y1) < Math.abs(x2 - x1);
    const dx = horizontal ? 0 : 1.5;
    const dy = horizontal ? 1.5 : 0;
    if (x1 !== x2 && y1 !== y2) {
      return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${color}" stroke-width="3" /><line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#fff" stroke-width="1" />`;
    }
    return `<line x1="${x1 - dx}" y1="${y1 - dy}" x2="${x2 - dx}" y2="${y2 - dy}" stroke="${color}" stroke-width="1" /><line x1="${x1 + dx}" y1="${y1 + dy}" x2="${x2 + dx}" y2="${y2 + dy}" stroke="${color}" stroke-width="1" />`;
  }
  const dash = LINE_DASH[style] ? ` stroke-dasharray="${LINE_DASH[style]}"` : "";
  const opacity = weight === "Hairline" ? ` stroke-opacity="0.6"` : "";
  return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${color}" stroke-width="${LINE_WIDTH[weight]}"${dash}${opacity} />`;
}

const EDGE_LABELS: Record<BorderEdgeField, string> = {
  borderTop: "Top border",
  borderBottom: "Bottom border",
  borderLeft: "Left border",
  borderRight: "Right border",
  borderInsideHorizontal: "Inside horizontal border",
  borderInsideVertical: "Inside vertical border",
  borderDiagonalDown: "Diagonal down border",
  borderDiagonalUp: "Diagonal up border",
};

/** Small icon like Excel's edge buttons: dotted square with the edge drawn solid. */
function edgeIcon(edge: BorderEdgeField): string {
  const line: Record<BorderEdgeField, string> = {
    borderTop: "M2 2H16",
    borderBottom: "M2 16H16",
    borderLeft: "M2 2V16",
    borderRight: "M16 2V16",
    borderInsideHorizontal: "M2 9H16",
    borderInsideVertical: "M9 2V16",
    borderDiagonalDown: "M2 2L16 16",
    borderDiagonalUp: "M2 16L16 2",
  };
  return `<svg viewBox="0 0 18 18" width="18" height="18" aria-hidden="true"><rect x="2" y="2" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1" stroke-dasharray="1 2" opacity="0.5"/><path d="${line[edge]}" stroke="currentColor" stroke-width="2" fill="none"/></svg>`;
}

function edgeButton(edge: BorderEdgeField, edges: BorderEdgeFlags): string {
  return `<button type="button" class="bd-edge${edges[edge] ? " bd-edge-on" : ""}" data-action="toggle-border-edge" data-edge="${edge}" aria-pressed="${edges[edge]}" title="${EDGE_LABELS[edge]}" aria-label="${EDGE_LABELS[edge]}">${edgeIcon(edge)}</button>`;
}

function renderBorderDesigner(options: {
  style: BorderLineStyle;
  weight: BorderWeight;
  color: string;
  edges: BorderEdgeFlags;
}): string {
  const { style, weight, color, edges } = options;
  const selected = findBorderLinePreset(style, weight);
  const drawStyle: BorderLineStyle = style === "None" ? "Continuous" : style;
  const drawColor = style === "None" ? "#B8B2A7" : color;
  const W = 140;
  const H = 90;
  const p = 10;
  const mx = W / 2;
  const my = H / 2;
  const seg = (on: boolean, x1: number, y1: number, x2: number, y2: number) =>
    on && style !== "None" ? svgLine(x1, y1, x2, y2, drawStyle, weight, drawColor) : "";
  const guides = `
    <rect x="${p}" y="${p}" width="${W - 2 * p}" height="${H - 2 * p}" fill="#fff" stroke="#CFC9BE" stroke-width="1" stroke-dasharray="2 3"/>
    <line x1="${mx}" y1="${p}" x2="${mx}" y2="${H - p}" stroke="#E2DDD3" stroke-dasharray="2 3"/>
    <line x1="${p}" y1="${my}" x2="${W - p}" y2="${my}" stroke="#E2DDD3" stroke-dasharray="2 3"/>
    <text x="${(p + mx) / 2}" y="${(p + my) / 2 + 4}" text-anchor="middle" font-size="10" fill="#8C857A">Text</text>
    <text x="${(mx + W - p) / 2}" y="${(p + my) / 2 + 4}" text-anchor="middle" font-size="10" fill="#8C857A">Text</text>
    <text x="${(p + mx) / 2}" y="${(my + H - p) / 2 + 4}" text-anchor="middle" font-size="10" fill="#8C857A">Text</text>
    <text x="${(mx + W - p) / 2}" y="${(my + H - p) / 2 + 4}" text-anchor="middle" font-size="10" fill="#8C857A">Text</text>`;
  const lines = [
    seg(edges.borderTop, p, p, W - p, p),
    seg(edges.borderBottom, p, H - p, W - p, H - p),
    seg(edges.borderLeft, p, p, p, H - p),
    seg(edges.borderRight, W - p, p, W - p, H - p),
    seg(edges.borderInsideHorizontal, p, my, W - p, my),
    seg(edges.borderInsideVertical, mx, p, mx, H - p),
    seg(edges.borderDiagonalDown, p, p, W - p, H - p),
    seg(edges.borderDiagonalUp, p, H - p, W - p, p),
  ].join("");

  const gallery = EXCEL_BORDER_LINE_PRESETS.map((preset) => {
    const isSel = preset.id === selected.id;
    const art =
      preset.style === "None"
        ? `<span class="bd-line-none">None</span>`
        : `<svg viewBox="0 0 60 10" width="60" height="10" aria-hidden="true">${svgLine(2, 5, 58, 5, preset.style, preset.weight, "#1D1B18")}</svg>`;
    return `<button type="button" class="bd-line${isSel ? " bd-line-on" : ""}" data-action="set-border-line" data-style="${preset.style}" data-weight="${preset.weight}" aria-pressed="${isSel}" title="${escapeHtml(preset.label)}" aria-label="${escapeHtml(preset.label)}">${art}</button>`;
  }).join("");

  return `
    <div class="border-designer">
      <div class="bd-col bd-col-line">
        <span class="field-label">Line style</span>
        <div class="bd-line-grid">${gallery}</div>
        <span class="field-label">Color</span>
        ${colorControl("borderColor", color, "Border color")}
      </div>
      <div class="bd-col bd-col-edges">
        <span class="field-label">Presets</span>
        <div class="bd-presets">
          <button type="button" class="bd-preset" data-action="border-edge-preset" data-preset="none">None</button>
          <button type="button" class="bd-preset" data-action="border-edge-preset" data-preset="outline">Outline</button>
          <button type="button" class="bd-preset" data-action="border-edge-preset" data-preset="inside">Inside</button>
        </div>
        <span class="field-label">Border</span>
        <div class="bd-board">
          <div class="bd-side">
            ${edgeButton("borderTop", edges)}
            ${edgeButton("borderInsideHorizontal", edges)}
            ${edgeButton("borderBottom", edges)}
          </div>
          <svg class="bd-preview" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="Border preview">${guides}${lines}</svg>
        </div>
        <div class="bd-bottom">
          ${edgeButton("borderDiagonalUp", edges)}
          ${edgeButton("borderLeft", edges)}
          ${edgeButton("borderInsideVertical", edges)}
          ${edgeButton("borderRight", edges)}
          ${edgeButton("borderDiagonalDown", edges)}
        </div>
        <span class="field-help">Inside borders apply only when several cells are selected.</span>
      </div>
    </div>
  `;
}

type BorderEditable = {
  borderStyle: BorderLineStyle;
  borderWeight?: BorderWeight;
} & Partial<BorderEdgeFlags>;

function getSelectedBorderItem(
  state: SettingsWorkspaceState
): { item: BorderEditable; edges: BorderEdgeFlags } | null {
  if (state.currentView === "cellFormats") {
    const index = state.selectedIndex.cellFormats;
    const item = index === null ? undefined : state.draft.cellFormats[index];
    return item ? { item, edges: cellFormatEdges(item) } : null;
  }
  if (state.currentView === "textStyles") {
    const index = state.selectedIndex.textStyles;
    const item = index === null ? undefined : state.draft.textStyles[index];
    return item ? { item, edges: textStyleEdges(item) } : null;
  }
  return null;
}

function writeEdges(item: BorderEditable, edges: BorderEdgeFlags): void {
  for (const field of BORDER_EDGE_FIELDS) {
    item[field] = edges[field];
  }
  const anyOn = BORDER_EDGE_FIELDS.some((field) => edges[field]);
  if (anyOn && item.borderStyle === "None") {
    item.borderStyle = "Continuous";
    item.borderWeight = "Thin";
  }
}

function handleBorderAction(state: SettingsWorkspaceState, target: HTMLElement): boolean {
  const action = target.dataset.action;
  if (action !== "set-border-line" && action !== "toggle-border-edge" && action !== "border-edge-preset") {
    return false;
  }
  const selected = getSelectedBorderItem(state);
  if (!selected) return true;
  const { item, edges } = selected;

  if (action === "set-border-line") {
    const style = target.dataset.style;
    const weight = target.dataset.weight;
    if (isBorderLineStyle(style) && isBorderWeight(weight)) {
      item.borderStyle = style;
      item.borderWeight = weight;
    }
    return true;
  }
  if (action === "toggle-border-edge") {
    const edge = target.dataset.edge as BorderEdgeField | undefined;
    if (edge && (BORDER_EDGE_FIELDS as readonly string[]).includes(edge)) {
      writeEdges(item, { ...edges, [edge]: !edges[edge] });
    }
    return true;
  }
  const preset = target.dataset.preset;
  if (preset === "none" || preset === "outline" || preset === "inside") {
    writeEdges(item, applyBorderEdgePreset(edges, preset));
  }
  return true;
}

function renderNumberDateEditorContent(
  item: NumberFormatDefinition | DateFormatDefinition,
  typeLabel: string
): string {
  return `
    <div class="field-grid">
      ${renderField(
        "Preset name",
        `<input data-action="update-list-field" data-field="name" type="text" value="${escapeHtml(item.name)}" />`
      )}
      ${renderField(
        "Excel format code",
        `<input data-action="update-list-field" data-field="formatCode" type="text" value="${escapeHtml(item.formatCode)}" />`,
        `This code is applied when ${typeLabel.toLowerCase()} cycles reach this preset.`,
        true
      )}
    </div>
    <div class="preview-panel">
      <span class="preview-label">Format code</span>
      <code>${escapeHtml(item.formatCode)}</code>
    </div>
  `;
}

function renderCellFormatEditorContent(item: CellFormatDefinition): string {
  return `
    <div class="field-grid">
      ${renderField(
        "Preset name",
        `<input data-action="update-list-field" data-field="name" type="text" value="${escapeHtml(item.name)}" />`
      )}
      ${renderField(
        "Fill pattern",
        `<select data-action="update-list-field" data-field="fillPattern">
          <option value="Solid" ${item.fillPattern === "Solid" ? "selected" : ""}>Solid</option>
          <option value="None" ${item.fillPattern === "None" ? "selected" : ""}>None</option>
        </select>`
      )}
      ${renderField(
        "Fill color",
        colorControl("fillColor", safeColor(item.fillColor, "#FFFFFF"), "Fill color", item.fillPattern === "None")
      )}
      ${renderField("Font color", colorControl("fontColor", safeColor(item.fontColor, "#1C1917"), "Font color"))}
    </div>
    ${renderBorderDesigner({
      style: item.borderStyle,
      weight: item.borderWeight ?? "Thin",
      color: safeColor(item.borderColor, "#000000"),
      edges: cellFormatEdges(item),
    })}
    <div class="toggle-grid">
      <label class="toggle"><input data-action="update-list-field" data-field="fontBold" type="checkbox" ${item.fontBold ? "checked" : ""} /> Bold</label>
      <label class="toggle"><input data-action="update-list-field" data-field="fontItalic" type="checkbox" ${item.fontItalic ? "checked" : ""} /> Italic</label>
      <label class="toggle"><input data-action="update-list-field" data-field="fontUnderline" type="checkbox" ${item.fontUnderline ? "checked" : ""} /> Underline</label>
      <label class="toggle"><input data-action="update-list-field" data-field="fontStrikethrough" type="checkbox" ${item.fontStrikethrough ? "checked" : ""} /> Strike</label>
    </div>
    <div class="preview-panel">
      <span class="preview-label">Preview</span>
      ${renderCellFormatPreview(item)}
    </div>
  `;
}

function renderTextStyleEditorContent(item: TextStyleDefinition): string {
  return `
    <div class="field-grid">
      ${renderField(
        "Preset name",
        `<input data-action="update-list-field" data-field="name" type="text" value="${escapeHtml(item.name)}" />`
      )}
      ${renderField(
        "Font name",
        `<select data-action="update-list-field" data-field="fontName">
          ${renderFontOptions(item.fontName)}
        </select>`
      )}
      ${renderField(
        "Font size",
        `<input data-action="update-list-field" data-field="fontSize" type="number" min="1" step="1" value="${item.fontSize}" />`
      )}
      ${renderField("Font color", colorControl("fontColor", safeColor(item.fontColor, "#1C1917"), "Font color"))}
      ${renderField(
        "Background fill",
        `<select data-action="update-list-field" data-field="fillPattern">
          <option value="Solid" ${item.fillPattern === "Solid" ? "selected" : ""}>Solid</option>
          <option value="None" ${item.fillPattern === "None" ? "selected" : ""}>No fill</option>
        </select>`
      )}
      ${renderField(
        "Background color",
        colorControl(
          "backColor",
          safeColor(item.backColor, "#FFFFFF"),
          "Background color",
          item.fillPattern === "None"
        ),
        item.fillPattern === "None" ? "Choose Solid to enable a fill color." : undefined
      )}
    </div>
    <div class="toggle-grid">
      <label class="toggle"><input data-action="update-list-field" data-field="bold" type="checkbox" ${item.bold ? "checked" : ""} /> Bold</label>
      <label class="toggle"><input data-action="update-list-field" data-field="italic" type="checkbox" ${item.italic ? "checked" : ""} /> Italic</label>
      <label class="toggle"><input data-action="update-list-field" data-field="underline" type="checkbox" ${item.underline ? "checked" : ""} /> Underline</label>
    </div>
    ${renderBorderDesigner({
      style: item.borderStyle,
      weight: textStyleBorderWeight(item),
      color: safeColor(item.borderColor ?? item.fontColor, "#000000"),
      edges: textStyleEdges(item),
    })}
    <div class="preview-panel">
      <span class="preview-label">Preview</span>
      ${renderTextStylePreview(item)}
    </div>
  `;
}

function renderPaletteEditor(palette: AutoColorPalette): string {
  return `
    <div class="content-card">
      <div class="content-card-header">
        <div>
          <p class="section-kicker">Auto-color palette</p>
          <h3>Semantic workbook colors</h3>
        </div>
      </div>
      <div class="palette-grid">
        ${AUTO_COLOR_FIELDS.map(
          (field) => `
            <label class="palette-row">
              <span class="palette-copy">
                <span class="palette-label">${escapeHtml(field.label)}</span>
                <span class="palette-help">${escapeHtml(field.description)}</span>
              </span>
              <span class="palette-inputs">
                ${renderColorControl({
                  dataAttrs: `data-action="update-palette" data-key="${field.key}"`,
                  value: safeColor(palette[field.key], DEFAULT_AUTO_COLOR_PALETTE[field.key]),
                  label: field.label,
                })}
              </span>
            </label>
          `
        ).join("")}
      </div>
    </div>
  `;
}

function renderTraceEditor(settings: ResolvedFormatSettings["trace"]): string {
  return `
    <div class="content-card">
      <div class="content-card-header">
        <div>
          <p class="section-kicker">Trace dialog defaults</p>
          <h3>Expansion controls</h3>
        </div>
      </div>
      <div class="field-grid">
        ${renderField(
          "Maximum depth",
          `<input data-action="update-trace" data-key="maxDepth" type="number" min="1" max="${MAX_TRACE_MAX_DEPTH}" step="1" value="${settings.maxDepth}" />`,
          `XLerate clamps this between 1 and ${MAX_TRACE_MAX_DEPTH}.`
        )}
        ${renderField(
          "Safety limit",
          `<input data-action="update-trace" data-key="safetyLimit" type="number" min="1" max="${MAX_TRACE_SAFETY_LIMIT}" step="1" value="${settings.safetyLimit}" />`,
          `Hard stop for rendered rows. XLerate clamps this between 1 and ${MAX_TRACE_SAFETY_LIMIT}.`
        )}
      </div>
      <div class="preview-panel trace-stats-panel">
        <span class="preview-label">Current defaults</span>
        <div class="trace-stat-grid">
          <div><strong>${settings.maxDepth}</strong><span>levels</span></div>
          <div><strong>${settings.safetyLimit}</strong><span>rows</span></div>
        </div>
      </div>
    </div>
  `;
}

function renderDetailContent(state: SettingsWorkspaceState, tab: SettingsTabKey): string {
  switch (tab) {
    case "numberFormats":
    case "percentFormats":
      return renderExpandablePresetList(state, tab);
    case "dateFormats":
      return renderExpandablePresetList(state, tab);
    case "cellFormats":
      return renderExpandablePresetList(state, tab);
    case "textStyles":
      return renderExpandablePresetList(state, tab);
    case "autoColorPalette":
      return `<div class="detail-grid detail-grid-single">${renderPaletteEditor(state.draft.autoColorPalette)}</div>`;
    case "trace":
      return `<div class="detail-grid detail-grid-single">${renderTraceEditor(state.draft.trace)}</div>`;
  }
}

function renderExpandablePresetEditor(
  state: SettingsWorkspaceState,
  tab: ListTabKey,
  index: number
): string {
  switch (tab) {
    case "numberFormats":
      return renderNumberDateEditorContent(state.draft.numberFormats[index], "Number format");
    case "percentFormats":
      return renderNumberDateEditorContent(state.draft.percentFormats[index], "Percent format");
    case "dateFormats":
      return renderNumberDateEditorContent(state.draft.dateFormats[index], "Date format");
    case "cellFormats":
      return renderCellFormatEditorContent(state.draft.cellFormats[index]);
    case "textStyles":
      return renderTextStyleEditorContent(state.draft.textStyles[index]);
  }
}

function renderExpandablePresetList(state: SettingsWorkspaceState, tab: ListTabKey): string {
  const items = getListItems(state, tab);
  const openIndex = state.selectedIndex[tab];

  return `
    <div class="detail-grid detail-grid-single">
      <section class="accordion-panel">
        <div class="accordion-panel-header">
          <div>
            <p class="section-kicker">Presets</p>
            <h3>${escapeHtml(TAB_META[tab].title)}</h3>
          </div>
          <button type="button" class="secondary-btn compact-btn" data-action="add-item">Add</button>
        </div>
        <div class="accordion-list">
          ${items
            .map((item, index) => {
              const isOpen = openIndex === index;
              const title = "name" in item ? escapeHtml(String(item.name)) : `Preset ${index + 1}`;
              const subtitle =
                tab === "numberFormats" || tab === "percentFormats" || tab === "dateFormats"
                  ? renderNumberDatePreview(item as NumberFormatDefinition | DateFormatDefinition)
                  : tab === "cellFormats"
                    ? renderCellFormatPreview(item as CellFormatDefinition)
                    : renderTextStylePreview(item as TextStyleDefinition);

              return `
                <div class="accordion-item${isOpen ? " accordion-item-open" : ""}">
                  <div class="accordion-item-row">
                    <button type="button" class="accordion-item-button" data-action="select-item" data-index="${index}" aria-expanded="${isOpen}">
                      <span class="accordion-item-order">${index + 1}</span>
                      <span class="accordion-item-copy">
                        <span class="accordion-item-title">${title}</span>
                        ${subtitle}
                      </span>
                      <span class="accordion-item-chevron">${isOpen ? "▾" : "▸"}</span>
                    </button>
                  </div>
                  <div class="accordion-item-actions">
                    <button type="button" class="mini-btn" data-action="move-item" data-index="${index}" data-direction="-1" ${index === 0 ? "disabled" : ""}>Up</button>
                    <button type="button" class="mini-btn" data-action="move-item" data-index="${index}" data-direction="1" ${index === items.length - 1 ? "disabled" : ""}>Down</button>
                    <button type="button" class="mini-btn danger-btn" data-action="delete-item" data-index="${index}" ${items.length <= 1 ? "disabled" : ""}>Delete</button>
                  </div>
                  ${
                    isOpen
                      ? `
                        <div class="accordion-item-body">
                          ${renderExpandablePresetEditor(state, tab, index)}
                        </div>
                      `
                      : ""
                  }
                </div>
              `;
            })
            .join("")}
        </div>
      </section>
    </div>
  `;
}

function renderIndex(state: SettingsWorkspaceState): string {
  const groups: Array<"Formatting" | "Behavior"> = ["Formatting", "Behavior"];

  return `
    <div class="settings-shell settings-shell-home">
      <header class="app-brand">
        <span class="app-wordmark">XLerate</span>
        <span class="app-brand-meta">Workbook tools</span>
      </header>
      <section class="settings-intro">
        <div class="settings-home-copy">
          <p class="section-kicker">Workbook</p>
          <h1>Settings</h1>
          <p class="section-copy">
            Choose an area to edit. Changes apply to this workbook after saving.
          </p>
        </div>
      </section>
      <nav class="settings-index" aria-label="Settings areas">
        ${groups
          .map((group) => {
            const tabs = (Object.keys(TAB_META) as SettingsTabKey[]).filter(
              (tab) => TAB_META[tab].group === group
            );

            return `
              <section class="index-group" aria-labelledby="settings-group-${group.toLowerCase()}">
                <h2 class="index-group-label" id="settings-group-${group.toLowerCase()}">${escapeHtml(group)}</h2>
                <div class="index-link-list">
                  ${tabs
                    .map(
                      (tab) => `
                        <button type="button" class="index-link" data-action="open-section" data-tab="${tab}" aria-label="Open ${escapeHtml(TAB_META[tab].title)}">
                          <span class="index-link-copy">
                            <span class="index-link-title">${escapeHtml(TAB_META[tab].title)}</span>
                          </span>
                          <span class="index-link-meta">
                            <span class="index-link-summary">${escapeHtml(getSectionSummary(state, tab))}</span>
                            <span class="index-link-arrow">&rsaquo;</span>
                          </span>
                        </button>
                      `
                    )
                    .join("")}
                </div>
              </section>
            `;
          })
          .join("")}
      </nav>
    </div>
  `;
}

function renderDetail(state: SettingsWorkspaceState, tab: SettingsTabKey): string {
  const dirty = serializeSettings(state.draft) !== serializeSettings(state.savedSnapshot);
  const validationError = getFormatSettingsValidationError(state.draft);

  return `
    <div class="settings-shell settings-shell-detail">
      <section class="detail-surface" aria-labelledby="settings-detail-title">
        <div class="detail-shell-bar">
          <button type="button" class="back-button" data-action="go-home" aria-label="Back to settings">
            <span aria-hidden="true">&larr;</span>
          </button>
          <strong>Workbook settings</strong>
        </div>
        <div class="detail-main">
          <div class="detail-topbar">
            <div class="detail-topbar-copy">
              <p class="section-kicker">Settings</p>
              <h1 id="settings-detail-title">${escapeHtml(TAB_META[tab].title)}</h1>
              <p class="section-copy">${escapeHtml(TAB_META[tab].description)}</p>
            </div>
          </div>
          <div class="detail-toolbar">
            <button type="button" class="secondary-btn" data-action="load-saved">Load Saved</button>
            <button type="button" class="secondary-btn" data-action="export-settings">Export</button>
            <button type="button" class="secondary-btn" data-action="restore-defaults">Restore Defaults</button>
            <button type="button" class="primary-btn" data-action="save-settings" ${validationError ? "disabled" : ""}>Save Settings</button>
          </div>
          <div class="detail-meta" aria-label="Settings state">
            <span class="state-pill ${dirty ? "dirty-pill" : "saved-pill"}">${dirty ? "Unsaved changes" : "Saved draft"}</span>
            <span class="state-pill">${escapeHtml(getSectionSummary(state, tab))}</span>
            ${validationError ? `<span class="state-pill warning-pill">${escapeHtml(validationError)}</span>` : ""}
          </div>
          ${renderDetailContent(state, tab)}
        </div>
      </section>
    </div>
  `;
}

function renderWorkspace(state: SettingsWorkspaceState): string {
  normalizeSelection(state);
  return state.currentView === "index"
    ? renderIndex(state)
    : renderDetail(state, state.currentView);
}

export function initSettingsWorkspace(options: SettingsWorkspaceOptions): void {
  const root = document.getElementById("settings-workspace");
  if (!(root instanceof HTMLDivElement)) {
    return;
  }

  const state: SettingsWorkspaceState = {
    draft: cloneResolvedFormatSettings(options.initialSettings),
    savedSnapshot: cloneResolvedFormatSettings(options.initialSettings),
    currentView: "index",
    selectedIndex: {
      numberFormats: null,
      percentFormats: null,
      dateFormats: null,
      cellFormats: null,
      textStyles: null,
    },
  };

  const render = (): void => {
    root.innerHTML = renderWorkspace(state);
  };

  initColorPicker(root);

  const runAsync = async (work: () => Promise<void>): Promise<void> => {
    try {
      await work();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      options.onStatus(`Error: ${message}`);
    }
  };

  root.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target.closest("[data-action]") : null;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const action = target.dataset.action;
    if (!action) {
      return;
    }

    syncDraftFromDom(root, state);

    if (action === "open-section") {
      const nextTab = target.dataset.tab as SettingsTabKey | undefined;
      if (nextTab) {
        state.currentView = nextTab;
        render();
      }
      return;
    }

    if (action === "go-home") {
      state.currentView = "index";
      render();
      return;
    }

    if (!isDetailView(state.currentView)) {
      return;
    }

    if (handleBorderAction(state, target)) {
      render();
      return;
    }

    if (action === "select-item" && isListTab(state.currentView)) {
      const index = Number(target.dataset.index ?? "0");
      const normalizedIndex = clampIndex(index, getListItems(state, state.currentView).length);
      state.selectedIndex[state.currentView] =
        state.selectedIndex[state.currentView] === normalizedIndex ? null : normalizedIndex;
      render();
      return;
    }

    if (action === "add-item" && isListTab(state.currentView)) {
      addPreset(state, state.currentView);
      options.onStatus(`${TAB_META[state.currentView].title.slice(0, -1)} preset added.`);
      render();
      return;
    }

    if (action === "move-item" && isListTab(state.currentView)) {
      const index = Number(target.dataset.index ?? "0");
      const direction = Number(target.dataset.direction ?? "0");
      movePreset(state, state.currentView, index, direction);
      render();
      return;
    }

    if (action === "delete-item" && isListTab(state.currentView)) {
      const index = Number(target.dataset.index ?? "0");
      deletePreset(state, state.currentView, index);
      render();
      return;
    }

    if (action === "load-saved") {
      void runAsync(async () => {
        const imported = await options.loadSavedSettings();
        state.draft = cloneResolvedFormatSettings(imported);
        normalizeSelection(state);
        options.onStatus("Settings file loaded. Click Save Settings to apply it to this workbook.");
        render();
      });
      return;
    }

    if (action === "export-settings") {
      void runAsync(async () => {
        await options.exportSettings(cloneResolvedFormatSettings(state.draft));
        options.onStatus("Settings exported.");
      });
      return;
    }

    if (action === "restore-defaults") {
      state.draft = buildDefaultFormatSettings();
      normalizeSelection(state);
      options.onStatus("Built-in defaults loaded into the editor. Save Settings to apply them.");
      render();
      return;
    }

    if (action === "save-settings") {
      void runAsync(async () => {
        const validationError = getFormatSettingsValidationError(state.draft);
        if (validationError) {
          options.onStatus(validationError);
          render();
          return;
        }

        const toSave = cloneResolvedFormatSettings(state.draft);
        await options.saveSettings(toSave);
        state.savedSnapshot = cloneResolvedFormatSettings(toSave);
        options.onStatus("Workbook settings saved. Text style cycle reset.");
        render();
      });
    }
  });

  root.addEventListener("change", (event) => {
    const target = event.target;
    if (
      !(
        target instanceof HTMLInputElement ||
        target instanceof HTMLSelectElement ||
        target instanceof HTMLTextAreaElement
      )
    ) {
      return;
    }

    const action = target.dataset.action;
    if (!action) {
      return;
    }

    if (action === "update-list-field") {
      updateSelectedListField(state, String(target.dataset.field ?? ""), parseEditorValue(target));
      render();
      return;
    }

    if (action === "update-palette") {
      const key = target.dataset.key as keyof AutoColorPalette | undefined;
      if (key) {
        state.draft.autoColorPalette[key] = target.value;
        render();
      }
      return;
    }

    if (action === "update-trace") {
      const key = target.dataset.key as keyof ResolvedFormatSettings["trace"] | undefined;
      if (key) {
        state.draft.trace[key] = target.value === "" ? NaN : Number(target.value);
        render();
      }
    }
  });

  render();
}
