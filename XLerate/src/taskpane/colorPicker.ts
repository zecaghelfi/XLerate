/**
 * Excel-style color picker for the settings panel: Theme Colors (from the
 * open workbook's theme, with Excel's tint/shade rows), Standard Colors,
 * Recent Colors and More Colors (native picker or hex).
 *
 * Each control is a hidden <input> carrying the settings data-attributes plus
 * a swatch button. Picking a color writes the hidden input and dispatches a
 * bubbling "change" event, so the settings workspace handles it exactly like
 * the old <input type="color">.
 */

import {
  buildThemeGrid,
  EXCEL_STANDARD_COLORS,
  normalizeHex,
  pushRecentColor,
  type NamedColor,
} from "../core/excelColors";
import { currentThemeOrDefault, refreshWorkbookTheme, type LoadedTheme } from "./themeColors";

const RECENT_KEY = "xlerate_recent_colors_v1";

let controlSequence = 0;
let popover: HTMLDivElement | null = null;
let activeInput: HTMLInputElement | null = null;
let activeButton: HTMLElement | null = null;
let themeLoading = false;

function escapeAttr(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function readRecent(): string[] {
  try {
    const raw = window.localStorage.getItem(RECENT_KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed)
      ? parsed.filter((v): v is string => typeof v === "string" && normalizeHex(v) !== null)
      : [];
  } catch {
    return [];
  }
}

function writeRecent(list: string[]): void {
  try {
    window.localStorage.setItem(RECENT_KEY, JSON.stringify(list));
  } catch {
    // storage unavailable: recent colors stay session-less
  }
}

/**
 * HTML for one color control. `dataAttrs` are the attributes the settings
 * workspace reads (e.g. `data-action="update-list-field" data-field="fillColor"`).
 */
export function renderColorControl(options: {
  dataAttrs: string;
  value: string;
  disabled?: boolean;
  label?: string;
}): string {
  const id = `xl-color-${++controlSequence}`;
  const hex = normalizeHex(options.value) ?? "#000000";
  const label = options.label ?? "Choose color";
  return `
    <span class="xl-color${options.disabled ? " xl-color-disabled" : ""}">
      <input type="hidden" id="${id}" ${options.dataAttrs} value="${hex}" />
      <button type="button" class="xl-color-button" data-color-target="${id}" aria-haspopup="dialog" aria-label="${escapeAttr(label)}: ${hex}" ${options.disabled ? "disabled" : ""}>
        <span class="xl-color-swatch" style="background:${hex}"></span>
        <span class="xl-color-hex">${hex}</span>
        <span class="xl-color-caret" aria-hidden="true">▾</span>
      </button>
    </span>
  `;
}

function swatch(color: NamedColor, current: string | null, extraClass = ""): string {
  const selected = current !== null && normalizeHex(color.hex) === current;
  return `<button type="button" class="xlc-swatch${extraClass}${selected ? " xlc-selected" : ""}" data-xlc-color="${color.hex}" title="${escapeAttr(color.label)} (${color.hex})" aria-label="${escapeAttr(color.label)}" style="background:${color.hex}"></button>`;
}

function renderPopover(theme: LoadedTheme, current: string | null): string {
  const grid = buildThemeGrid(theme);
  const recent = readRecent();
  const themeLabel =
    theme.source === "workbook" ? `${theme.name} (this workbook)` : `${theme.name} (default)`;
  return `
    <div class="xlc-header">
      <span class="xlc-title">Theme Colors</span>
      <span class="xlc-theme-name" title="${escapeAttr(themeLabel)}">${escapeAttr(themeLabel)}${themeLoading ? " · reading…" : ""}</span>
      <button type="button" class="xlc-icon-btn" data-xlc="reload" title="Reload workbook theme" aria-label="Reload workbook theme">↻</button>
    </div>
    <div class="xlc-grid xlc-grid-base">${grid[0].map((c) => swatch(c, current)).join("")}</div>
    <div class="xlc-grid xlc-grid-tints">${grid
      .slice(1)
      .map((row) => row.map((c) => swatch(c, current, " xlc-tint")).join(""))
      .join("")}</div>
    <div class="xlc-header"><span class="xlc-title">Standard Colors</span></div>
    <div class="xlc-grid">${EXCEL_STANDARD_COLORS.map((c) => swatch(c, current)).join("")}</div>
    ${
      recent.length > 0
        ? `<div class="xlc-header"><span class="xlc-title">Recent Colors</span></div>
           <div class="xlc-grid">${recent.map((hex) => swatch({ hex, label: hex }, current)).join("")}</div>`
        : ""
    }
    <div class="xlc-more">
      <label class="xlc-more-native" title="More Colors…">
        <span>More Colors…</span>
        <input type="color" data-xlc="native" value="${current ?? "#000000"}" />
      </label>
      <input type="text" class="xlc-hex" data-xlc="hex" maxlength="7" placeholder="#RRGGBB" value="${current ?? ""}" aria-label="Hex color" />
      <button type="button" class="xlc-ok" data-xlc="apply-hex">OK</button>
    </div>
  `;
}

function position(): void {
  if (!popover || !activeButton) return;
  const rect = activeButton.getBoundingClientRect();
  const width = popover.offsetWidth;
  const height = popover.offsetHeight;
  const margin = 6;
  let left = Math.min(rect.left, window.innerWidth - width - margin);
  left = Math.max(margin, left);
  let top = rect.bottom + 4;
  if (top + height > window.innerHeight - margin && rect.top - height - 4 > margin) {
    top = rect.top - height - 4;
  }
  popover.style.left = `${left}px`;
  popover.style.top = `${Math.max(margin, top)}px`;
}

function rerender(): void {
  if (!popover || !activeInput) return;
  popover.innerHTML = renderPopover(currentThemeOrDefault(), normalizeHex(activeInput.value));
  position();
}

function close(): void {
  popover?.remove();
  popover = null;
  activeInput = null;
  activeButton?.focus();
  activeButton = null;
}

function choose(hex: string): void {
  const value = normalizeHex(hex);
  const input = activeInput;
  if (!value || !input) return;
  writeRecent(pushRecentColor(readRecent(), value));
  close();
  input.value = value;
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

function loadTheme(force: boolean): void {
  themeLoading = true;
  rerender();
  void refreshWorkbookTheme(force).finally(() => {
    themeLoading = false;
    rerender();
  });
}

function open(button: HTMLElement): void {
  const targetId = button.dataset.colorTarget;
  const input = targetId ? document.getElementById(targetId) : null;
  if (!(input instanceof HTMLInputElement)) return;
  if (popover && activeInput === input) {
    close();
    return;
  }
  close();
  activeInput = input;
  activeButton = button;
  popover = document.createElement("div");
  popover.className = "xlc-popover";
  popover.setAttribute("role", "dialog");
  popover.setAttribute("aria-label", "Color picker");
  document.body.appendChild(popover);

  popover.addEventListener("click", (event) => {
    const el = event.target instanceof Element ? event.target.closest<HTMLElement>("[data-xlc-color], [data-xlc]") : null;
    if (!el) return;
    if (el.dataset.xlcColor) {
      choose(el.dataset.xlcColor);
      return;
    }
    if (el.dataset.xlc === "reload") {
      loadTheme(true);
      return;
    }
    if (el.dataset.xlc === "apply-hex") {
      const hexInput = popover?.querySelector<HTMLInputElement>("[data-xlc='hex']");
      if (hexInput) choose(hexInput.value);
    }
  });
  popover.addEventListener("change", (event) => {
    const el = event.target;
    if (el instanceof HTMLInputElement && el.dataset.xlc === "native") choose(el.value);
  });
  popover.addEventListener("keydown", (event) => {
    const el = event.target;
    if (event.key === "Enter" && el instanceof HTMLInputElement && el.dataset.xlc === "hex") {
      event.preventDefault();
      choose(el.value);
    }
  });

  rerender();
  loadTheme(false);
}

/** Wires the color controls inside `root`. Call once. */
export function initColorPicker(root: HTMLElement): void {
  root.addEventListener("click", (event) => {
    const button =
      event.target instanceof Element ? event.target.closest<HTMLElement>(".xl-color-button") : null;
    if (button && !(button as HTMLButtonElement).disabled) {
      event.preventDefault();
      open(button);
    }
  });

  document.addEventListener(
    "mousedown",
    (event) => {
      if (!popover) return;
      const target = event.target instanceof Node ? event.target : null;
      if (target && (popover.contains(target) || activeButton?.contains(target))) return;
      close();
    },
    true
  );
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && popover) close();
  });
  window.addEventListener("resize", () => position());
  window.addEventListener(
    "scroll",
    (event) => {
      if (popover && !(event.target instanceof Node && popover.contains(event.target))) close();
    },
    true
  );
}
