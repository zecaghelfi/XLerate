/* global Office */

/**
 * Loads the active workbook's theme colors so the settings color picker can
 * show the same Theme Colors grid as Excel.
 *
 * Office.js has no theme API, so we read the document itself
 * (getFileAsync, compressed = the .xlsx zip) and parse xl/theme/theme1.xml.
 * The result is cached for the session; if anything fails (unsupported host,
 * very large file, timeout) we fall back to Excel's default Office theme.
 */

import { unzipSync } from "fflate";
import { OFFICE_DEFAULT_THEME, parseThemeXml, type ThemeColors } from "../core/excelColors";

export type LoadedTheme = ThemeColors & { source: "workbook" | "default" };

const SLICE_SIZE = 4 * 1024 * 1024;
const MAX_FILE_BYTES = 80 * 1024 * 1024;
const TIMEOUT_MS = 20000;

let cached: Promise<LoadedTheme> | null = null;

function defaultTheme(): LoadedTheme {
  return { ...OFFICE_DEFAULT_THEME, colors: [...OFFICE_DEFAULT_THEME.colors], source: "default" };
}

function getFile(): Promise<Office.File> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: SLICE_SIZE },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error.message));
      }
    );
  });
}

function getSlice(file: Office.File, index: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(new Uint8Array(result.value.data as number[]));
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

function closeFile(file: Office.File): void {
  try {
    file.closeAsync();
  } catch {
    // ignore: the file handle is released when the runtime is torn down
  }
}

async function readWorkbookBytes(): Promise<Uint8Array> {
  const file = await getFile();
  try {
    if (file.size > MAX_FILE_BYTES) {
      throw new Error("Workbook too large to read the theme");
    }
    const parts: Uint8Array[] = [];
    for (let i = 0; i < file.sliceCount; i++) {
      parts.push(await getSlice(file, i));
    }
    const total = parts.reduce((n, p) => n + p.length, 0);
    const out = new Uint8Array(total);
    let offset = 0;
    for (const p of parts) {
      out.set(p, offset);
      offset += p.length;
    }
    return out;
  } finally {
    closeFile(file);
  }
}

async function loadFromWorkbook(): Promise<LoadedTheme> {
  const bytes = await readWorkbookBytes();
  const files = unzipSync(bytes, { filter: (f) => /^xl\/theme\/theme1\.xml$/i.test(f.name) });
  const entry = Object.values(files)[0];
  if (!entry) throw new Error("Theme part not found");
  const parsed = parseThemeXml(new TextDecoder("utf-8").decode(entry));
  if (!parsed) throw new Error("Theme colors not found");
  return { ...parsed, source: "workbook" };
}

function withTimeout<T>(work: Promise<T>, ms: number): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error("Timed out reading the theme")), ms);
    work.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (error) => {
        clearTimeout(timer);
        reject(error);
      }
    );
  });
}

/** Theme colors of the open workbook (cached). Never rejects. */
export function getWorkbookTheme(forceReload = false): Promise<LoadedTheme> {
  if (!cached || forceReload) {
    cached = withTimeout(loadFromWorkbook(), TIMEOUT_MS).catch((error) => {
      // eslint-disable-next-line no-console
      console.warn("[XLerate theme] using the default Office theme:", error);
      return defaultTheme();
    });
  }
  return cached;
}

/** Synchronous peek: the loaded theme if ready, otherwise the default. */
let lastResolved: LoadedTheme = defaultTheme();
export function currentThemeOrDefault(): LoadedTheme {
  return lastResolved;
}

export async function refreshWorkbookTheme(forceReload = false): Promise<LoadedTheme> {
  lastResolved = await getWorkbookTheme(forceReload);
  return lastResolved;
}
