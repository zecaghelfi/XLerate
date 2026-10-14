import { ExcelPort } from "../adapters/excelPort";
import {
  computeNextNumberFormat,
  DEFAULT_NUMBER_FORMATS,
  hasMixedNumberFormats,
  NumberFormatDefinition,
} from "../core/numberFormatCycle";

/**
 * Implements spec §3.7. Reads the current number format from the selection,
 * asks the core which preset comes next, and applies it to every selected
 * cell in one batch.
 */
export async function runCycleNumberFormat(
  port: ExcelPort,
  configuredFormats: NumberFormatDefinition[] = DEFAULT_NUMBER_FORMATS
): Promise<void> {
  // Only number formats are read (no fonts/fills/borders), and the new code
  // is written to the whole selection in one call.
  const summary = await port.getSelectionNumberFormats();
  if (!summary) return;

  const mixed = hasMixedNumberFormats(summary.formats);
  const nextFormat = computeNextNumberFormat(summary.formats[0], mixed, configuredFormats);
  await port.applySelectionNumberFormat(nextFormat);
}
