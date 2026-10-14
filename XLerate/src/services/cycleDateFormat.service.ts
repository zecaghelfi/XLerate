import { ExcelPort } from "../adapters/excelPort";
import {
  computeNextDateFormat,
  DateFormatDefinition,
  DEFAULT_DATE_FORMATS,
  hasMixedDateFormats,
} from "../core/dateFormatCycle";

/**
 * Implements spec §3.9. Same shape as runCycleNumberFormat but using the
 * date-format preset list.
 */
export async function runCycleDateFormat(
  port: ExcelPort,
  configuredFormats: DateFormatDefinition[] = DEFAULT_DATE_FORMATS
): Promise<void> {
  const summary = await port.getSelectionNumberFormats();
  if (!summary) return;

  const mixed = hasMixedDateFormats(summary.formats);
  const nextFormat = computeNextDateFormat(summary.formats[0], mixed, configuredFormats);
  await port.applySelectionNumberFormat(nextFormat);
}
