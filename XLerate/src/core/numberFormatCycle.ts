export type NumberFormatDefinition = {
  name: string;
  formatCode: string;
};

export const DEFAULT_NUMBER_FORMATS: NumberFormatDefinition[] = [
  {
    name: "No Decimals",
    formatCode: "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"
  },
  {
    name: "1 Decimal",
    formatCode: "_(* #,##0.0_);_(* (#,##0.0);_(* \"-\"_);_(@_)"
  },
  {
    name: "2 Decimals",
    formatCode: "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"_);_(@_)"
  }
];

/**
 * Presets for the percent cycle (Ctrl+Shift+5). Same cycling rules as the
 * number cycle; kept as a separate list so the two shortcuts never fight
 * over the same format codes.
 */
export const DEFAULT_PERCENT_FORMATS: NumberFormatDefinition[] = [
  {
    name: "Percent, no decimals",
    formatCode: "0%_);(0%);\"-\"_)"
  },
  {
    name: "Percent, 1 decimal",
    formatCode: "0.0%_);(0.0%);\"-\"_)"
  },
  {
    name: "Percent, 2 decimals",
    formatCode: "0.00%_);(0.00%);\"-\"_)"
  }
];

export function hasMixedNumberFormats(formats: string[]): boolean {
  if (formats.length <= 1) {
    return false;
  }

  const first = formats[0];
  return formats.some((value) => value !== first);
}

export function computeNextNumberFormat(
  currentFormat: string,
  hasMixedSelection: boolean,
  configuredFormats: NumberFormatDefinition[] = DEFAULT_NUMBER_FORMATS
): string {
  if (configuredFormats.length === 0) {
    throw new Error("configuredFormats must contain at least one format");
  }

  if (hasMixedSelection) {
    return configuredFormats[0].formatCode;
  }

  const index = configuredFormats.findIndex((item) => item.formatCode === currentFormat);
  if (index < 0) {
    return configuredFormats[0].formatCode;
  }

  const nextIndex = index < configuredFormats.length - 1 ? index + 1 : 0;
  return configuredFormats[nextIndex].formatCode;
}
