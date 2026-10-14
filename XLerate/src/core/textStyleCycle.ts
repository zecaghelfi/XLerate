import {
  resolveBorderEdges,
  TEXT_STYLE_DEFAULT_EDGES,
  type BorderEdgeFlags,
  type BorderLineStyle,
  type BorderWeight,
} from "./borders";

export type TextStyleFillPattern = "Solid" | "None";

export type TextStyleDefinition = {
  name: string;
  fontName: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  fontColor: string;
  fillPattern: TextStyleFillPattern;
  backColor: string;
  /** Any Excel line style. */
  borderStyle: BorderLineStyle;
  /** Excel border weight. Missing = legacy mapping from the style (mapBorderWeight). */
  borderWeight?: BorderWeight;
  /** Border color. Missing = font color (legacy behavior). */
  borderColor?: string;
  borderTop: boolean;
  borderBottom: boolean;
  borderLeft: boolean;
  borderRight: boolean;
  borderInsideHorizontal?: boolean;
  borderInsideVertical?: boolean;
  borderDiagonalDown?: boolean;
  borderDiagonalUp?: boolean;
};

export function textStyleEdges(style: TextStyleDefinition): BorderEdgeFlags {
  return resolveBorderEdges(style, TEXT_STYLE_DEFAULT_EDGES);
}

export function textStyleBorderWeight(style: TextStyleDefinition): BorderWeight {
  return style.borderWeight ?? mapBorderWeight(style.borderStyle);
}

export const DEFAULT_TEXT_STYLES: TextStyleDefinition[] = [
  {
    name: "Heading",
    fontName: "Calibri",
    fontSize: 14,
    bold: true,
    italic: false,
    underline: false,
    fontColor: "#000000",
    fillPattern: "Solid",
    backColor: "#F0F0F0",
    borderStyle: "Continuous",
    borderTop: true,
    borderBottom: true,
    borderLeft: false,
    borderRight: false
  },
  {
    name: "Subheading",
    fontName: "Calibri",
    fontSize: 12,
    bold: true,
    italic: false,
    underline: false,
    fontColor: "#595959",
    fillPattern: "Solid",
    backColor: "#F5F5F5",
    borderStyle: "Continuous",
    borderTop: false,
    borderBottom: true,
    borderLeft: false,
    borderRight: false
  },
  {
    name: "Sum",
    fontName: "Calibri",
    fontSize: 11,
    bold: true,
    italic: false,
    underline: true,
    fontColor: "#000000",
    fillPattern: "Solid",
    backColor: "#FFFFFF",
    borderStyle: "Double",
    borderTop: true,
    borderBottom: false,
    borderLeft: false,
    borderRight: false
  },
  {
    name: "Normal",
    fontName: "Calibri",
    fontSize: 11,
    bold: false,
    italic: false,
    underline: false,
    fontColor: "#000000",
    fillPattern: "Solid",
    backColor: "#FFFFFF",
    borderStyle: "None",
    borderTop: false,
    borderBottom: false,
    borderLeft: false,
    borderRight: false
  }
];

export function computeNextTextStyleIndex(currentIndex: number, styles: TextStyleDefinition[] = DEFAULT_TEXT_STYLES): number {
  if (styles.length === 0) {
    throw new Error("styles must contain at least one item");
  }
  return (currentIndex + 1) % styles.length;
}

export function computeNextTextStyle(
  currentIndex: number,
  styles: TextStyleDefinition[] = DEFAULT_TEXT_STYLES
): { index: number; style: TextStyleDefinition } {
  const index = computeNextTextStyleIndex(currentIndex, styles);
  return { index, style: styles[index] };
}

export function mapBorderWeight(style: BorderLineStyle): "Thin" | "Medium" | "Thick" {
  if (style === "Continuous") {
    return "Medium";
  }
  if (style === "Double") {
    return "Thick";
  }
  if (style === "Dash" || style === "Dot" || style === "DashDot" || style === "DashDotDot") {
    return "Thin";
  }
  if (style === "SlantDashDot") {
    return "Medium";
  }
  return "Thin";
}
