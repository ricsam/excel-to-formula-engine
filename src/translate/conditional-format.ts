/**
 * Translates Excel conditional formatting into engine conditional styles.
 *
 * The engine offers two kinds of rule: a formula rule, which paints a cell one
 * colour when a formula is true, and a gradient rule, which interpolates
 * between two colour stops. Excel's rule catalogue is much larger, so this maps
 * what fits and reports the rest.
 *
 * Excel writes each rule's condition relative to the top-left cell of the range
 * it applies to, which is exactly the convention the engine's formula rules use,
 * so conditions carry over without rebasing.
 */

import {
  hexToLch,
  type ConditionalStyle,
  type LCHColor,
  type StyleCondition,
} from "@ricsam/formula-engine";
import type { RawConditionalFormat } from "../xlsx/worksheet";
import type { Diagnostic } from "../types";
import type { WorkbookDataArea } from "@ricsam/formula-engine";
import { rangeReferenceToSpreadsheetRange } from "../a1";
import { translateFormula } from "./formula";

export interface TranslatedConditionalFormat {
  areas: WorkbookDataArea[];
  condition: StyleCondition;
}

export interface TranslateConditionalFormatOptions {
  sheetName: string;
  /** Differential formats from styles.xml, indexed by `dxfId`. */
  differentialFormats: { backgroundColor?: string; color?: string }[];
}

export function translateConditionalFormat(
  rule: RawConditionalFormat,
  options: TranslateConditionalFormatOptions
): { result?: TranslatedConditionalFormat; diagnostics: Diagnostic[] } {
  const diagnostics: Diagnostic[] = [];

  const areas: WorkbookDataArea[] = [];
  for (const reference of rule.ranges) {
    const range = rangeReferenceToSpreadsheetRange(reference);
    if (range) {
      areas.push({ sheetName: options.sheetName, range });
    }
  }
  if (areas.length === 0) {
    return { diagnostics };
  }

  const anchor = rule.ranges[0]?.split(":")[0]?.replace(/\$/g, "") ?? "A1";

  if (rule.type === "colorScale") {
    return translateColorScale(rule, areas, options, diagnostics);
  }

  const condition = buildFormulaCondition(rule, anchor, options, diagnostics);
  if (!condition) {
    return { diagnostics };
  }

  return { result: { areas, condition }, diagnostics };
}

function buildFormulaCondition(
  rule: RawConditionalFormat,
  anchor: string,
  options: TranslateConditionalFormatOptions,
  diagnostics: Diagnostic[]
): StyleCondition | undefined {
  const color = ruleColor(rule, options);
  if (!color) {
    diagnostics.push({
      severity: "warning",
      code: "unsupported-conditional-format",
      message:
        `A ${rule.type} rule was skipped because its format has no fill or ` +
        `font colour the engine can apply.`,
      sheetName: options.sheetName,
    });
    return undefined;
  }

  const expression = buildExpression(rule, anchor, options, diagnostics);
  if (!expression) {
    return undefined;
  }

  const translated = translateFormula(expression, {
    sheetName: options.sheetName,
    cellReference: anchor,
  });
  diagnostics.push(...translated.diagnostics);

  return { type: "formula", formula: translated.formula, color };
}

/**
 * Build the boolean expression for a rule, in terms of the range's anchor cell.
 */
function buildExpression(
  rule: RawConditionalFormat,
  anchor: string,
  options: TranslateConditionalFormatOptions,
  diagnostics: Diagnostic[]
): string | undefined {
  const [first, second] = rule.formulas;

  switch (rule.type) {
    case "expression":
      return first ? `=${first}` : undefined;

    case "cellIs":
      return cellIsExpression(rule.operator, anchor, first, second);

    case "containsText":
      return rule.text === undefined
        ? undefined
        : `=IFERROR(FIND(${quote(rule.text)},${anchor})>0,FALSE)`;

    case "notContainsText":
      return rule.text === undefined
        ? undefined
        : `=IFERROR(FIND(${quote(rule.text)},${anchor})>0,FALSE)=FALSE`;

    case "beginsWith":
      return rule.text === undefined
        ? undefined
        : `=EXACT(LEFT(${anchor},${rule.text.length}),${quote(rule.text)})`;

    case "endsWith":
      return rule.text === undefined
        ? undefined
        : `=EXACT(RIGHT(${anchor},${rule.text.length}),${quote(rule.text)})`;

    case "containsBlanks":
      return `=LEN(${anchor})=0`;

    case "notContainsBlanks":
      return `=LEN(${anchor})>0`;

    default:
      diagnostics.push({
        severity: "warning",
        code: "unsupported-conditional-format",
        message:
          `The engine cannot express a "${rule.type}" conditional formatting ` +
          `rule, so it was skipped.`,
        sheetName: options.sheetName,
      });
      return undefined;
  }
}

function cellIsExpression(
  operator: string | undefined,
  anchor: string,
  first: string | undefined,
  second: string | undefined
): string | undefined {
  if (first === undefined) {
    return undefined;
  }

  switch (operator) {
    case "equal":
      return `=${anchor}=${first}`;
    case "notEqual":
      return `=${anchor}<>${first}`;
    case "greaterThan":
      return `=${anchor}>${first}`;
    case "greaterThanOrEqual":
      return `=${anchor}>=${first}`;
    case "lessThan":
      return `=${anchor}<${first}`;
    case "lessThanOrEqual":
      return `=${anchor}<=${first}`;
    case "between":
      return second === undefined
        ? undefined
        : `=AND(${anchor}>=${first},${anchor}<=${second})`;
    case "notBetween":
      return second === undefined
        ? undefined
        : `=OR(${anchor}<${first},${anchor}>${second})`;
    default:
      return undefined;
  }
}

function translateColorScale(
  rule: RawConditionalFormat,
  areas: WorkbookDataArea[],
  options: TranslateConditionalFormatOptions,
  diagnostics: Diagnostic[]
): { result?: TranslatedConditionalFormat; diagnostics: Diagnostic[] } {
  const scale = rule.colorScale;
  if (!scale || scale.cfvo.length < 2 || scale.colors.length < 2) {
    return { diagnostics };
  }

  if (scale.cfvo.length > 2) {
    diagnostics.push({
      severity: "info",
      code: "unsupported-conditional-format",
      message:
        "A three-colour scale was imported as a two-colour gradient between " +
        "its lowest and highest stops; the midpoint colour was dropped.",
      sheetName: options.sheetName,
    });
  }

  const firstStop = scale.cfvo[0]!;
  const lastStop = scale.cfvo[scale.cfvo.length - 1]!;
  const firstColor = toLch(scale.colors[0]);
  const lastColor = toLch(scale.colors[scale.colors.length - 1]);

  if (!firstColor || !lastColor) {
    return { diagnostics };
  }

  const min =
    firstStop.type === "num" && firstStop.val !== undefined
      ? ({
          type: "number" as const,
          color: firstColor,
          valueFormula: `=${firstStop.val}`,
        })
      : ({ type: "lowest_value" as const, color: firstColor });

  const max =
    lastStop.type === "num" && lastStop.val !== undefined
      ? ({
          type: "number" as const,
          color: lastColor,
          valueFormula: `=${lastStop.val}`,
        })
      : ({ type: "highest_value" as const, color: lastColor });

  if (
    (firstStop.type !== "num" && firstStop.type !== "min") ||
    (lastStop.type !== "num" && lastStop.type !== "max")
  ) {
    diagnostics.push({
      severity: "info",
      code: "unsupported-conditional-format",
      message:
        `A colour scale using "${firstStop.type}"/"${lastStop.type}" stops was ` +
        "imported as a gradient across the range's own lowest and highest " +
        "values.",
      sheetName: options.sheetName,
    });
  }

  return {
    result: { areas, condition: { type: "gradient", min, max } },
    diagnostics,
  };
}

/**
 * The colour a rule paints. The engine's formula rule takes a single colour, so
 * a differential format's fill wins over its font colour.
 */
function ruleColor(
  rule: RawConditionalFormat,
  options: TranslateConditionalFormatOptions
): LCHColor | undefined {
  if (rule.dxfId === undefined) {
    return undefined;
  }
  const format = options.differentialFormats[rule.dxfId];
  if (!format) {
    return undefined;
  }
  const hex = format.backgroundColor ?? format.color;
  return hex ? hexToLch(hex) : undefined;
}

function toLch(argb: string | undefined): LCHColor | undefined {
  if (!argb) {
    return undefined;
  }
  const value = argb.replace(/^#/, "");
  const rgb = value.length === 8 ? value.slice(2) : value;
  if (!/^[0-9A-Fa-f]{6}$/.test(rgb)) {
    return undefined;
  }
  return hexToLch(`#${rgb.toLowerCase()}`);
}

function quote(text: string): string {
  return `"${text.replace(/"/g, '""')}"`;
}

export type { ConditionalStyle };
