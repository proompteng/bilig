import { builtinCapabilityManifest, type BuiltinCapabilityCategory } from "@bilig/formula";
import type {
  WorkbookDefinedNameSnapshot,
  WorkbookDefinedNameValueSnapshot,
} from "@bilig/protocol";

export interface FormulaHelpArg {
  readonly label: string;
  readonly optional?: boolean;
}

export interface FormulaHelpEntry {
  readonly kind: "function";
  readonly name: string;
  readonly category: BuiltinCapabilityCategory;
  readonly summary: string;
  readonly args: readonly FormulaHelpArg[];
  readonly variadic?: boolean;
}

export interface DefinedNameSuggestion {
  readonly kind: "defined-name";
  readonly name: string;
  readonly summary: string;
  readonly insertText: string;
}

export interface FunctionSuggestion {
  readonly kind: "function";
  readonly name: string;
  readonly category: BuiltinCapabilityCategory;
  readonly summary: string;
  readonly signature: string;
}

export type FormulaSuggestion = FunctionSuggestion | DefinedNameSuggestion;

export interface FormulaReplaceResult {
  readonly value: string;
  readonly caret: number;
}

export interface FormulaAssistState {
  readonly tokenStart: number | null;
  readonly tokenEnd: number | null;
  readonly suggestions: readonly FormulaSuggestion[];
  readonly activeFunction: {
    readonly entry: FormulaHelpEntry;
    readonly activeArgumentIndex: number;
    readonly signature: string;
  } | null;
}

const IDENTIFIER_PATTERN = /[A-Za-z0-9_.]/;
const COMMON_FUNCTIONS = new Set([
  "SUM",
  "AVERAGE",
  "COUNT",
  "COUNTA",
  "MIN",
  "MAX",
  "IF",
  "IFERROR",
  "IFNA",
  "AND",
  "OR",
  "NOT",
  "XLOOKUP",
  "XMATCH",
  "VLOOKUP",
  "HLOOKUP",
  "INDEX",
  "MATCH",
  "SUMIF",
  "SUMIFS",
  "COUNTIF",
  "COUNTIFS",
  "TEXTJOIN",
  "CONCAT",
  "FILTER",
  "UNIQUE",
  "SORT",
  "ROUND",
  "DATE",
  "TODAY",
  "NOW",
]);

const FALLBACK_ARGS: readonly FormulaHelpArg[] = [
  { label: "value1" },
  { label: "[value2]", optional: true },
  { label: "…", optional: true },
];

const CURATED_HELP: Record<
  string,
  {
    readonly summary: string;
    readonly args: readonly FormulaHelpArg[];
    readonly variadic?: boolean;
  }
> = {
  SUM: {
    summary: "Add numbers, ranges, and spill results.",
    args: [{ label: "number1" }, { label: "[number2]", optional: true }],
    variadic: true,
  },
  AVERAGE: {
    summary: "Return the arithmetic mean for the supplied values.",
    args: [{ label: "number1" }, { label: "[number2]", optional: true }],
    variadic: true,
  },
  COUNT: {
    summary: "Count numeric values in cells, ranges, and arguments.",
    args: [{ label: "value1" }, { label: "[value2]", optional: true }],
    variadic: true,
  },
  COUNTA: {
    summary: "Count non-empty values in cells, ranges, and arguments.",
    args: [{ label: "value1" }, { label: "[value2]", optional: true }],
    variadic: true,
  },
  MIN: {
    summary: "Return the smallest numeric value from the supplied inputs.",
    args: [{ label: "number1" }, { label: "[number2]", optional: true }],
    variadic: true,
  },
  MAX: {
    summary: "Return the largest numeric value from the supplied inputs.",
    args: [{ label: "number1" }, { label: "[number2]", optional: true }],
    variadic: true,
  },
  IF: {
    summary: "Choose between two results based on a logical test.",
    args: [
      { label: "logical_test" },
      { label: "value_if_true" },
      { label: "[value_if_false]", optional: true },
    ],
  },
  IFERROR: {
    summary: "Replace an error result with a fallback value.",
    args: [{ label: "value" }, { label: "value_if_error" }],
  },
  IFNA: {
    summary: "Replace a #N/A result with a fallback value.",
    args: [{ label: "value" }, { label: "value_if_na" }],
  },
  AND: {
    summary: "Return TRUE only when every argument is truthy.",
    args: [{ label: "logical1" }, { label: "[logical2]", optional: true }],
    variadic: true,
  },
  OR: {
    summary: "Return TRUE when any supplied argument is truthy.",
    args: [{ label: "logical1" }, { label: "[logical2]", optional: true }],
    variadic: true,
  },
  NOT: {
    summary: "Invert a logical value.",
    args: [{ label: "logical" }],
  },
  XLOOKUP: {
    summary: "Look up a value in one array and return the matching item from another.",
    args: [
      { label: "lookup_value" },
      { label: "lookup_array" },
      { label: "return_array" },
      { label: "[if_not_found]", optional: true },
      { label: "[match_mode]", optional: true },
      { label: "[search_mode]", optional: true },
    ],
  },
  XMATCH: {
    summary: "Return the relative position of a lookup value inside an array.",
    args: [
      { label: "lookup_value" },
      { label: "lookup_array" },
      { label: "[match_mode]", optional: true },
      { label: "[search_mode]", optional: true },
    ],
  },
  VLOOKUP: {
    summary: "Search the first column of a table and return a value from a target column.",
    args: [
      { label: "lookup_value" },
      { label: "table_array" },
      { label: "col_index_num" },
      { label: "[range_lookup]", optional: true },
    ],
  },
  HLOOKUP: {
    summary: "Search the first row of a table and return a value from a target row.",
    args: [
      { label: "lookup_value" },
      { label: "table_array" },
      { label: "row_index_num" },
      { label: "[range_lookup]", optional: true },
    ],
  },
  INDEX: {
    summary: "Return a value or reference at the given row and column position.",
    args: [{ label: "array" }, { label: "row_num" }, { label: "[column_num]", optional: true }],
  },
  MATCH: {
    summary: "Return the relative position of a lookup value in a one-dimensional range.",
    args: [
      { label: "lookup_value" },
      { label: "lookup_array" },
      { label: "[match_type]", optional: true },
    ],
  },
  SUMIF: {
    summary: "Sum cells that match a single condition.",
    args: [{ label: "range" }, { label: "criteria" }, { label: "[sum_range]", optional: true }],
  },
  SUMIFS: {
    summary: "Sum cells that match multiple conditions.",
    args: [
      { label: "sum_range" },
      { label: "criteria_range1" },
      { label: "criteria1" },
      { label: "[criteria_range2]", optional: true },
      { label: "[criteria2]", optional: true },
    ],
    variadic: true,
  },
  COUNTIF: {
    summary: "Count cells that match a single condition.",
    args: [{ label: "range" }, { label: "criteria" }],
  },
  COUNTIFS: {
    summary: "Count cells that match multiple conditions.",
    args: [
      { label: "criteria_range1" },
      { label: "criteria1" },
      { label: "[criteria_range2]", optional: true },
      { label: "[criteria2]", optional: true },
    ],
    variadic: true,
  },
  TEXTJOIN: {
    summary: "Join multiple values with a delimiter and optional blank skipping.",
    args: [
      { label: "delimiter" },
      { label: "ignore_empty" },
      { label: "text1" },
      { label: "[text2]", optional: true },
    ],
    variadic: true,
  },
  CONCAT: {
    summary: "Concatenate text values without a delimiter.",
    args: [{ label: "text1" }, { label: "[text2]", optional: true }],
    variadic: true,
  },
  FILTER: {
    summary: "Return rows or columns that satisfy a filter condition.",
    args: [{ label: "array" }, { label: "include" }, { label: "[if_empty]", optional: true }],
  },
  UNIQUE: {
    summary: "Return unique rows or columns from an array.",
    args: [
      { label: "array" },
      { label: "[by_col]", optional: true },
      { label: "[exactly_once]", optional: true },
    ],
  },
  SORT: {
    summary: "Sort rows or columns in ascending or descending order.",
    args: [
      { label: "array" },
      { label: "[sort_index]", optional: true },
      { label: "[sort_order]", optional: true },
      { label: "[by_col]", optional: true },
    ],
  },
  ROUND: {
    summary: "Round a number to the specified number of digits.",
    args: [{ label: "number" }, { label: "num_digits" }],
  },
  ROUNDUP: {
    summary: "Round a number away from zero.",
    args: [{ label: "number" }, { label: "num_digits" }],
  },
  ROUNDDOWN: {
    summary: "Round a number toward zero.",
    args: [{ label: "number" }, { label: "num_digits" }],
  },
  LEFT: {
    summary: "Return the leftmost characters from a text value.",
    args: [{ label: "text" }, { label: "[num_chars]", optional: true }],
  },
  RIGHT: {
    summary: "Return the rightmost characters from a text value.",
    args: [{ label: "text" }, { label: "[num_chars]", optional: true }],
  },
  MID: {
    summary: "Return a substring from the middle of a text value.",
    args: [{ label: "text" }, { label: "start_num" }, { label: "num_chars" }],
  },
  DATE: {
    summary: "Build a date serial from year, month, and day components.",
    args: [{ label: "year" }, { label: "month" }, { label: "day" }],
  },
  TODAY: {
    summary: "Return the current date in the workbook volatile context.",
    args: [],
  },
  NOW: {
    summary: "Return the current date and time in the workbook volatile context.",
    args: [],
  },
  YEAR: {
    summary: "Extract the year from a date serial.",
    args: [{ label: "serial_number" }],
  },
  MONTH: {
    summary: "Extract the month from a date serial.",
    args: [{ label: "serial_number" }],
  },
  DAY: {
    summary: "Extract the day of month from a date serial.",
    args: [{ label: "serial_number" }],
  },
};

const CATEGORY_SUMMARIES: Record<BuiltinCapabilityCategory, string> = {
  aggregation: "Combine values across ranges and arrays.",
  logical: "Branch or test logical conditions.",
  information: "Inspect value types, sheet state, or cell metadata.",
  text: "Manipulate and search text values.",
  "date-time": "Compute or extract date and time values.",
  "lookup-reference": "Look up positions, references, or matching values.",
  statistical: "Compute descriptive and inferential statistics.",
  "dynamic-array": "Produce or transform spill-aware array results.",
  lambda: "Compose spreadsheet functions with lambda semantics.",
  math: "Perform scalar math, rounding, and numeric transforms.",
};

const functionHelpEntries: readonly FormulaHelpEntry[] = builtinCapabilityManifest
  .map((capability) => {
    const curated = CURATED_HELP[capability.name];
    return Object.assign(
      {
        kind: `function` as const,
        name: capability.name,
        category: capability.category,
        summary: curated?.summary ?? CATEGORY_SUMMARIES[capability.category],
        args: curated?.args ?? FALLBACK_ARGS,
      },
      curated?.variadic ? { variadic: true } : !curated ? { variadic: true } : {},
    );
  })
  .toSorted((left, right) => left.name.localeCompare(right.name));

const functionHelpByName = new Map(functionHelpEntries.map((entry) => [entry.name, entry]));

function isIdentifierCharacter(char: string | undefined): boolean {
  return char !== undefined && IDENTIFIER_PATTERN.test(char);
}

function isEscapedDoubleQuote(source: string, index: number): boolean {
  return source[index] === '"' && source[index + 1] === '"';
}

function isEscapedSingleQuote(source: string, index: number): boolean {
  return source[index] === "'" && source[index + 1] === "'";
}

function previousNonWhitespaceChar(source: string, index: number): string | null {
  for (let cursor = index; cursor >= 0; cursor -= 1) {
    const char = source[cursor];
    if (char && !/\s/.test(char)) {
      return char;
    }
  }
  return null;
}

function findCalleeBeforeParen(source: string, parenIndex: number): string | null {
  let cursor = parenIndex - 1;
  while (cursor >= 0 && /\s/.test(source[cursor]!)) {
    cursor -= 1;
  }
  const end = cursor + 1;
  while (cursor >= 0 && isIdentifierCharacter(source[cursor])) {
    cursor -= 1;
  }
  const start = cursor + 1;
  if (start >= end) {
    return null;
  }
  const name = source.slice(start, end).toUpperCase();
  const previous = previousNonWhitespaceChar(source, start - 1);
  if (previous === "!" || previous === "]" || previous === "#") {
    return null;
  }
  return name;
}

function formatArgumentLabel(arg: FormulaHelpArg): string {
  return arg.optional ? `[${arg.label.replace(/^\[(.*)\]$/, "$1")}]` : arg.label;
}

export function formatFormulaSignature(entry: FormulaHelpEntry): string {
  if (entry.args.length === 0) {
    return `${entry.name}()`;
  }
  const labels = entry.args.map(formatArgumentLabel);
  if (entry.variadic && entry.args.at(-1)?.label !== "…") {
    labels.push("…");
  }
  return `${entry.name}(${labels.join(", ")})`;
}

function formatDefinedNameSummary(value: WorkbookDefinedNameValueSnapshot): string {
  if (value === null) {
    return "Defined name = blank";
  }
  if (typeof value === "string") {
    return `Defined name = "${value}"`;
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return `Defined name = ${String(value)}`;
  }
  switch (value.kind) {
    case "scalar":
      return `Defined name = ${String(value.value)}`;
    case "cell-ref":
      return `${value.sheetName}!${value.address}`;
    case "range-ref":
      return `${value.sheetName}!${value.startAddress}:${value.endAddress}`;
    case "structured-ref":
      return `${value.tableName}[${value.columnName}]`;
    case "formula":
      return `=${value.formula}`;
  }
}

export function resolveNameBoxDisplayValue(input: {
  readonly sheetName: string;
  readonly address: string;
  readonly definedNames?: readonly WorkbookDefinedNameSnapshot[];
}): string {
  const match = input.definedNames?.find(
    (entry) =>
      entry.value &&
      typeof entry.value === "object" &&
      "kind" in entry.value &&
      entry.value.kind === "cell-ref" &&
      entry.value.sheetName === input.sheetName &&
      entry.value.address.toUpperCase() === input.address.toUpperCase(),
  );
  return match?.name ?? input.address;
}

function buildDefinedNameSuggestions(
  prefix: string,
  definedNames: readonly WorkbookDefinedNameSnapshot[],
): readonly DefinedNameSuggestion[] {
  return definedNames
    .filter((entry) => entry.name.toUpperCase().startsWith(prefix))
    .toSorted((left, right) => left.name.localeCompare(right.name))
    .map((entry) => ({
      kind: "defined-name" as const,
      name: entry.name,
      summary: formatDefinedNameSummary(entry.value),
      insertText: entry.name,
    }));
}

function buildFunctionSuggestions(prefix: string): readonly FunctionSuggestion[] {
  const candidates =
    prefix.length === 0
      ? functionHelpEntries.filter((entry) => COMMON_FUNCTIONS.has(entry.name))
      : functionHelpEntries.filter((entry) => entry.name.startsWith(prefix));
  return candidates
    .toSorted((left, right) => {
      const leftCommon = COMMON_FUNCTIONS.has(left.name) ? 0 : 1;
      const rightCommon = COMMON_FUNCTIONS.has(right.name) ? 0 : 1;
      if (leftCommon !== rightCommon) {
        return leftCommon - rightCommon;
      }
      return left.name.localeCompare(right.name);
    })
    .map((entry) => ({
      kind: "function" as const,
      name: entry.name,
      category: entry.category,
      summary: entry.summary,
      signature: formatFormulaSignature(entry),
    }));
}

function resolveTokenBounds(
  value: string,
  caret: number,
): { start: number; end: number; prefix: string } | null {
  if (!value.startsWith("=") || caret < 1) {
    return null;
  }
  let start = caret;
  while (start > 1 && isIdentifierCharacter(value[start - 1])) {
    start -= 1;
  }
  const prefix = value.slice(start, caret).toUpperCase();
  const previous = previousNonWhitespaceChar(value, start - 1);
  const allowedPrevious =
    previous === null ||
    previous === "=" ||
    previous === "(" ||
    previous === "," ||
    previous === "+" ||
    previous === "-" ||
    previous === "*" ||
    previous === "/" ||
    previous === "^" ||
    previous === "&" ||
    previous === "<" ||
    previous === ">";
  if (prefix.length === 0 && previous !== "=" && previous !== "(" && previous !== ",") {
    return null;
  }
  if (!allowedPrevious) {
    return null;
  }
  return { start, end: caret, prefix };
}

function resolveActiveFunction(value: string, caret: number) {
  if (!value.startsWith("=") || caret <= 1) {
    return null;
  }
  const stack: Array<{ name: string | null; argIndex: number }> = [];
  let insideString = false;
  let insideQuotedIdentifier = false;
  let bracketDepth = 0;
  const end = Math.min(caret, value.length);

  for (let index = 1; index < end; index += 1) {
    const char = value[index]!;
    if (insideString) {
      if (char === '"' && isEscapedDoubleQuote(value, index)) {
        index += 1;
        continue;
      }
      if (char === '"') {
        insideString = false;
      }
      continue;
    }
    if (insideQuotedIdentifier) {
      if (char === "'" && isEscapedSingleQuote(value, index)) {
        index += 1;
        continue;
      }
      if (char === "'") {
        insideQuotedIdentifier = false;
      }
      continue;
    }
    if (char === '"') {
      insideString = true;
      continue;
    }
    if (char === "'") {
      insideQuotedIdentifier = true;
      continue;
    }
    if (char === "[") {
      bracketDepth += 1;
      continue;
    }
    if (char === "]" && bracketDepth > 0) {
      bracketDepth -= 1;
      continue;
    }
    if (bracketDepth > 0) {
      continue;
    }
    if (char === "(") {
      stack.push({ name: findCalleeBeforeParen(value, index), argIndex: 0 });
      continue;
    }
    if (char === ")") {
      stack.pop();
      continue;
    }
    if (char === "," && stack.length > 0) {
      const top = stack.at(-1);
      if (top && top.name) {
        top.argIndex += 1;
      }
    }
  }

  for (let index = stack.length - 1; index >= 0; index -= 1) {
    const active = stack[index];
    if (!active?.name) {
      continue;
    }
    const entry = functionHelpByName.get(active.name);
    if (!entry) {
      return null;
    }
    return {
      entry,
      activeArgumentIndex: active.argIndex,
      signature: formatFormulaSignature(entry),
    };
  }
  return null;
}

export function resolveFormulaAssistState(input: {
  readonly value: string;
  readonly caret: number;
  readonly definedNames?: readonly WorkbookDefinedNameSnapshot[];
}): FormulaAssistState {
  const token = resolveTokenBounds(input.value, input.caret);
  const prefix = token?.prefix ?? "";
  const functionSuggestions = token ? buildFunctionSuggestions(prefix) : [];
  const definedNameSuggestions =
    token && input.definedNames ? buildDefinedNameSuggestions(prefix, input.definedNames) : [];
  return {
    tokenStart: token?.start ?? null,
    tokenEnd: token?.end ?? null,
    suggestions: [...definedNameSuggestions, ...functionSuggestions].slice(0, 12),
    activeFunction: resolveActiveFunction(input.value, input.caret),
  };
}

export function applyFormulaSuggestion(input: {
  readonly value: string;
  readonly tokenStart: number;
  readonly tokenEnd: number;
  readonly suggestion: FormulaSuggestion;
}): FormulaReplaceResult {
  const before = input.value.slice(0, input.tokenStart);
  const after = input.value.slice(input.tokenEnd);
  if (input.suggestion.kind === "defined-name") {
    const value = `${before}${input.suggestion.insertText}${after}`;
    return {
      value,
      caret: before.length + input.suggestion.insertText.length,
    };
  }
  if (after.startsWith("(")) {
    const value = `${before}${input.suggestion.name}${after}`;
    return {
      value,
      caret: before.length + input.suggestion.name.length + 1,
    };
  }
  const insertText = `${input.suggestion.name}()`;
  return {
    value: `${before}${insertText}${after}`,
    caret: before.length + input.suggestion.name.length + 1,
  };
}
