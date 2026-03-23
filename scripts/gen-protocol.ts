#!/usr/bin/env bun

import path from "node:path";

const repoRoot = path.resolve(import.meta.dir, "..");

const enumManifest = {
  ValueTag: [
    ["Empty", 0],
    ["Number", 1],
    ["Boolean", 2],
    ["String", 3],
    ["Error", 4],
  ],
  ErrorCode: [
    ["None", 0],
    ["Div0", 1],
    ["Ref", 2],
    ["Value", 3],
    ["Name", 4],
    ["NA", 5],
    ["Cycle", 6],
    ["Spill", 7],
    ["Blocked", 8],
  ],
  FormulaMode: [
    ["JsOnly", 0],
    ["WasmFastPath", 1],
  ],
  Opcode: [
    ["PushNumber", 1],
    ["PushBoolean", 2],
    ["PushCell", 3],
    ["PushRange", 4],
    ["Add", 5],
    ["Sub", 6],
    ["Mul", 7],
    ["Div", 8],
    ["Pow", 9],
    ["Concat", 10],
    ["Neg", 11],
    ["Eq", 12],
    ["Neq", 13],
    ["Gt", 14],
    ["Gte", 15],
    ["Lt", 16],
    ["Lte", 17],
    ["Jump", 18],
    ["JumpIfFalse", 19],
    ["CallBuiltin", 20],
    ["PushString", 21],
    ["PushError", 22],
    ["Ret", 255],
  ],
  BuiltinId: [
    ["Sum", 1],
    ["Avg", 2],
    ["Min", 3],
    ["Max", 4],
    ["Count", 5],
    ["CountA", 6],
    ["Abs", 7],
    ["Round", 8],
    ["Floor", 9],
    ["Ceiling", 10],
    ["Mod", 11],
    ["If", 12],
    ["And", 13],
    ["Or", 14],
    ["Not", 15],
    ["Len", 16],
    ["Concat", 17],
    ["IsBlank", 18],
    ["IsNumber", 19],
    ["IsText", 20],
    ["Date", 21],
    ["Year", 22],
    ["Month", 23],
    ["Day", 24],
    ["Edate", 25],
    ["Eomonth", 26],
    ["Exact", 27],
    ["Int", 28],
    ["RoundUp", 29],
    ["RoundDown", 30],
    ["Time", 31],
    ["Hour", 32],
    ["Minute", 33],
    ["Second", 34],
    ["Weekday", 35],
    ["Sin", 36],
    ["Cos", 37],
    ["Tan", 38],
    ["Asin", 39],
    ["Acos", 40],
    ["Atan", 41],
    ["Atan2", 42],
    ["Degrees", 43],
    ["Radians", 44],
    ["Exp", 45],
    ["Ln", 46],
    ["Log", 47],
    ["Log10", 48],
    ["Power", 49],
    ["Sqrt", 50],
    ["Pi", 51],
    ["Ifs", 52],
    ["Switch", 53],
    ["Xor", 54],
    ["Ifna", 55],
    ["Days", 56],
    ["Workday", 57],
    ["Networkdays", 58],
    ["Weeknum", 59],
    ["Replace", 60],
    ["Substitute", 61],
    ["Rept", 62],
    ["Na", 63],
    ["Iferror", 64],
    ["Match", 65],
    ["Index", 66],
    ["Vlookup", 67],
    ["Hlookup", 68],
    ["Xlookup", 69],
    ["Xmatch", 70],
    ["Countif", 71],
    ["Countifs", 72],
    ["Sumif", 73],
    ["Sumifs", 74],
    ["Averageif", 75],
    ["Averageifs", 76],
    ["Sumproduct", 77],
    ["Left", 78],
    ["Right", 79],
    ["Mid", 80],
    ["Trim", 81],
    ["Upper", 82],
    ["Lower", 83],
    ["Find", 84],
    ["Search", 85],
    ["Value", 86],
    ["Today", 87],
    ["Now", 88],
    ["Rand", 89],
    ["Sequence", 90],
    ["Offset", 91],
    ["Take", 92],
    ["Drop", 93],
    ["Choosecols", 94],
    ["Chooserows", 95],
    ["Sort", 96],
    ["Sortby", 97],
    ["Tocol", 98],
    ["Torow", 99],
    ["Wraprows", 100],
    ["Wrapcols", 101],
  ],
};

const builtinManifest = [
  { id: "Sum", name: "SUM", supportsWasm: true },
  { id: "Avg", name: "AVG", supportsWasm: true },
  { id: "Min", name: "MIN", supportsWasm: true },
  { id: "Max", name: "MAX", supportsWasm: true },
  { id: "Count", name: "COUNT", supportsWasm: true },
  { id: "CountA", name: "COUNTA", supportsWasm: true },
  { id: "Abs", name: "ABS", supportsWasm: true },
  { id: "Round", name: "ROUND", supportsWasm: true },
  { id: "Floor", name: "FLOOR", supportsWasm: true },
  { id: "Ceiling", name: "CEILING", supportsWasm: true },
  { id: "Mod", name: "MOD", supportsWasm: true },
  { id: "If", name: "IF", supportsWasm: true },
  { id: "And", name: "AND", supportsWasm: true },
  { id: "Or", name: "OR", supportsWasm: true },
  { id: "Not", name: "NOT", supportsWasm: true },
  { id: "Len", name: "LEN", supportsWasm: true },
  { id: "Concat", name: "CONCAT", supportsWasm: true },
  { id: "IsBlank", name: "ISBLANK", supportsWasm: true },
  { id: "IsNumber", name: "ISNUMBER", supportsWasm: true },
  { id: "IsText", name: "ISTEXT", supportsWasm: true },
  { id: "Date", name: "DATE", supportsWasm: true },
  { id: "Year", name: "YEAR", supportsWasm: true },
  { id: "Month", name: "MONTH", supportsWasm: true },
  { id: "Day", name: "DAY", supportsWasm: true },
  { id: "Edate", name: "EDATE", supportsWasm: true },
  { id: "Eomonth", name: "EOMONTH", supportsWasm: true },
  { id: "Exact", name: "EXACT", supportsWasm: true },
  { id: "Int", name: "INT", supportsWasm: true },
  { id: "RoundUp", name: "ROUNDUP", supportsWasm: true },
  { id: "RoundDown", name: "ROUNDDOWN", supportsWasm: true },
  { id: "Time", name: "TIME", supportsWasm: true },
  { id: "Hour", name: "HOUR", supportsWasm: true },
  { id: "Minute", name: "MINUTE", supportsWasm: true },
  { id: "Second", name: "SECOND", supportsWasm: true },
  { id: "Weekday", name: "WEEKDAY", supportsWasm: true },
  { id: "Sin", name: "SIN", supportsWasm: true },
  { id: "Cos", name: "COS", supportsWasm: true },
  { id: "Tan", name: "TAN", supportsWasm: true },
  { id: "Asin", name: "ASIN", supportsWasm: true },
  { id: "Acos", name: "ACOS", supportsWasm: true },
  { id: "Atan", name: "ATAN", supportsWasm: true },
  { id: "Atan2", name: "ATAN2", supportsWasm: true },
  { id: "Degrees", name: "DEGREES", supportsWasm: true },
  { id: "Radians", name: "RADIANS", supportsWasm: true },
  { id: "Exp", name: "EXP", supportsWasm: true },
  { id: "Ln", name: "LN", supportsWasm: true },
  { id: "Log", name: "LOG", supportsWasm: true },
  { id: "Log10", name: "LOG10", supportsWasm: true },
  { id: "Power", name: "POWER", supportsWasm: true },
  { id: "Sqrt", name: "SQRT", supportsWasm: true },
  { id: "Pi", name: "PI", supportsWasm: true },
  { id: "Ifs", name: "IFS", supportsWasm: true },
  { id: "Switch", name: "SWITCH", supportsWasm: true },
  { id: "Xor", name: "XOR", supportsWasm: true },
  { id: "Ifna", name: "IFNA", supportsWasm: true },
  { id: "Days", name: "DAYS", supportsWasm: true },
  { id: "Workday", name: "WORKDAY", supportsWasm: true },
  { id: "Networkdays", name: "NETWORKDAYS", supportsWasm: true },
  { id: "Weeknum", name: "WEEKNUM", supportsWasm: true },
  { id: "Replace", name: "REPLACE", supportsWasm: true },
  { id: "Substitute", name: "SUBSTITUTE", supportsWasm: true },
  { id: "Rept", name: "REPT", supportsWasm: true },
  { id: "Na", name: "NA", supportsWasm: true },
  { id: "Iferror", name: "IFERROR", supportsWasm: true },
  { id: "Match", name: "MATCH", supportsWasm: true },
  { id: "Index", name: "INDEX", supportsWasm: true },
  { id: "Vlookup", name: "VLOOKUP", supportsWasm: true },
  { id: "Hlookup", name: "HLOOKUP", supportsWasm: true },
  { id: "Xlookup", name: "XLOOKUP", supportsWasm: true },
  { id: "Xmatch", name: "XMATCH", supportsWasm: true },
  { id: "Countif", name: "COUNTIF", supportsWasm: true },
  { id: "Countifs", name: "COUNTIFS", supportsWasm: true },
  { id: "Sumif", name: "SUMIF", supportsWasm: true },
  { id: "Sumifs", name: "SUMIFS", supportsWasm: true },
  { id: "Averageif", name: "AVERAGEIF", supportsWasm: true },
  { id: "Averageifs", name: "AVERAGEIFS", supportsWasm: true },
  { id: "Sumproduct", name: "SUMPRODUCT", supportsWasm: true },
  { id: "Left", name: "LEFT", supportsWasm: true },
  { id: "Right", name: "RIGHT", supportsWasm: true },
  { id: "Mid", name: "MID", supportsWasm: true },
  { id: "Trim", name: "TRIM", supportsWasm: true },
  { id: "Upper", name: "UPPER", supportsWasm: true },
  { id: "Lower", name: "LOWER", supportsWasm: true },
  { id: "Find", name: "FIND", supportsWasm: true },
  { id: "Search", name: "SEARCH", supportsWasm: true },
  { id: "Value", name: "VALUE", supportsWasm: true },
  { id: "Today", name: "TODAY", supportsWasm: true },
  { id: "Now", name: "NOW", supportsWasm: true },
  { id: "Rand", name: "RAND", supportsWasm: true },
  { id: "Sequence", name: "SEQUENCE", supportsWasm: true },
  { id: "Offset", name: "OFFSET", supportsWasm: true },
  { id: "Take", name: "TAKE", supportsWasm: true },
  { id: "Drop", name: "DROP", supportsWasm: true },
  { id: "Choosecols", name: "CHOOSECOLS", supportsWasm: true },
  { id: "Chooserows", name: "CHOOSEROWS", supportsWasm: true },
  { id: "Sort", name: "SORT", supportsWasm: true },
  { id: "Sortby", name: "SORTBY", supportsWasm: true },
  { id: "Tocol", name: "TOCOL", supportsWasm: true },
  { id: "Torow", name: "TOROW", supportsWasm: true },
  { id: "Wraprows", name: "WRAPROWS", supportsWasm: true },
  { id: "Wrapcols", name: "WRAPCOLS", supportsWasm: true },
];

const generatedHeader = `// GENERATED FILE. DO NOT EDIT DIRECTLY.\n// Source: scripts/gen-protocol.ts\n\n`;

function renderEnum(name, entries) {
  const lines = entries.map(([key, value]) => `  ${key} = ${value}`);
  return `export enum ${name} {\n${lines.join(",\n")}\n}\n`;
}

function renderProtocolEnums() {
  return (
    generatedHeader +
    Object.entries(enumManifest)
      .map(([name, entries]) => renderEnum(name, entries))
      .join("\n")
  );
}

function renderOpcodeNames() {
  return enumManifest.Opcode.map(([name]) => `  [Opcode.${name}]: "${name}"`).join(",\n");
}

function renderBuiltins() {
  return builtinManifest
    .map(
      ({ id, name, supportsWasm }) =>
        `  { id: BuiltinId.${id}, name: "${name}", supportsWasm: ${supportsWasm} }`,
    )
    .join(",\n");
}

function renderOpcodesModule() {
  return `${generatedHeader}import { BuiltinId, Opcode } from "./enums.js";

export interface BuiltinDescriptor {
  readonly id: BuiltinId;
  readonly name: string;
  readonly supportsWasm: boolean;
}

export const OPCODE_NAMES: Record<Opcode, string> = {
${renderOpcodeNames()}
};

export const BUILTINS: BuiltinDescriptor[] = [
${renderBuiltins()}
];
`;
}

const generatedFiles = [
  {
    path: path.join(repoRoot, "packages/protocol/src/enums.ts"),
    contents: renderProtocolEnums(),
  },
  {
    path: path.join(repoRoot, "packages/protocol/src/opcodes.ts"),
    contents: renderOpcodesModule(),
  },
  {
    path: path.join(repoRoot, "packages/wasm-kernel/assembly/protocol.ts"),
    contents: renderProtocolEnums(),
  },
];

async function main() {
  const checkMode = Bun.argv.includes("--check");
  const staleFiles = (
    await Promise.all(
      generatedFiles.map(async (file) => {
        let existing = "";
        try {
          existing = await Bun.file(file.path).text();
        } catch {
          existing = "";
        }

        if (existing === file.contents) {
          return null;
        }
        if (!checkMode) {
          await Bun.write(file.path, file.contents);
        }
        return path.relative(repoRoot, file.path);
      }),
    )
  ).filter((entry) => entry !== null);

  if (checkMode && staleFiles.length > 0) {
    console.error(
      `Protocol artifacts are stale:\n${staleFiles.map((entry) => `- ${entry}`).join("\n")}`,
    );
    process.exitCode = 1;
    return;
  }

  if (!checkMode) {
    if (staleFiles.length === 0) {
      console.log("Protocol artifacts are already up to date.");
      return;
    }
    console.log(
      `Updated protocol artifacts:\n${staleFiles.map((entry) => `- ${entry}`).join("\n")}`,
    );
  }
}

await main();
