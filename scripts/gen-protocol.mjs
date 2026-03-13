import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");

const enumManifest = {
  ValueTag: [
    ["Empty", 0],
    ["Number", 1],
    ["Boolean", 2],
    ["String", 3],
    ["Error", 4]
  ],
  ErrorCode: [
    ["None", 0],
    ["Div0", 1],
    ["Ref", 2],
    ["Value", 3],
    ["Name", 4],
    ["NA", 5],
    ["Cycle", 6]
  ],
  FormulaMode: [
    ["JsOnly", 0],
    ["WasmFastPath", 1]
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
    ["Ret", 255]
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
    ["Concat", 17]
  ]
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
  { id: "Len", name: "LEN", supportsWasm: false },
  { id: "Concat", name: "CONCAT", supportsWasm: false }
];

const generatedHeader = `// GENERATED FILE. DO NOT EDIT DIRECTLY.\n// Source: scripts/gen-protocol.mjs\n\n`;

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
  return enumManifest.Opcode.map(
    ([name]) => `  [Opcode.${name}]: "${name}"`
  ).join(",\n");
}

function renderBuiltins() {
  return builtinManifest
    .map(
      ({ id, name, supportsWasm }) =>
        `  { id: BuiltinId.${id}, name: "${name}", supportsWasm: ${supportsWasm} }`
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
    contents: renderProtocolEnums()
  },
  {
    path: path.join(repoRoot, "packages/protocol/src/opcodes.ts"),
    contents: renderOpcodesModule()
  },
  {
    path: path.join(repoRoot, "packages/wasm-kernel/assembly/protocol.ts"),
    contents: renderProtocolEnums()
  }
];

async function main() {
  const checkMode = process.argv.includes("--check");
  const staleFiles = [];

  for (const file of generatedFiles) {
    let existing = "";
    try {
      existing = await readFile(file.path, "utf8");
    } catch {
      existing = "";
    }

    if (existing !== file.contents) {
      staleFiles.push(path.relative(repoRoot, file.path));
      if (!checkMode) {
        await writeFile(file.path, file.contents, "utf8");
      }
    }
  }

  if (checkMode && staleFiles.length > 0) {
    console.error(`Protocol artifacts are stale:\n${staleFiles.map((entry) => `- ${entry}`).join("\n")}`);
    process.exitCode = 1;
    return;
  }

  if (!checkMode) {
    if (staleFiles.length === 0) {
      console.log("Protocol artifacts are already up to date.");
      return;
    }
    console.log(`Updated protocol artifacts:\n${staleFiles.map((entry) => `- ${entry}`).join("\n")}`);
  }
}

await main();
