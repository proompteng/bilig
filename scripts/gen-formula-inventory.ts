#!/usr/bin/env bun

import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { BUILTINS } from "../packages/protocol/src/opcodes.ts";
import { getBuiltinCapability } from "../packages/formula/src/builtin-capabilities.ts";
import { scalarPlaceholderBuiltinNames } from "../packages/formula/src/builtins/placeholder.ts";
import { hasBuiltin } from "../packages/formula/src/builtins.ts";

interface FormulaInventorySourceEntry {
  name: string;
  odfStatus: string;
  inOfficeList: boolean;
}

interface FormulaInventorySource {
  version: number;
  entries: FormulaInventorySourceEntry[];
}

type FormulaDeterminism = "deterministic" | "provider-backed";
type FormulaRuntimeStatus = "missing" | "placeholder" | "implemented";

const repoRoot = path.resolve(import.meta.dir, "..");
const sourcePath = path.join(repoRoot, "packages/formula/src/formula-inventory-source.json");
const generatedReportPath = path.join(
  repoRoot,
  "packages/formula/src/generated/formula-inventory.ts",
);
const generatedDocPath = path.join(repoRoot, "docs/odf-1.4-mandatory-office-excel-functions.md");
const missingProtocolPath = path.join(repoRoot, "docs/formula-inventory-missing-from-protocol.md");

const providerBackedNames = new Set([
  "CALL",
  "COPILOT",
  "CUBEKPIMEMBER",
  "CUBEMEMBER",
  "CUBEMEMBERPROPERTY",
  "CUBERANKEDMEMBER",
  "CUBESET",
  "CUBESETCOUNT",
  "CUBEVALUE",
  "DDE",
  "DETECTLANGUAGE",
  "FILTERXML",
  "HYPERLINK",
  "IMAGE",
  "RTD",
  "STOCKHISTORY",
  "WEBSERVICE",
]);

const placeholderNames = new Set(scalarPlaceholderBuiltinNames.map((name) => name.toUpperCase()));
const protocolBuiltinsByName = new Map(
  BUILTINS.map((builtin) => [builtin.name.toUpperCase(), builtin]),
);
const runtimeAliasByCanonicalName = new Map<string, string>([["AVERAGE", "AVG"]]);

function escapeTsString(value: string): string {
  return JSON.stringify(value);
}

function escapeMd(value: string): string {
  return value.replaceAll("|", "\\|");
}

function inferDeterminism(name: string): FormulaDeterminism {
  return providerBackedNames.has(name) ? "provider-backed" : "deterministic";
}

function runtimeLookupNames(name: string): string[] {
  const canonical = name.trim().toUpperCase();
  const alias = runtimeAliasByCanonicalName.get(canonical);
  return alias ? [canonical, alias] : [canonical];
}

function inferRuntimeStatus(name: string): FormulaRuntimeStatus {
  const lookupNames = runtimeLookupNames(name);
  if (lookupNames.some((entry) => placeholderNames.has(entry))) {
    return "placeholder";
  }
  return lookupNames.some((entry) => hasBuiltin(entry)) ? "implemented" : "missing";
}

function summarizeOdfMembership(odfStatus: string): boolean {
  return odfStatus === "Implemented";
}

function renderInventoryReport(source: FormulaInventorySource): string {
  const entries = source.entries.map((entry) => {
    const name = entry.name.trim().toUpperCase();
    const capability = runtimeLookupNames(name)
      .map((entryName) => getBuiltinCapability(entryName))
      .find((entryCapability) => entryCapability !== undefined);
    const protocolBuiltin = runtimeLookupNames(name)
      .map((entryName) => protocolBuiltinsByName.get(entryName))
      .find((entryBuiltin) => entryBuiltin !== undefined);
    const runtimeStatus = inferRuntimeStatus(name);
    return {
      name,
      odfStatus: entry.odfStatus,
      inOfficeList: entry.inOfficeList,
      deterministic: inferDeterminism(name),
      runtimeStatus,
      protocolId: protocolBuiltin?.id,
      protocolName: protocolBuiltin?.name,
      protocolSupportsWasm: protocolBuiltin?.supportsWasm ?? false,
      jsStatus:
        capability?.jsStatus ??
        (runtimeStatus === "implemented"
          ? "implemented"
          : runtimeStatus === "placeholder"
            ? "placeholder"
            : "missing"),
      wasmStatus: capability?.wasmStatus ?? "not-started",
      placeholder: runtimeStatus === "placeholder",
      registeredInCodebase: runtimeStatus !== "missing",
    };
  });

  const summary = {
    total: entries.length,
    odfMandatory: entries.filter((entry) => summarizeOdfMembership(entry.odfStatus)).length,
    officeListed: entries.filter((entry) => entry.inOfficeList).length,
    overlap: entries.filter(
      (entry) => summarizeOdfMembership(entry.odfStatus) && entry.inOfficeList,
    ).length,
    odfOnly: entries.filter(
      (entry) => summarizeOdfMembership(entry.odfStatus) && !entry.inOfficeList,
    ).length,
    officeOnly: entries.filter(
      (entry) => !summarizeOdfMembership(entry.odfStatus) && entry.inOfficeList,
    ).length,
    registeredInCodebase: entries.filter((entry) => entry.registeredInCodebase).length,
    missingInCodebase: entries.filter((entry) => !entry.registeredInCodebase).length,
    placeholders: entries.filter((entry) => entry.placeholder).length,
    protocolBuiltins: entries.filter((entry) => entry.protocolId !== undefined).length,
    runtimeRegisteredMissingProtocol: entries.filter(
      (entry) =>
        entry.deterministic === "deterministic" &&
        entry.registeredInCodebase &&
        entry.protocolId === undefined,
    ).length,
  };

  const lines = entries.map((entry) => {
    return `  {
    name: ${escapeTsString(entry.name)},
    odfStatus: ${escapeTsString(entry.odfStatus)},
    inOfficeList: ${entry.inOfficeList},
    deterministic: ${escapeTsString(entry.deterministic)},
    runtimeStatus: ${escapeTsString(entry.runtimeStatus)},
    protocolId: ${entry.protocolId ?? "undefined"},
    protocolName: ${entry.protocolName ? escapeTsString(entry.protocolName) : "undefined"},
    protocolSupportsWasm: ${entry.protocolSupportsWasm},
    jsStatus: ${escapeTsString(entry.jsStatus)},
    wasmStatus: ${escapeTsString(entry.wasmStatus)},
    placeholder: ${entry.placeholder},
    registeredInCodebase: ${entry.registeredInCodebase},
  }`;
  });

  return `// GENERATED FILE. DO NOT EDIT DIRECTLY.
// Source: scripts/gen-formula-inventory.ts

export const formulaInventorySummary = ${JSON.stringify(summary, null, 2)} as const;

export const formulaInventory = [
${lines.join(",\n")}
] as const;
`;
}

function renderInventoryDoc(source: FormulaInventorySource): string {
  const rows = source.entries.map((entry) => {
    const name = entry.name.trim().toUpperCase();
    const runtimeStatus = inferRuntimeStatus(name);
    const implemented = runtimeStatus === "missing" ? "No" : "Yes";
    return `| ${escapeMd(name)} | ${escapeMd(entry.odfStatus)} | ${entry.inOfficeList ? "Yes" : "No"} | ${implemented} |`;
  });

  const odfMandatory = source.entries.filter((entry) =>
    summarizeOdfMembership(entry.odfStatus),
  ).length;
  const officeCount = source.entries.filter((entry) => entry.inOfficeList).length;
  const overlap = source.entries.filter(
    (entry) => summarizeOdfMembership(entry.odfStatus) && entry.inOfficeList,
  ).length;
  const odfOnly = source.entries.filter(
    (entry) => summarizeOdfMembership(entry.odfStatus) && !entry.inOfficeList,
  ).length;
  const officeOnly = source.entries.filter(
    (entry) => !summarizeOdfMembership(entry.odfStatus) && entry.inOfficeList,
  ).length;
  const implementedCount = source.entries.filter(
    (entry) => inferRuntimeStatus(entry.name.trim().toUpperCase()) !== "missing",
  ).length;

  return `# ODF 1.4 and Office Excel Function Coverage

## Source
- Canonical inventory: \`packages/formula/src/formula-inventory-source.json\`
- Generated by: \`bun scripts/gen-formula-inventory.ts\`
- ODF 1.4 Spreadsheet OpenFormula Formula Functions
- OASIS OpenDocument v1.4 function requirements:
  - https://docs.oasis-open.org/office/OpenDocument/v1.4/os/v1.4-os.html
- Microsoft Office Excel functions by category:
  - https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb
- Scope: ODF mandatory functions + Office by-category function list (merged into one list).

## Coverage Summary
- Unified function count: **${source.entries.length}**
- ODF mandatory function count: **${odfMandatory}**
- Office function count (cleaned scrape): **${officeCount}**
- Overlap (present in both): **${overlap}**
- ODF-only (mandatory, not listed by Office): **${odfOnly}**
- Office-only (not in ODF 1.4 mandatory): **${officeOnly}**

## Current code coverage snapshot
- Registered in codebase: **${implementedCount}**
- Not yet registered in codebase: **${source.entries.length - implementedCount}**
- Placeholder-backed registrations: **${source.entries.filter((entry) => inferRuntimeStatus(entry.name.trim().toUpperCase()) === "placeholder").length}**
- The "Implemented in codebase" column reflects runtime registration, including blocked placeholder registrations.

## Full unified function list (ODF 1.4 mandatory + Office category)

| Function | ODF status | In Office list | Implemented in codebase |
| --- | --- | --- | --- |
${rows.join("\n")}
`;
}

function renderMissingProtocolDoc(source: FormulaInventorySource): string {
  const entries = source.entries
    .map((entry) => {
      const name = entry.name.trim().toUpperCase();
      const runtimeStatus = inferRuntimeStatus(name);
      return {
        name,
        runtimeStatus,
        deterministic: inferDeterminism(name),
      };
    })
    .filter(
      (entry) => entry.deterministic === "deterministic" && !protocolBuiltinsByName.has(entry.name),
    )
    .toSorted((left, right) => left.name.localeCompare(right.name));

  const registered = entries.filter((entry) => entry.runtimeStatus !== "missing");
  const missing = entries.filter((entry) => entry.runtimeStatus === "missing");

  return `# Formula Inventory Missing From Protocol

Generated by \`bun scripts/gen-formula-inventory.ts\`.

## Deterministic formulas registered in codebase but missing \`BuiltinId\`

${registered.length === 0 ? "- None" : registered.map((entry) => `- \`${entry.name}\` (${entry.runtimeStatus})`).join("\n")}

## Deterministic formulas not yet in protocol and not yet registered

${missing.length === 0 ? "- None" : missing.map((entry) => `- \`${entry.name}\``).join("\n")}
`;
}

function isFormulaInventorySource(value: unknown): value is FormulaInventorySource {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const candidate = value as { version?: unknown; entries?: unknown };
  if (candidate.version !== 1 || !Array.isArray(candidate.entries)) {
    return false;
  }
  return candidate.entries.every((entry) => {
    if (typeof entry !== "object" || entry === null) {
      return false;
    }
    const candidateEntryName = Reflect.get(entry, "name");
    const candidateEntryOdfStatus = Reflect.get(entry, "odfStatus");
    const candidateEntryInOfficeList = Reflect.get(entry, "inOfficeList");
    return (
      typeof candidateEntryName === "string" &&
      typeof candidateEntryOdfStatus === "string" &&
      typeof candidateEntryInOfficeList === "boolean"
    );
  });
}

async function readSource(): Promise<FormulaInventorySource> {
  const sourceText = await readFile(sourcePath, "utf8");
  const parsed = JSON.parse(sourceText) as unknown;
  if (!isFormulaInventorySource(parsed)) {
    throw new Error(`Invalid formula inventory source at ${path.relative(repoRoot, sourcePath)}`);
  }
  const source = parsed;
  const seen = new Set<string>();
  for (const entry of source.entries) {
    const normalized = entry.name.trim().toUpperCase();
    if (seen.has(normalized)) {
      throw new Error(`Duplicate formula inventory entry: ${normalized}`);
    }
    seen.add(normalized);
  }
  for (const builtin of BUILTINS) {
    const builtinName = builtin.name.trim().toUpperCase();
    if (builtinName.startsWith("__")) {
      continue;
    }
    const hasDirect = seen.has(builtinName);
    const hasAlias = [...runtimeAliasByCanonicalName.entries()].some(
      ([canonical, alias]) => alias === builtinName && seen.has(canonical),
    );
    if (!hasDirect && !hasAlias) {
      throw new Error(`Protocol builtin missing from inventory: ${builtin.name}`);
    }
  }
  return source;
}

async function writeGeneratedFiles(checkMode: boolean): Promise<void> {
  const source = await readSource();
  const outputs = [
    { path: generatedReportPath, contents: renderInventoryReport(source) },
    { path: generatedDocPath, contents: renderInventoryDoc(source) },
    { path: missingProtocolPath, contents: renderMissingProtocolDoc(source) },
  ];

  const stalePaths: string[] = [];
  await mkdir(path.dirname(generatedReportPath), { recursive: true });

  const existingOutputs = await Promise.all(
    outputs.map(async (output) => {
      try {
        return await readFile(output.path, "utf8");
      } catch {
        return "";
      }
    }),
  );

  const staleOutputs = outputs.filter(
    (output, index) => existingOutputs[index] !== output.contents,
  );
  if (!checkMode) {
    await Promise.all(
      staleOutputs.map((output) => writeFile(output.path, output.contents, "utf8")),
    );
  } else {
    staleOutputs.forEach((output) => stalePaths.push(path.relative(repoRoot, output.path)));
  }

  if (stalePaths.length > 0) {
    throw new Error(
      `Formula inventory outputs are stale. Regenerate with \`bun scripts/gen-formula-inventory.ts\`:\n${stalePaths.map((entry) => `- ${entry}`).join("\n")}`,
    );
  }
}

await writeGeneratedFiles(Bun.argv.includes("--check"));
