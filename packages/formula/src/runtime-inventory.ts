import {
  builtinJsSpecialNames,
  builtinWasmEnabledNames,
  type BuiltinJsStatus,
  type BuiltinWasmStatus,
} from "./builtin-capabilities.js";
import { getBuiltin, hasBuiltin } from "./builtins.js";
import { lookupBuiltins } from "./builtins/lookup.js";
import { placeholderBuiltinNames } from "./builtins/placeholder.js";

export type FormulaRuntimeStatus = "missing" | "placeholder" | "implemented";
export type FormulaRuntimeJsStatus = BuiltinJsStatus | FormulaRuntimeStatus;

const lookupBuiltinNameSet = new Set(Object.keys(lookupBuiltins).map(normalizeFormulaName));
const placeholderBuiltinNameSet = new Set(placeholderBuiltinNames.map(normalizeFormulaName));
const runtimeAliasByCanonicalName = new Map<string, string>([
  ["AVERAGE", "AVG"],
  ["USE.THE.COUNTIF", "COUNTIF"],
]);

export function normalizeFormulaName(name: string): string {
  return name.trim().toUpperCase();
}

export function getFormulaRuntimeLookupNames(name: string): string[] {
  const canonical = normalizeFormulaName(name);
  const alias = runtimeAliasByCanonicalName.get(canonical);
  return alias ? [canonical, alias] : [canonical];
}

export function isLookupBuiltinRuntime(name: string): boolean {
  return getFormulaRuntimeLookupNames(name).some((entry) => lookupBuiltinNameSet.has(entry));
}

export function isScalarBuiltinRuntime(name: string): boolean {
  return getFormulaRuntimeLookupNames(name).some((entry) => getBuiltin(entry) !== undefined);
}

export function isPlaceholderBuiltinRuntime(name: string): boolean {
  return getFormulaRuntimeLookupNames(name).some((entry) => placeholderBuiltinNameSet.has(entry));
}

export function getFormulaRuntimeStatus(name: string): FormulaRuntimeStatus {
  if (isPlaceholderBuiltinRuntime(name)) {
    return "placeholder";
  }
  return getFormulaRuntimeLookupNames(name).some((entry) => hasBuiltin(entry))
    ? "implemented"
    : "missing";
}

export function getFormulaRuntimeJsStatus(name: string): FormulaRuntimeJsStatus {
  const runtimeStatus = getFormulaRuntimeStatus(name);
  if (runtimeStatus !== "implemented") {
    return runtimeStatus;
  }
  return getFormulaRuntimeLookupNames(name).some((entry) => builtinJsSpecialNames.has(entry))
    ? "special-js-only"
    : "implemented";
}

export function getFormulaRuntimeWasmStatus(name: string): BuiltinWasmStatus {
  return getFormulaRuntimeLookupNames(name).some((entry) => builtinWasmEnabledNames.has(entry))
    ? "production"
    : "not-started";
}
