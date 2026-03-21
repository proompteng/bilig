import {
  ErrorCode,
  ValueTag,
  type CellValue,
  type PivotAggregation,
  type WorkbookPivotSnapshot,
  type WorkbookPivotValueSnapshot
} from "@bilig/protocol";

export type PivotDefinitionInput = Pick<WorkbookPivotSnapshot, "groupBy" | "values">;

export type PivotMaterializationResult =
  | {
      kind: "ok";
      rows: number;
      cols: number;
      values: CellValue[];
    }
  | {
      kind: "error";
      code: ErrorCode.Value;
      rows: 1;
      cols: 1;
      values: [CellValue];
    };

interface MaterializedPivotField extends WorkbookPivotValueSnapshot {
  columnIndex: number;
  headerLabel: string;
}

interface GroupBucket {
  keyValues: CellValue[];
  aggregates: number[];
}

export function materializePivotTable(
  definition: PivotDefinitionInput,
  sourceRows: readonly (readonly CellValue[])[]
): PivotMaterializationResult {
  if (definition.values.length === 0) {
    return pivotConfigError();
  }

  const headerRow = sourceRows[0];
  if (!headerRow || headerRow.length === 0) {
    return pivotConfigError();
  }

  const headerLookup = new Map<string, { index: number; label: string }>();
  for (let columnIndex = 0; columnIndex < headerRow.length; columnIndex += 1) {
    const label = headerLabel(headerRow[columnIndex] ?? emptyValue());
    const normalized = normalizeHeader(label);
    if (normalized.length === 0 || headerLookup.has(normalized)) {
      continue;
    }
    headerLookup.set(normalized, { index: columnIndex, label });
  }

  const groupFields = definition.groupBy.map((fieldName) => {
    const resolved = headerLookup.get(normalizeHeader(fieldName));
    return resolved ? { columnIndex: resolved.index, headerLabel: resolved.label } : undefined;
  });
  if (groupFields.some((field) => field === undefined)) {
    return pivotConfigError();
  }
  const materializedGroupFields = groupFields.filter((field): field is { columnIndex: number; headerLabel: string } => field !== undefined);

  const valueFields = definition.values.map((field) => resolveValueField(field, headerLookup));
  if (valueFields.some((field) => field === undefined)) {
    return pivotConfigError();
  }
  const materializedValueFields = valueFields.filter((field): field is MaterializedPivotField => field !== undefined);

  const buckets = new Map<string, GroupBucket>();
  for (let rowIndex = 1; rowIndex < sourceRows.length; rowIndex += 1) {
    const row = sourceRows[rowIndex] ?? [];
    const keyValues = materializedGroupFields.map((field) => cloneOutputCell(row[field.columnIndex] ?? emptyValue()));
    const hasObservedValue = hasMeaningfulRowValue(keyValues, materializedValueFields, row);
    if (!hasObservedValue) {
      continue;
    }
    const key = keyValues.map(cellValueKey).join("\u001f");
    let bucket = buckets.get(key);
    if (!bucket) {
      bucket = {
        keyValues,
        aggregates: new Array(materializedValueFields.length).fill(0)
      };
      buckets.set(key, bucket);
    }
    for (let valueIndex = 0; valueIndex < materializedValueFields.length; valueIndex += 1) {
      const field = materializedValueFields[valueIndex]!;
      const cell = row[field.columnIndex] ?? emptyValue();
      bucket.aggregates[valueIndex] = (bucket.aggregates[valueIndex] ?? 0) + accumulateValue(field.summarizeBy, cell);
    }
  }

  const cols = definition.groupBy.length + materializedValueFields.length;
  const values: CellValue[] = [];
  for (let groupIndex = 0; groupIndex < materializedGroupFields.length; groupIndex += 1) {
    values.push(stringValue(materializedGroupFields[groupIndex]!.headerLabel));
  }
  for (let valueIndex = 0; valueIndex < materializedValueFields.length; valueIndex += 1) {
    values.push(stringValue(outputLabel(materializedValueFields[valueIndex]!)));
  }
  buckets.forEach((bucket) => {
    bucket.keyValues.forEach((keyValue) => {
      values.push(cloneOutputCell(keyValue));
    });
    bucket.aggregates.forEach((aggregate) => {
      values.push(numberValue(aggregate));
    });
  });

  return {
    kind: "ok",
    rows: buckets.size + 1,
    cols,
    values
  };
}

function resolveValueField(
  field: WorkbookPivotValueSnapshot,
  headerLookup: Map<string, { index: number; label: string }>
): MaterializedPivotField | undefined {
  const resolved = headerLookup.get(normalizeHeader(field.sourceColumn));
  if (!resolved) {
    return undefined;
  }
  return {
    ...field,
    columnIndex: resolved.index,
    headerLabel: resolved.label
  };
}

function accumulateValue(mode: PivotAggregation, value: CellValue): number {
  if (mode === "count") {
    return isEmptyValue(value) ? 0 : 1;
  }
  return value.tag === ValueTag.Number ? value.value : 0;
}

function outputLabel(field: MaterializedPivotField): string {
  const customLabel = field.outputLabel?.trim();
  if (customLabel && customLabel.length > 0) {
    return customLabel;
  }
  return `${field.summarizeBy.toUpperCase()} of ${field.headerLabel}`;
}

function hasMeaningfulRowValue(
  keyValues: readonly CellValue[],
  valueFields: readonly MaterializedPivotField[],
  row: readonly CellValue[]
): boolean {
  if (keyValues.some((value) => !isEmptyValue(value))) {
    return true;
  }
  return valueFields.some((field) => !isEmptyValue(row[field.columnIndex] ?? emptyValue()));
}

function cellValueKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "E";
    case ValueTag.Number:
      return `N:${Object.is(value.value, -0) ? "-0" : String(value.value)}`;
    case ValueTag.Boolean:
      return value.value ? "B:1" : "B:0";
    case ValueTag.String:
      return `S:${value.value}`;
    case ValueTag.Error:
      return `R:${value.code}`;
  }
}

function headerLabel(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return value.value.trim();
    case ValueTag.Error:
      return "";
  }
}

function normalizeHeader(value: string): string {
  return value.trim().toUpperCase();
}

function isEmptyValue(value: CellValue): boolean {
  return value.tag === ValueTag.Empty;
}

function cloneOutputCell(value: CellValue): CellValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return emptyValue();
    case ValueTag.Number:
      return numberValue(value.value);
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: value.value };
    case ValueTag.String:
      return stringValue(value.value);
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: value.code };
  }
}

function pivotConfigError(): PivotMaterializationResult {
  return {
    kind: "error",
    code: ErrorCode.Value,
    rows: 1,
    cols: 1,
    values: [{ tag: ValueTag.Error, code: ErrorCode.Value }]
  };
}

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}
