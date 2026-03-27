import type {
  CellRangeRef,
  CompatibilityMode,
  PivotAggregation,
  LiteralInput,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationMode,
  WorkbookDefinedNameValueSnapshot,
  WorkbookPivotValueSnapshot,
} from "@bilig/protocol";
import type {
  EngineOp,
  EngineOpBatch,
  WorkbookSortDirection,
  WorkbookSortKey,
  WorkbookTableOp,
} from "@bilig/crdt";

export const PROTOCOL_MAGIC = 0x424c4731;
export const PROTOCOL_VERSION = 1;
export const WORKBOOK_SNAPSHOT_CONTENT_TYPE = "application/vnd.bilig.workbook+json";

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

export type FrameKind =
  | "hello"
  | "appendBatch"
  | "ack"
  | "snapshotChunk"
  | "cursorWatermark"
  | "heartbeat"
  | "error";

export interface HelloFrame {
  kind: "hello";
  documentId: string;
  replicaId: string;
  sessionId: string;
  protocolVersion: number;
  lastServerCursor: number;
  capabilities: string[];
}

export interface BatchAppendFrame {
  kind: "appendBatch";
  documentId: string;
  cursor: number;
  batch: EngineOpBatch;
}

export interface AckFrame {
  kind: "ack";
  documentId: string;
  batchId: string;
  cursor: number;
  acceptedAtUnixMs: number;
}

export interface SnapshotChunkFrame {
  kind: "snapshotChunk";
  documentId: string;
  snapshotId: string;
  cursor: number;
  chunkIndex: number;
  chunkCount: number;
  contentType: string;
  bytes: Uint8Array;
}

export interface SnapshotChunkOptions {
  documentId: string;
  snapshotId: string;
  cursor: number;
  contentType: string;
  bytes: Uint8Array;
  chunkSize?: number;
}

export interface CursorWatermarkFrame {
  kind: "cursorWatermark";
  documentId: string;
  cursor: number;
  compactedCursor: number;
}

export interface HeartbeatFrame {
  kind: "heartbeat";
  documentId: string;
  cursor: number;
  sentAtUnixMs: number;
}

export interface ErrorFrame {
  kind: "error";
  documentId: string;
  code: string;
  message: string;
  retryable: boolean;
}

export type ProtocolFrame =
  | HelloFrame
  | BatchAppendFrame
  | AckFrame
  | SnapshotChunkFrame
  | CursorWatermarkFrame
  | HeartbeatFrame
  | ErrorFrame;

export class BinaryProtocolError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "BinaryProtocolError";
  }
}

const FRAME_TAGS: Record<FrameKind, number> = {
  hello: 1,
  appendBatch: 2,
  ack: 3,
  snapshotChunk: 4,
  cursorWatermark: 5,
  heartbeat: 6,
  error: 7,
};

const FRAME_ENTRIES: ReadonlyArray<readonly [FrameKind, number]> = [
  ["hello", 1],
  ["appendBatch", 2],
  ["ack", 3],
  ["snapshotChunk", 4],
  ["cursorWatermark", 5],
  ["heartbeat", 6],
  ["error", 7],
];

const FRAME_BY_TAG = new Map<number, FrameKind>(FRAME_ENTRIES.map(([kind, tag]) => [tag, kind]));

const OP_TAGS: Record<EngineOp["kind"], number> = {
  upsertWorkbook: 1,
  setWorkbookMetadata: 2,
  upsertSheet: 3,
  deleteSheet: 4,
  updateRowMetadata: 5,
  updateColumnMetadata: 6,
  setFreezePane: 7,
  clearFreezePane: 8,
  setFilter: 9,
  clearFilter: 10,
  setSort: 11,
  clearSort: 12,
  setCellValue: 13,
  setCellFormula: 14,
  setCellFormat: 15,
  clearCell: 16,
  upsertDefinedName: 17,
  deleteDefinedName: 18,
  upsertTable: 19,
  deleteTable: 20,
  upsertSpillRange: 21,
  deleteSpillRange: 22,
  upsertPivotTable: 23,
  deletePivotTable: 24,
  setCalculationSettings: 25,
  setVolatileContext: 26,
  insertRows: 27,
  deleteRows: 28,
  moveRows: 29,
  insertColumns: 30,
  deleteColumns: 31,
  moveColumns: 32,
};

type LiteralTag = 0 | 1 | 2 | 3;
type DefinedNameValueTag = 0 | 1 | 2 | 3 | 4 | 5;

function assertNever(value: never): never {
  throw new BinaryProtocolError(`Unsupported value: ${String(value)}`);
}

function decodeLiteralTag(tag: number): LiteralTag {
  switch (tag) {
    case 0:
    case 1:
    case 2:
    case 3:
      return tag;
    default:
      throw new BinaryProtocolError(`Unknown literal tag ${tag}`);
  }
}

function decodeDefinedNameValueTag(tag: number): DefinedNameValueTag {
  switch (tag) {
    case 0:
    case 1:
    case 2:
    case 3:
    case 4:
    case 5:
      return tag;
    default:
      throw new BinaryProtocolError(`Unknown defined-name value tag ${tag}`);
  }
}

export class BinaryWriter {
  private readonly chunks: number[] = [];

  u8(value: number): void {
    this.chunks.push(value & 0xff);
  }

  u32(value: number): void {
    const next = value >>> 0;
    this.chunks.push(next & 0xff, (next >>> 8) & 0xff, (next >>> 16) & 0xff, (next >>> 24) & 0xff);
  }

  f64(value: number): void {
    const buffer = new ArrayBuffer(8);
    new DataView(buffer).setFloat64(0, value, true);
    this.bytes(new Uint8Array(buffer));
  }

  bool(value: boolean): void {
    this.u8(value ? 1 : 0);
  }

  string(value: string): void {
    this.bytes(textEncoder.encode(value));
  }

  stringArray(values: readonly string[]): void {
    this.u32(values.length);
    values.forEach((value) => this.string(value));
  }

  bytes(value: Uint8Array): void {
    this.u32(value.byteLength);
    value.forEach((byte) => this.chunks.push(byte));
  }

  finish(): Uint8Array {
    return Uint8Array.from(this.chunks);
  }
}

export class BinaryReader {
  private offset = 0;

  constructor(private readonly bytes: Uint8Array) {}

  private ensure(size: number): void {
    if (this.offset + size > this.bytes.byteLength) {
      throw new BinaryProtocolError("Unexpected end of binary frame");
    }
  }

  u8(): number {
    this.ensure(1);
    const value = this.bytes[this.offset]!;
    this.offset += 1;
    return value;
  }

  u32(): number {
    this.ensure(4);
    const view = new DataView(this.bytes.buffer, this.bytes.byteOffset + this.offset, 4);
    const value = view.getUint32(0, true);
    this.offset += 4;
    return value;
  }

  f64(): number {
    const buffer = this.bytesView();
    return new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength).getFloat64(0, true);
  }

  bool(): boolean {
    return this.u8() === 1;
  }

  string(): string {
    return textDecoder.decode(this.bytesView());
  }

  stringArray(): string[] {
    const count = this.u32();
    const values: string[] = [];
    for (let index = 0; index < count; index += 1) {
      values.push(this.string());
    }
    return values;
  }

  bytesView(): Uint8Array {
    const length = this.u32();
    this.ensure(length);
    const slice = this.bytes.subarray(this.offset, this.offset + length);
    this.offset += length;
    return slice;
  }

  done(): boolean {
    return this.offset === this.bytes.byteLength;
  }
}

function encodeLiteral(writer: BinaryWriter, literal: LiteralInput): void {
  if (literal === null) {
    writer.u8(0 satisfies LiteralTag);
    return;
  }

  if (typeof literal === "number") {
    writer.u8(1 satisfies LiteralTag);
    writer.f64(literal);
    return;
  }
  if (typeof literal === "string") {
    writer.u8(2 satisfies LiteralTag);
    writer.string(literal);
    return;
  }
  if (typeof literal === "boolean") {
    writer.u8(3 satisfies LiteralTag);
    writer.bool(literal);
    return;
  }
  throw new BinaryProtocolError(`Unsupported literal type: ${typeof literal}`);
}

function decodeLiteral(reader: BinaryReader): LiteralInput {
  switch (decodeLiteralTag(reader.u8())) {
    case 0:
      return null;
    case 1:
      return reader.f64();
    case 2:
      return reader.string();
    case 3:
      return reader.bool();
    default:
      throw new BinaryProtocolError("Unknown literal tag");
  }
}

function encodeDefinedNameValue(
  writer: BinaryWriter,
  value: WorkbookDefinedNameValueSnapshot,
): void {
  if (typeof value !== "object" || value === null || !("kind" in value)) {
    writer.u8(0 satisfies DefinedNameValueTag);
    encodeLiteral(writer, value);
    return;
  }
  switch (value.kind) {
    case "scalar":
      writer.u8(1 satisfies DefinedNameValueTag);
      encodeLiteral(writer, value.value);
      return;
    case "cell-ref":
      writer.u8(2 satisfies DefinedNameValueTag);
      writer.string(value.sheetName);
      writer.string(value.address);
      return;
    case "range-ref":
      writer.u8(3 satisfies DefinedNameValueTag);
      encodeCellRangeRef(writer, value);
      return;
    case "structured-ref":
      writer.u8(4 satisfies DefinedNameValueTag);
      writer.string(value.tableName);
      writer.string(value.columnName);
      return;
    case "formula":
      writer.u8(5 satisfies DefinedNameValueTag);
      writer.string(value.formula);
      return;
  }
}

function decodeDefinedNameValue(reader: BinaryReader): WorkbookDefinedNameValueSnapshot {
  switch (decodeDefinedNameValueTag(reader.u8())) {
    case 0:
      return decodeLiteral(reader);
    case 1:
      return { kind: "scalar", value: decodeLiteral(reader) };
    case 2:
      return { kind: "cell-ref", sheetName: reader.string(), address: reader.string() };
    case 3:
      return { kind: "range-ref", ...decodeCellRangeRef(reader) };
    case 4:
      return { kind: "structured-ref", tableName: reader.string(), columnName: reader.string() };
    case 5:
      return { kind: "formula", formula: reader.string() };
    default:
      throw new BinaryProtocolError("Unknown defined-name value tag");
  }
}

function encodeNullableNumber(writer: BinaryWriter, value: number | null): void {
  writer.bool(value !== null);
  if (value !== null) {
    writer.f64(value);
  }
}

function decodeNullableNumber(reader: BinaryReader): number | null {
  return reader.bool() ? reader.f64() : null;
}

function encodeNullableBoolean(writer: BinaryWriter, value: boolean | null): void {
  writer.u8(value === null ? 0 : value ? 2 : 1);
}

function decodeNullableBoolean(reader: BinaryReader): boolean | null {
  switch (reader.u8()) {
    case 0:
      return null;
    case 1:
      return false;
    case 2:
      return true;
    default:
      throw new BinaryProtocolError("Unknown nullable boolean tag");
  }
}

function encodeCellRangeRef(writer: BinaryWriter, ref: CellRangeRef): void {
  writer.string(ref.sheetName);
  writer.string(ref.startAddress);
  writer.string(ref.endAddress);
}

function decodeCellRangeRef(reader: BinaryReader): CellRangeRef {
  return {
    sheetName: reader.string(),
    startAddress: reader.string(),
    endAddress: reader.string(),
  };
}

function encodeSortDirection(writer: BinaryWriter, direction: WorkbookSortDirection): void {
  switch (direction) {
    case "asc":
      writer.u8(1);
      return;
    case "desc":
      writer.u8(2);
      return;
  }
}

function encodeCalculationMode(writer: BinaryWriter, mode: WorkbookCalculationMode): void {
  writer.u8(mode === "manual" ? 2 : 1);
}

function decodeCalculationMode(reader: BinaryReader): WorkbookCalculationMode {
  return reader.u8() === 2 ? "manual" : "automatic";
}

function encodeCompatibilityMode(writer: BinaryWriter, mode: CompatibilityMode): void {
  writer.u8(mode === "odf-1.4" ? 2 : 1);
}

function decodeCompatibilityMode(reader: BinaryReader): CompatibilityMode {
  return reader.u8() === 2 ? "odf-1.4" : "excel-modern";
}

function encodeAxisEntries(
  writer: BinaryWriter,
  entries: readonly WorkbookAxisEntrySnapshot[] | undefined,
): void {
  writer.u32(entries?.length ?? 0);
  entries?.forEach((entry) => {
    writer.string(entry.id);
    writer.u32(entry.index);
    encodeNullableNumber(writer, entry.size ?? null);
    encodeNullableBoolean(writer, entry.hidden ?? null);
  });
}

function decodeAxisEntries(reader: BinaryReader): WorkbookAxisEntrySnapshot[] {
  const count = reader.u32();
  const entries: WorkbookAxisEntrySnapshot[] = [];
  for (let index = 0; index < count; index += 1) {
    const entry: WorkbookAxisEntrySnapshot = {
      id: reader.string(),
      index: reader.u32(),
    };
    const size = decodeNullableNumber(reader);
    const hidden = decodeNullableBoolean(reader);
    if (size !== null) {
      entry.size = size;
    }
    if (hidden !== null) {
      entry.hidden = hidden;
    }
    entries.push(entry);
  }
  return entries;
}

function decodeSortDirection(reader: BinaryReader): WorkbookSortDirection {
  switch (reader.u8()) {
    case 1:
      return "asc";
    case 2:
      return "desc";
    default:
      throw new BinaryProtocolError("Unknown sort direction tag");
  }
}

function encodeSortKey(writer: BinaryWriter, key: WorkbookSortKey): void {
  writer.string(key.keyAddress);
  encodeSortDirection(writer, key.direction);
}

function decodeSortKey(reader: BinaryReader): WorkbookSortKey {
  return {
    keyAddress: reader.string(),
    direction: decodeSortDirection(reader),
  };
}

function encodePivotAggregation(writer: BinaryWriter, agg: PivotAggregation): void {
  switch (agg) {
    case "sum":
      writer.u8(1);
      return;
    case "count":
      writer.u8(2);
      return;
  }
}

function decodePivotAggregation(reader: BinaryReader): PivotAggregation {
  switch (reader.u8()) {
    case 1:
      return "sum";
    case 2:
      return "count";
    default:
      throw new BinaryProtocolError("Unknown pivot aggregation tag");
  }
}

function encodePivotValue(writer: BinaryWriter, value: WorkbookPivotValueSnapshot): void {
  writer.string(value.sourceColumn);
  encodePivotAggregation(writer, value.summarizeBy);
  writer.bool(value.outputLabel !== undefined);
  if (value.outputLabel !== undefined) {
    writer.string(value.outputLabel);
  }
}

function decodePivotValue(reader: BinaryReader): WorkbookPivotValueSnapshot {
  const sourceColumn = reader.string();
  const summarizeBy = decodePivotAggregation(reader);
  const hasLabel = reader.bool();
  const result: WorkbookPivotValueSnapshot = {
    sourceColumn,
    summarizeBy,
  };
  if (hasLabel) {
    result.outputLabel = reader.string();
  }
  return result;
}

function encodeTable(writer: BinaryWriter, table: WorkbookTableOp): void {
  writer.string(table.name);
  writer.string(table.sheetName);
  writer.string(table.startAddress);
  writer.string(table.endAddress);
  writer.stringArray(table.columnNames);
  writer.bool(table.headerRow);
  writer.bool(table.totalsRow);
}

function decodeTable(reader: BinaryReader): WorkbookTableOp {
  return {
    name: reader.string(),
    sheetName: reader.string(),
    startAddress: reader.string(),
    endAddress: reader.string(),
    columnNames: reader.stringArray(),
    headerRow: reader.bool(),
    totalsRow: reader.bool(),
  };
}

function encodeEngineOp(writer: BinaryWriter, op: EngineOp): void {
  writer.u8(OP_TAGS[op.kind]);
  switch (op.kind) {
    case "upsertWorkbook":
      writer.string(op.name);
      return;
    case "setWorkbookMetadata":
      writer.string(op.key);
      encodeLiteral(writer, op.value);
      return;
    case "setCalculationSettings":
      encodeCalculationMode(writer, op.settings.mode);
      encodeCompatibilityMode(writer, op.settings.compatibilityMode ?? "excel-modern");
      return;
    case "setVolatileContext":
      writer.u32(op.context.recalcEpoch);
      return;
    case "upsertSheet":
      writer.string(op.name);
      writer.u32(op.order);
      return;
    case "deleteSheet":
      writer.string(op.name);
      return;
    case "insertRows":
    case "insertColumns":
      writer.string(op.sheetName);
      writer.u32(op.start);
      writer.u32(op.count);
      encodeAxisEntries(writer, op.entries);
      return;
    case "deleteRows":
    case "deleteColumns":
      writer.string(op.sheetName);
      writer.u32(op.start);
      writer.u32(op.count);
      return;
    case "moveRows":
    case "moveColumns":
      writer.string(op.sheetName);
      writer.u32(op.start);
      writer.u32(op.count);
      writer.u32(op.target);
      return;
    case "updateRowMetadata":
    case "updateColumnMetadata":
      writer.string(op.sheetName);
      writer.u32(op.start);
      writer.u32(op.count);
      encodeNullableNumber(writer, op.size);
      encodeNullableBoolean(writer, op.hidden);
      return;
    case "setFreezePane":
      writer.string(op.sheetName);
      writer.u32(op.rows);
      writer.u32(op.cols);
      return;
    case "clearFreezePane":
      writer.string(op.sheetName);
      return;
    case "setFilter":
    case "clearFilter":
      writer.string(op.sheetName);
      encodeCellRangeRef(writer, op.range);
      return;
    case "setSort":
      writer.string(op.sheetName);
      encodeCellRangeRef(writer, op.range);
      writer.u32(op.keys.length);
      op.keys.forEach((key) => encodeSortKey(writer, key));
      return;
    case "clearSort":
      writer.string(op.sheetName);
      encodeCellRangeRef(writer, op.range);
      return;
    case "setCellValue":
      writer.string(op.sheetName);
      writer.string(op.address);
      encodeLiteral(writer, op.value);
      return;
    case "setCellFormula":
      writer.string(op.sheetName);
      writer.string(op.address);
      writer.string(op.formula);
      return;
    case "setCellFormat":
      writer.string(op.sheetName);
      writer.string(op.address);
      writer.bool(op.format !== null);
      if (op.format !== null) {
        writer.string(op.format);
      }
      return;
    case "clearCell":
      writer.string(op.sheetName);
      writer.string(op.address);
      return;
    case "upsertDefinedName":
      writer.string(op.name);
      encodeDefinedNameValue(writer, op.value);
      return;
    case "deleteDefinedName":
      writer.string(op.name);
      return;
    case "upsertTable":
      encodeTable(writer, op.table);
      return;
    case "deleteTable":
      writer.string(op.name);
      return;
    case "upsertSpillRange":
      writer.string(op.sheetName);
      writer.string(op.address);
      writer.u32(op.rows);
      writer.u32(op.cols);
      return;
    case "deleteSpillRange":
      writer.string(op.sheetName);
      writer.string(op.address);
      return;
    case "upsertPivotTable":
      writer.string(op.name);
      writer.string(op.sheetName);
      writer.string(op.address);
      encodeCellRangeRef(writer, op.source);
      writer.stringArray(op.groupBy);
      writer.u32(op.values.length);
      op.values.forEach((v) => encodePivotValue(writer, v));
      writer.u32(op.rows);
      writer.u32(op.cols);
      return;
    case "deletePivotTable":
      writer.string(op.sheetName);
      writer.string(op.address);
      return;
    default:
      assertNever(op);
  }
}

function decodeEngineOp(reader: BinaryReader): EngineOp {
  switch (reader.u8()) {
    case 1:
      return { kind: "upsertWorkbook", name: reader.string() };
    case 2:
      return { kind: "setWorkbookMetadata", key: reader.string(), value: decodeLiteral(reader) };
    case 25:
      return {
        kind: "setCalculationSettings",
        settings: {
          mode: decodeCalculationMode(reader),
          compatibilityMode: decodeCompatibilityMode(reader),
        },
      };
    case 26:
      return { kind: "setVolatileContext", context: { recalcEpoch: reader.u32() } };
    case 3:
      return { kind: "upsertSheet", name: reader.string(), order: reader.u32() };
    case 4:
      return { kind: "deleteSheet", name: reader.string() };
    case 27:
      return {
        kind: "insertRows",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        entries: decodeAxisEntries(reader),
      };
    case 28:
      return {
        kind: "deleteRows",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
      };
    case 29:
      return {
        kind: "moveRows",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        target: reader.u32(),
      };
    case 30:
      return {
        kind: "insertColumns",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        entries: decodeAxisEntries(reader),
      };
    case 31:
      return {
        kind: "deleteColumns",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
      };
    case 32:
      return {
        kind: "moveColumns",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        target: reader.u32(),
      };
    case 5:
      return {
        kind: "updateRowMetadata",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        size: decodeNullableNumber(reader),
        hidden: decodeNullableBoolean(reader),
      };
    case 6:
      return {
        kind: "updateColumnMetadata",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        size: decodeNullableNumber(reader),
        hidden: decodeNullableBoolean(reader),
      };
    case 7:
      return {
        kind: "setFreezePane",
        sheetName: reader.string(),
        rows: reader.u32(),
        cols: reader.u32(),
      };
    case 8:
      return { kind: "clearFreezePane", sheetName: reader.string() };
    case 9:
      return {
        kind: "setFilter",
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
      };
    case 10:
      return {
        kind: "clearFilter",
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
      };
    case 11:
      return {
        kind: "setSort",
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
        keys: (() => {
          const count = reader.u32();
          const keys: WorkbookSortKey[] = [];
          for (let index = 0; index < count; index += 1) {
            keys.push(decodeSortKey(reader));
          }
          return keys;
        })(),
      };
    case 12:
      return {
        kind: "clearSort",
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
      };
    case 13:
      return {
        kind: "setCellValue",
        sheetName: reader.string(),
        address: reader.string(),
        value: decodeLiteral(reader),
      };
    case 14:
      return {
        kind: "setCellFormula",
        sheetName: reader.string(),
        address: reader.string(),
        formula: reader.string(),
      };
    case 15: {
      const sheetName = reader.string();
      const address = reader.string();
      const hasFormat = reader.bool();
      return {
        kind: "setCellFormat",
        sheetName,
        address,
        format: hasFormat ? reader.string() : null,
      };
    }
    case 16:
      return { kind: "clearCell", sheetName: reader.string(), address: reader.string() };
    case 17:
      return {
        kind: "upsertDefinedName",
        name: reader.string(),
        value: decodeDefinedNameValue(reader),
      };
    case 18:
      return { kind: "deleteDefinedName", name: reader.string() };
    case 19:
      return { kind: "upsertTable", table: decodeTable(reader) };
    case 20:
      return { kind: "deleteTable", name: reader.string() };
    case 21:
      return {
        kind: "upsertSpillRange",
        sheetName: reader.string(),
        address: reader.string(),
        rows: reader.u32(),
        cols: reader.u32(),
      };
    case 22:
      return {
        kind: "deleteSpillRange",
        sheetName: reader.string(),
        address: reader.string(),
      };
    case 23:
      return {
        kind: "upsertPivotTable",
        name: reader.string(),
        sheetName: reader.string(),
        address: reader.string(),
        source: decodeCellRangeRef(reader),
        groupBy: reader.stringArray(),
        values: (() => {
          const count = reader.u32();
          const values: WorkbookPivotValueSnapshot[] = [];
          for (let i = 0; i < count; i++) values.push(decodePivotValue(reader));
          return values;
        })(),
        rows: reader.u32(),
        cols: reader.u32(),
      };
    case 24:
      return {
        kind: "deletePivotTable",
        sheetName: reader.string(),
        address: reader.string(),
      };
    default:
      throw new BinaryProtocolError("Unknown engine op tag");
  }
}

function encodeBatch(writer: BinaryWriter, batch: EngineOpBatch): void {
  writer.string(batch.id);
  writer.string(batch.replicaId);
  writer.u32(batch.clock.counter);
  writer.u32(batch.ops.length);
  batch.ops.forEach((op) => encodeEngineOp(writer, op));
}

function decodeBatch(reader: BinaryReader): EngineOpBatch {
  const id = reader.string();
  const replicaId = reader.string();
  const counter = reader.u32();
  const opCount = reader.u32();
  const ops: EngineOp[] = [];
  for (let index = 0; index < opCount; index += 1) {
    ops.push(decodeEngineOp(reader));
  }
  return {
    id,
    replicaId,
    clock: { counter },
    ops,
  };
}

function encodePayload(frame: ProtocolFrame): Uint8Array {
  const writer = new BinaryWriter();
  switch (frame.kind) {
    case "hello":
      writer.string(frame.documentId);
      writer.string(frame.replicaId);
      writer.string(frame.sessionId);
      writer.u32(frame.protocolVersion);
      writer.u32(frame.lastServerCursor);
      writer.stringArray(frame.capabilities);
      return writer.finish();
    case "appendBatch":
      writer.string(frame.documentId);
      writer.u32(frame.cursor);
      encodeBatch(writer, frame.batch);
      return writer.finish();
    case "ack":
      writer.string(frame.documentId);
      writer.string(frame.batchId);
      writer.u32(frame.cursor);
      writer.f64(frame.acceptedAtUnixMs);
      return writer.finish();
    case "snapshotChunk":
      writer.string(frame.documentId);
      writer.string(frame.snapshotId);
      writer.u32(frame.cursor);
      writer.u32(frame.chunkIndex);
      writer.u32(frame.chunkCount);
      writer.string(frame.contentType);
      writer.bytes(frame.bytes);
      return writer.finish();
    case "cursorWatermark":
      writer.string(frame.documentId);
      writer.u32(frame.cursor);
      writer.u32(frame.compactedCursor);
      return writer.finish();
    case "heartbeat":
      writer.string(frame.documentId);
      writer.u32(frame.cursor);
      writer.f64(frame.sentAtUnixMs);
      return writer.finish();
    case "error":
      writer.string(frame.documentId);
      writer.string(frame.code);
      writer.string(frame.message);
      writer.bool(frame.retryable);
      return writer.finish();
    default:
      assertNever(frame);
  }
}

function decodePayload(kind: FrameKind, payload: Uint8Array): ProtocolFrame {
  const reader = new BinaryReader(payload);
  switch (kind) {
    case "hello":
      return {
        kind,
        documentId: reader.string(),
        replicaId: reader.string(),
        sessionId: reader.string(),
        protocolVersion: reader.u32(),
        lastServerCursor: reader.u32(),
        capabilities: reader.stringArray(),
      };
    case "appendBatch":
      return {
        kind,
        documentId: reader.string(),
        cursor: reader.u32(),
        batch: decodeBatch(reader),
      };
    case "ack":
      return {
        kind,
        documentId: reader.string(),
        batchId: reader.string(),
        cursor: reader.u32(),
        acceptedAtUnixMs: reader.f64(),
      };
    case "snapshotChunk":
      return {
        kind,
        documentId: reader.string(),
        snapshotId: reader.string(),
        cursor: reader.u32(),
        chunkIndex: reader.u32(),
        chunkCount: reader.u32(),
        contentType: reader.string(),
        bytes: reader.bytesView(),
      };
    case "cursorWatermark":
      return {
        kind,
        documentId: reader.string(),
        cursor: reader.u32(),
        compactedCursor: reader.u32(),
      };
    case "heartbeat":
      return {
        kind,
        documentId: reader.string(),
        cursor: reader.u32(),
        sentAtUnixMs: reader.f64(),
      };
    case "error":
      return {
        kind,
        documentId: reader.string(),
        code: reader.string(),
        message: reader.string(),
        retryable: reader.bool(),
      };
    default:
      assertNever(kind);
  }
}

export function encodeFrame(frame: ProtocolFrame): Uint8Array {
  const payload = encodePayload(frame);
  const output = new Uint8Array(11 + payload.byteLength);
  const view = new DataView(output.buffer);
  view.setUint32(0, PROTOCOL_MAGIC, true);
  view.setUint16(4, PROTOCOL_VERSION, true);
  view.setUint8(6, FRAME_TAGS[frame.kind]);
  view.setUint32(7, payload.byteLength, true);
  output.set(payload, 11);
  return output;
}

export function decodeFrame(bytesLike: Uint8Array | ArrayBuffer): ProtocolFrame {
  const bytes = bytesLike instanceof Uint8Array ? bytesLike : new Uint8Array(bytesLike);
  if (bytes.byteLength < 11) {
    throw new BinaryProtocolError("Binary frame too short");
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const magic = view.getUint32(0, true);
  if (magic !== PROTOCOL_MAGIC) {
    throw new BinaryProtocolError("Binary frame magic mismatch");
  }

  const version = view.getUint16(4, true);
  if (version !== PROTOCOL_VERSION) {
    throw new BinaryProtocolError(`Unsupported protocol version ${version}`);
  }

  const kind = FRAME_BY_TAG.get(view.getUint8(6));
  if (!kind) {
    throw new BinaryProtocolError("Unknown frame tag");
  }

  const payloadLength = view.getUint32(7, true);
  if (bytes.byteLength !== 11 + payloadLength) {
    throw new BinaryProtocolError("Binary frame length mismatch");
  }

  return decodePayload(kind, bytes.subarray(11));
}

export function createSnapshotChunkFrames(options: SnapshotChunkOptions): SnapshotChunkFrame[] {
  const chunkSize = Math.max(1, options.chunkSize ?? 64 * 1024);
  const chunkCount = Math.max(1, Math.ceil(options.bytes.byteLength / chunkSize));
  const frames: SnapshotChunkFrame[] = [];
  for (let chunkIndex = 0; chunkIndex < chunkCount; chunkIndex += 1) {
    const start = chunkIndex * chunkSize;
    const end = Math.min(options.bytes.byteLength, start + chunkSize);
    frames.push({
      kind: "snapshotChunk",
      documentId: options.documentId,
      snapshotId: options.snapshotId,
      cursor: options.cursor,
      chunkIndex,
      chunkCount,
      contentType: options.contentType,
      bytes: options.bytes.subarray(start, end),
    });
  }
  return frames;
}
