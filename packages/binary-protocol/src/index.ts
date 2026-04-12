import type {
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellNumberFormatKind,
  CellNumberFormatRecord,
  CellRangeRef,
  CellStylePatch,
  CompatibilityMode,
  CellStyleRecord,
  CellVerticalAlignment,
  PivotAggregation,
  LiteralInput,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatRuleSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationRuleSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookNoteSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationMode,
  WorkbookDefinedNameValueSnapshot,
  WorkbookValidationListSourceSnapshot,
  WorkbookPivotValueSnapshot,
} from "@bilig/protocol";
import type {
  EngineOp,
  EngineOpBatch,
  WorkbookSortDirection,
  WorkbookSortKey,
  WorkbookTableOp,
} from "@bilig/workbook-domain";

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
  renameSheet: 37,
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
  upsertCellStyle: 16,
  setStyleRange: 17,
  upsertCellNumberFormat: 35,
  setFormatRange: 36,
  clearCell: 18,
  upsertDefinedName: 19,
  deleteDefinedName: 20,
  upsertTable: 21,
  deleteTable: 22,
  upsertSpillRange: 23,
  deleteSpillRange: 24,
  upsertPivotTable: 25,
  deletePivotTable: 26,
  setCalculationSettings: 27,
  setVolatileContext: 28,
  insertRows: 29,
  deleteRows: 30,
  moveRows: 31,
  insertColumns: 32,
  deleteColumns: 33,
  moveColumns: 34,
  setSheetProtection: 46,
  clearSheetProtection: 47,
  setDataValidation: 38,
  clearDataValidation: 39,
  upsertCommentThread: 40,
  deleteCommentThread: 41,
  upsertNote: 42,
  deleteNote: 43,
  upsertConditionalFormat: 44,
  deleteConditionalFormat: 45,
  upsertRangeProtection: 48,
  deleteRangeProtection: 49,
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

function decodeValidationComparisonOperator(
  value: string,
):
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual" {
  switch (value) {
    case "between":
    case "notBetween":
    case "equal":
    case "notEqual":
    case "greaterThan":
    case "greaterThanOrEqual":
    case "lessThan":
    case "lessThanOrEqual":
      return value;
    default:
      throw new BinaryProtocolError("Unknown data validation comparison operator");
  }
}

function decodeValidationErrorStyle(value: string): "stop" | "warning" | "information" {
  switch (value) {
    case "stop":
    case "warning":
    case "information":
      return value;
    default:
      throw new BinaryProtocolError("Unknown data validation error style");
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

function encodeOptionalNullableNumber(
  writer: BinaryWriter,
  value: number | null | undefined,
): void {
  writer.bool(value !== undefined);
  if (value !== undefined) {
    encodeNullableNumber(writer, value);
  }
}

function decodeOptionalNullableNumber(reader: BinaryReader): number | null | undefined {
  return reader.bool() ? decodeNullableNumber(reader) : undefined;
}

function encodeOptionalNullableBoolean(
  writer: BinaryWriter,
  value: boolean | null | undefined,
): void {
  writer.bool(value !== undefined);
  if (value !== undefined) {
    encodeNullableBoolean(writer, value);
  }
}

function decodeOptionalNullableBoolean(reader: BinaryReader): boolean | null | undefined {
  return reader.bool() ? decodeNullableBoolean(reader) : undefined;
}

function encodeOptionalNullableString(
  writer: BinaryWriter,
  value: string | null | undefined,
): void {
  writer.bool(value !== undefined);
  if (value !== undefined) {
    writer.bool(value !== null);
    if (value !== null) {
      writer.string(value);
    }
  }
}

function decodeOptionalNullableString(reader: BinaryReader): string | null | undefined {
  if (!reader.bool()) {
    return undefined;
  }
  return reader.bool() ? reader.string() : null;
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

function encodeCellStyleRecord(writer: BinaryWriter, style: CellStyleRecord): void {
  writer.string(style.id);
  writer.bool(style.fill !== undefined);
  if (style.fill) {
    writer.string(style.fill.backgroundColor);
  }
  writer.bool(style.font !== undefined);
  if (style.font) {
    writer.bool(style.font.family !== undefined);
    if (style.font.family !== undefined) {
      writer.string(style.font.family);
    }
    writer.bool(style.font.size !== undefined);
    if (style.font.size !== undefined) {
      writer.f64(style.font.size);
    }
    writer.bool(style.font.bold !== undefined);
    if (style.font.bold !== undefined) {
      writer.bool(style.font.bold);
    }
    writer.bool(style.font.italic !== undefined);
    if (style.font.italic !== undefined) {
      writer.bool(style.font.italic);
    }
    writer.bool(style.font.underline !== undefined);
    if (style.font.underline !== undefined) {
      writer.bool(style.font.underline);
    }
    writer.bool(style.font.color !== undefined);
    if (style.font.color !== undefined) {
      writer.string(style.font.color);
    }
  }
  writer.bool(style.alignment !== undefined);
  if (style.alignment) {
    writer.bool(style.alignment.horizontal !== undefined);
    if (style.alignment.horizontal !== undefined) {
      writer.string(style.alignment.horizontal);
    }
    writer.bool(style.alignment.vertical !== undefined);
    if (style.alignment.vertical !== undefined) {
      writer.string(style.alignment.vertical);
    }
    writer.bool(style.alignment.wrap !== undefined);
    if (style.alignment.wrap !== undefined) {
      writer.bool(style.alignment.wrap);
    }
    writer.bool(style.alignment.indent !== undefined);
    if (style.alignment.indent !== undefined) {
      writer.u32(style.alignment.indent);
    }
  }
  writer.bool(style.borders !== undefined);
  if (style.borders) {
    encodeBorderSide(writer, style.borders.top);
    encodeBorderSide(writer, style.borders.right);
    encodeBorderSide(writer, style.borders.bottom);
    encodeBorderSide(writer, style.borders.left);
  }
}

function encodeCellNumberFormatRecord(writer: BinaryWriter, format: CellNumberFormatRecord): void {
  writer.string(format.id);
  writer.string(format.code);
  writer.string(format.kind);
}

function decodeCellNumberFormatRecord(reader: BinaryReader): CellNumberFormatRecord {
  const id = reader.string();
  const code = reader.string();
  const kind = decodeCellNumberFormatKind(reader.string());
  return {
    id,
    code,
    kind,
  };
}

function decodeCellStyleRecord(reader: BinaryReader): CellStyleRecord {
  const style: CellStyleRecord = { id: reader.string() };
  if (reader.bool()) {
    style.fill = { backgroundColor: reader.string() };
  }
  if (reader.bool()) {
    const font: NonNullable<CellStyleRecord["font"]> = {};
    if (reader.bool()) {
      font.family = reader.string();
    }
    if (reader.bool()) {
      font.size = reader.f64();
    }
    if (reader.bool()) {
      font.bold = reader.bool();
    }
    if (reader.bool()) {
      font.italic = reader.bool();
    }
    if (reader.bool()) {
      font.underline = reader.bool();
    }
    if (reader.bool()) {
      font.color = reader.string();
    }
    style.font = font;
  }
  if (reader.bool()) {
    const alignment: NonNullable<CellStyleRecord["alignment"]> = {};
    if (reader.bool()) {
      const horizontal = decodeHorizontalAlignment(reader.string());
      if (horizontal !== undefined) {
        alignment.horizontal = horizontal;
      }
    }
    if (reader.bool()) {
      const vertical = decodeVerticalAlignment(reader.string());
      if (vertical !== undefined) {
        alignment.vertical = vertical;
      }
    }
    if (reader.bool()) {
      alignment.wrap = reader.bool();
    }
    if (reader.bool()) {
      alignment.indent = reader.u32();
    }
    style.alignment = alignment;
  }
  if (reader.bool()) {
    const top = decodeBorderSide(reader);
    const right = decodeBorderSide(reader);
    const bottom = decodeBorderSide(reader);
    const left = decodeBorderSide(reader);
    style.borders = {
      ...(top ? { top } : {}),
      ...(right ? { right } : {}),
      ...(bottom ? { bottom } : {}),
      ...(left ? { left } : {}),
    };
  }
  return style;
}

function encodeCellStylePatch(writer: BinaryWriter, patch: CellStylePatch): void {
  writer.bool(patch.fill !== undefined);
  if (patch.fill !== undefined) {
    writer.bool(patch.fill !== null);
    if (patch.fill !== null) {
      encodeOptionalNullableString(writer, patch.fill.backgroundColor);
    }
  }
  writer.bool(patch.font !== undefined);
  if (patch.font !== undefined) {
    writer.bool(patch.font !== null);
    if (patch.font !== null) {
      encodeOptionalNullableString(writer, patch.font.family);
      encodeOptionalNullableNumber(writer, patch.font.size);
      encodeOptionalNullableBoolean(writer, patch.font.bold);
      encodeOptionalNullableBoolean(writer, patch.font.italic);
      encodeOptionalNullableBoolean(writer, patch.font.underline);
      encodeOptionalNullableString(writer, patch.font.color);
    }
  }
  writer.bool(patch.alignment !== undefined);
  if (patch.alignment !== undefined) {
    writer.bool(patch.alignment !== null);
    if (patch.alignment !== null) {
      encodeOptionalNullableString(writer, patch.alignment.horizontal);
      encodeOptionalNullableString(writer, patch.alignment.vertical);
      encodeOptionalNullableBoolean(writer, patch.alignment.wrap);
      encodeOptionalNullableNumber(writer, patch.alignment.indent);
    }
  }
  writer.bool(patch.borders !== undefined);
  if (patch.borders !== undefined) {
    writer.bool(patch.borders !== null);
    if (patch.borders !== null) {
      encodePatchBorderSide(writer, patch.borders.top);
      encodePatchBorderSide(writer, patch.borders.right);
      encodePatchBorderSide(writer, patch.borders.bottom);
      encodePatchBorderSide(writer, patch.borders.left);
    }
  }
}

function decodeCellStylePatch(reader: BinaryReader): CellStylePatch {
  const patch: CellStylePatch = {};
  if (reader.bool()) {
    if (reader.bool()) {
      const backgroundColor = decodeOptionalNullableString(reader);
      patch.fill = backgroundColor === undefined ? {} : { backgroundColor };
    } else {
      patch.fill = null;
    }
  }
  if (reader.bool()) {
    if (reader.bool()) {
      const font: NonNullable<CellStylePatch["font"]> = {};
      const family = decodeOptionalNullableString(reader);
      if (family !== undefined) {
        font.family = family;
      }
      const size = decodeOptionalNullableNumber(reader);
      if (size !== undefined) {
        font.size = size;
      }
      const bold = decodeOptionalNullableBoolean(reader);
      if (bold !== undefined) {
        font.bold = bold;
      }
      const italic = decodeOptionalNullableBoolean(reader);
      if (italic !== undefined) {
        font.italic = italic;
      }
      const underline = decodeOptionalNullableBoolean(reader);
      if (underline !== undefined) {
        font.underline = underline;
      }
      const color = decodeOptionalNullableString(reader);
      if (color !== undefined) {
        font.color = color;
      }
      patch.font = font;
    } else {
      patch.font = null;
    }
  }
  if (reader.bool()) {
    if (reader.bool()) {
      const alignment: NonNullable<CellStylePatch["alignment"]> = {};
      const horizontal = decodeOptionalNullableString(reader);
      if (horizontal !== undefined) {
        alignment.horizontal =
          horizontal === null ? null : (decodeHorizontalAlignment(horizontal) ?? null);
      }
      const vertical = decodeOptionalNullableString(reader);
      if (vertical !== undefined) {
        alignment.vertical = vertical === null ? null : (decodeVerticalAlignment(vertical) ?? null);
      }
      const wrap = decodeOptionalNullableBoolean(reader);
      if (wrap !== undefined) {
        alignment.wrap = wrap;
      }
      const indent = decodeOptionalNullableNumber(reader);
      if (indent !== undefined) {
        alignment.indent = indent;
      }
      patch.alignment = alignment;
    } else {
      patch.alignment = null;
    }
  }
  if (reader.bool()) {
    if (reader.bool()) {
      const borders: NonNullable<CellStylePatch["borders"]> = {};
      const top = decodePatchBorderSide(reader);
      if (top !== undefined) {
        borders.top = top;
      }
      const right = decodePatchBorderSide(reader);
      if (right !== undefined) {
        borders.right = right;
      }
      const bottom = decodePatchBorderSide(reader);
      if (bottom !== undefined) {
        borders.bottom = bottom;
      }
      const left = decodePatchBorderSide(reader);
      if (left !== undefined) {
        borders.left = left;
      }
      patch.borders = borders;
    } else {
      patch.borders = null;
    }
  }
  return patch;
}

type EncodedBorderSide = NonNullable<NonNullable<CellStyleRecord["borders"]>["top"]>;

function encodeBorderSide(writer: BinaryWriter, side: EncodedBorderSide | undefined): void {
  writer.bool(side !== undefined);
  if (!side) {
    return;
  }
  writer.string(side.style);
  writer.string(side.weight);
  writer.string(side.color);
}

function decodeBorderSide(reader: BinaryReader): EncodedBorderSide | undefined {
  if (!reader.bool()) {
    return undefined;
  }
  const style = decodeBorderStyle(reader.string());
  const weight = decodeBorderWeight(reader.string());
  return {
    style,
    weight,
    color: reader.string(),
  };
}

type EncodedPatchBorderSide = NonNullable<NonNullable<CellStylePatch["borders"]>["top"]>;

function encodePatchBorderSide(
  writer: BinaryWriter,
  side: EncodedPatchBorderSide | null | undefined,
): void {
  writer.bool(side !== undefined);
  if (side === undefined) {
    return;
  }
  writer.bool(side !== null);
  if (side === null) {
    return;
  }
  encodeOptionalNullableString(writer, side.style);
  encodeOptionalNullableString(writer, side.weight);
  encodeOptionalNullableString(writer, side.color);
}

function decodePatchBorderSide(reader: BinaryReader): EncodedPatchBorderSide | null | undefined {
  if (!reader.bool()) {
    return undefined;
  }
  if (!reader.bool()) {
    return null;
  }
  const side: EncodedPatchBorderSide = {};
  const style = decodeOptionalNullableString(reader);
  if (style !== undefined) {
    side.style = style === null ? null : decodeBorderStyle(style);
  }
  const weight = decodeOptionalNullableString(reader);
  if (weight !== undefined) {
    side.weight = weight === null ? null : decodeBorderWeight(weight);
  }
  const color = decodeOptionalNullableString(reader);
  if (color !== undefined) {
    side.color = color;
  }
  return side;
}

function decodeCellNumberFormatKind(value: string): CellNumberFormatKind {
  switch (value) {
    case "number":
    case "currency":
    case "accounting":
    case "percent":
    case "date":
    case "time":
    case "datetime":
    case "text":
      return value;
    default:
      return "general";
  }
}

function decodeHorizontalAlignment(value: string): CellHorizontalAlignment | undefined {
  switch (value) {
    case "general":
    case "left":
    case "center":
    case "right":
      return value;
    default:
      return undefined;
  }
}

function decodeVerticalAlignment(value: string): CellVerticalAlignment | undefined {
  switch (value) {
    case "top":
    case "middle":
    case "bottom":
      return value;
    default:
      return undefined;
  }
}

function decodeBorderStyle(value: string): CellBorderStyle {
  switch (value) {
    case "dashed":
    case "dotted":
    case "double":
      return value;
    default:
      return "solid";
  }
}

function decodeBorderWeight(value: string): CellBorderWeight {
  switch (value) {
    case "medium":
    case "thick":
      return value;
    default:
      return "thin";
  }
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

function encodeValidationListSource(
  writer: BinaryWriter,
  source: WorkbookValidationListSourceSnapshot,
): void {
  switch (source.kind) {
    case "named-range":
      writer.u8(0);
      writer.string(source.name);
      return;
    case "cell-ref":
      writer.u8(1);
      writer.string(source.sheetName);
      writer.string(source.address);
      return;
    case "range-ref":
      writer.u8(2);
      writer.string(source.sheetName);
      writer.string(source.startAddress);
      writer.string(source.endAddress);
      return;
    case "structured-ref":
      writer.u8(3);
      writer.string(source.tableName);
      writer.string(source.columnName);
      return;
    default:
      assertNever(source);
  }
}

function decodeValidationListSource(reader: BinaryReader): WorkbookValidationListSourceSnapshot {
  switch (reader.u8()) {
    case 0:
      return { kind: "named-range", name: reader.string() };
    case 1:
      return { kind: "cell-ref", sheetName: reader.string(), address: reader.string() };
    case 2:
      return {
        kind: "range-ref",
        sheetName: reader.string(),
        startAddress: reader.string(),
        endAddress: reader.string(),
      };
    case 3:
      return {
        kind: "structured-ref",
        tableName: reader.string(),
        columnName: reader.string(),
      };
    default:
      throw new BinaryProtocolError("Unknown validation source tag");
  }
}

function encodeDataValidationRule(
  writer: BinaryWriter,
  rule: WorkbookDataValidationRuleSnapshot,
): void {
  switch (rule.kind) {
    case "list":
      writer.u8(0);
      writer.bool(rule.values !== undefined);
      if (rule.values) {
        writer.u32(rule.values.length);
        rule.values.forEach((value) => encodeLiteral(writer, value));
      }
      writer.bool(rule.source !== undefined);
      if (rule.source) {
        encodeValidationListSource(writer, rule.source);
      }
      return;
    case "checkbox":
      writer.u8(1);
      writer.bool(rule.checkedValue !== undefined);
      if (rule.checkedValue !== undefined) {
        encodeLiteral(writer, rule.checkedValue);
      }
      writer.bool(rule.uncheckedValue !== undefined);
      if (rule.uncheckedValue !== undefined) {
        encodeLiteral(writer, rule.uncheckedValue);
      }
      return;
    case "whole":
    case "decimal":
    case "date":
    case "time":
    case "textLength":
      writer.u8(
        {
          whole: 2,
          decimal: 3,
          date: 4,
          time: 5,
          textLength: 6,
        }[rule.kind],
      );
      writer.string(rule.operator);
      writer.u32(rule.values.length);
      rule.values.forEach((value) => encodeLiteral(writer, value));
      return;
    default:
      assertNever(rule);
  }
}

function decodeDataValidationRule(reader: BinaryReader): WorkbookDataValidationRuleSnapshot {
  const tag = reader.u8();
  switch (tag) {
    case 0: {
      const hasValues = reader.bool();
      const values = hasValues
        ? (() => {
            const count = reader.u32();
            const items: LiteralInput[] = [];
            for (let index = 0; index < count; index += 1) {
              items.push(decodeLiteral(reader));
            }
            return items;
          })()
        : undefined;
      const hasSource = reader.bool();
      const rule: Extract<WorkbookDataValidationRuleSnapshot, { kind: "list" }> = {
        kind: "list",
      };
      if (values) {
        rule.values = values;
      }
      if (hasSource) {
        rule.source = decodeValidationListSource(reader);
      }
      return rule;
    }
    case 1: {
      const hasCheckedValue = reader.bool();
      const checkedValue = hasCheckedValue ? decodeLiteral(reader) : undefined;
      const hasUncheckedValue = reader.bool();
      const rule: Extract<WorkbookDataValidationRuleSnapshot, { kind: "checkbox" }> = {
        kind: "checkbox",
      };
      if (checkedValue !== undefined) {
        rule.checkedValue = checkedValue;
      }
      if (hasUncheckedValue) {
        rule.uncheckedValue = decodeLiteral(reader);
      }
      return rule;
    }
    case 2:
    case 3:
    case 4:
    case 5:
    case 6: {
      const operator = decodeValidationComparisonOperator(reader.string());
      const count = reader.u32();
      const values: LiteralInput[] = [];
      for (let index = 0; index < count; index += 1) {
        values.push(decodeLiteral(reader));
      }
      const kind = {
        2: "whole",
        3: "decimal",
        4: "date",
        5: "time",
        6: "textLength",
      } as const satisfies Record<
        2 | 3 | 4 | 5 | 6,
        "whole" | "decimal" | "date" | "time" | "textLength"
      >;
      const nextKind = kind[tag];
      if (!nextKind) {
        throw new BinaryProtocolError("Unknown scalar data validation rule tag");
      }
      return {
        kind: nextKind,
        operator,
        values,
      };
    }
    default:
      throw new BinaryProtocolError("Unknown data validation rule tag");
  }
}

function encodeDataValidation(
  writer: BinaryWriter,
  validation: WorkbookDataValidationSnapshot,
): void {
  encodeCellRangeRef(writer, validation.range);
  encodeDataValidationRule(writer, validation.rule);
  writer.bool(validation.allowBlank !== undefined);
  if (validation.allowBlank !== undefined) {
    writer.bool(validation.allowBlank);
  }
  writer.bool(validation.showDropdown !== undefined);
  if (validation.showDropdown !== undefined) {
    writer.bool(validation.showDropdown);
  }
  writer.bool(validation.promptTitle !== undefined);
  if (validation.promptTitle !== undefined) {
    writer.string(validation.promptTitle);
  }
  writer.bool(validation.promptMessage !== undefined);
  if (validation.promptMessage !== undefined) {
    writer.string(validation.promptMessage);
  }
  writer.bool(validation.errorStyle !== undefined);
  if (validation.errorStyle !== undefined) {
    writer.string(validation.errorStyle);
  }
  writer.bool(validation.errorTitle !== undefined);
  if (validation.errorTitle !== undefined) {
    writer.string(validation.errorTitle);
  }
  writer.bool(validation.errorMessage !== undefined);
  if (validation.errorMessage !== undefined) {
    writer.string(validation.errorMessage);
  }
}

function decodeDataValidation(reader: BinaryReader): WorkbookDataValidationSnapshot {
  const range = decodeCellRangeRef(reader);
  const rule = decodeDataValidationRule(reader);
  const validation: WorkbookDataValidationSnapshot = {
    range,
    rule,
  };
  if (reader.bool()) {
    validation.allowBlank = reader.bool();
  }
  if (reader.bool()) {
    validation.showDropdown = reader.bool();
  }
  if (reader.bool()) {
    validation.promptTitle = reader.string();
  }
  if (reader.bool()) {
    validation.promptMessage = reader.string();
  }
  if (reader.bool()) {
    validation.errorStyle = decodeValidationErrorStyle(reader.string());
  }
  if (reader.bool()) {
    validation.errorTitle = reader.string();
  }
  if (reader.bool()) {
    validation.errorMessage = reader.string();
  }
  return validation;
}

function encodeConditionalFormatRule(
  writer: BinaryWriter,
  rule: WorkbookConditionalFormatRuleSnapshot,
): void {
  switch (rule.kind) {
    case "cellIs":
      writer.u8(0);
      writer.string(rule.operator);
      writer.u32(rule.values.length);
      rule.values.forEach((value) => encodeLiteral(writer, value));
      return;
    case "textContains":
      writer.u8(1);
      writer.string(rule.text);
      writer.bool(rule.caseSensitive !== undefined);
      if (rule.caseSensitive !== undefined) {
        writer.bool(rule.caseSensitive);
      }
      return;
    case "formula":
      writer.u8(2);
      writer.string(rule.formula);
      return;
    case "blanks":
      writer.u8(3);
      return;
    case "notBlanks":
      writer.u8(4);
      return;
    default:
      assertNever(rule);
  }
}

function decodeConditionalFormatRule(reader: BinaryReader): WorkbookConditionalFormatRuleSnapshot {
  switch (reader.u8()) {
    case 0: {
      const operator = decodeValidationComparisonOperator(reader.string());
      const count = reader.u32();
      const values: LiteralInput[] = [];
      for (let index = 0; index < count; index += 1) {
        values.push(decodeLiteral(reader));
      }
      return {
        kind: "cellIs",
        operator,
        values,
      };
    }
    case 1: {
      const rule: Extract<WorkbookConditionalFormatRuleSnapshot, { kind: "textContains" }> = {
        kind: "textContains",
        text: reader.string(),
      };
      if (reader.bool()) {
        rule.caseSensitive = reader.bool();
      }
      return rule;
    }
    case 2:
      return { kind: "formula", formula: reader.string() };
    case 3:
      return { kind: "blanks" };
    case 4:
      return { kind: "notBlanks" };
    default:
      throw new BinaryProtocolError("Unknown conditional format rule tag");
  }
}

function encodeConditionalFormat(
  writer: BinaryWriter,
  format: WorkbookConditionalFormatSnapshot,
): void {
  writer.string(format.id);
  encodeCellRangeRef(writer, format.range);
  encodeConditionalFormatRule(writer, format.rule);
  encodeCellStylePatch(writer, format.style);
  writer.bool(format.stopIfTrue !== undefined);
  if (format.stopIfTrue !== undefined) {
    writer.bool(format.stopIfTrue);
  }
  writer.bool(format.priority !== undefined);
  if (format.priority !== undefined) {
    writer.f64(format.priority);
  }
}

function decodeConditionalFormat(reader: BinaryReader): WorkbookConditionalFormatSnapshot {
  const format: WorkbookConditionalFormatSnapshot = {
    id: reader.string(),
    range: decodeCellRangeRef(reader),
    rule: decodeConditionalFormatRule(reader),
    style: decodeCellStylePatch(reader),
  };
  if (reader.bool()) {
    format.stopIfTrue = reader.bool();
  }
  if (reader.bool()) {
    format.priority = reader.f64();
  }
  return format;
}

function encodeSheetProtection(
  writer: BinaryWriter,
  protection: WorkbookSheetProtectionSnapshot,
): void {
  writer.string(protection.sheetName);
  writer.bool(protection.hideFormulas !== undefined);
  if (protection.hideFormulas !== undefined) {
    writer.bool(protection.hideFormulas);
  }
}

function decodeSheetProtection(reader: BinaryReader): WorkbookSheetProtectionSnapshot {
  const protection: WorkbookSheetProtectionSnapshot = {
    sheetName: reader.string(),
  };
  if (reader.bool()) {
    protection.hideFormulas = reader.bool();
  }
  return protection;
}

function encodeRangeProtection(
  writer: BinaryWriter,
  protection: WorkbookRangeProtectionSnapshot,
): void {
  writer.string(protection.id);
  encodeCellRangeRef(writer, protection.range);
  writer.bool(protection.hideFormulas !== undefined);
  if (protection.hideFormulas !== undefined) {
    writer.bool(protection.hideFormulas);
  }
}

function decodeRangeProtection(reader: BinaryReader): WorkbookRangeProtectionSnapshot {
  const protection: WorkbookRangeProtectionSnapshot = {
    id: reader.string(),
    range: decodeCellRangeRef(reader),
  };
  if (reader.bool()) {
    protection.hideFormulas = reader.bool();
  }
  return protection;
}

function encodeCommentEntry(
  writer: BinaryWriter,
  entry: WorkbookCommentThreadSnapshot["comments"][number],
): void {
  writer.string(entry.id);
  writer.string(entry.body);
  writer.bool(entry.authorUserId !== undefined);
  if (entry.authorUserId !== undefined) {
    writer.string(entry.authorUserId);
  }
  writer.bool(entry.authorDisplayName !== undefined);
  if (entry.authorDisplayName !== undefined) {
    writer.string(entry.authorDisplayName);
  }
  writer.bool(entry.createdAtUnixMs !== undefined);
  if (entry.createdAtUnixMs !== undefined) {
    writer.u32(entry.createdAtUnixMs);
  }
}

function decodeCommentEntry(
  reader: BinaryReader,
): WorkbookCommentThreadSnapshot["comments"][number] {
  const entry: WorkbookCommentThreadSnapshot["comments"][number] = {
    id: reader.string(),
    body: reader.string(),
  };
  if (reader.bool()) {
    entry.authorUserId = reader.string();
  }
  if (reader.bool()) {
    entry.authorDisplayName = reader.string();
  }
  if (reader.bool()) {
    entry.createdAtUnixMs = reader.u32();
  }
  return entry;
}

function encodeCommentThread(writer: BinaryWriter, thread: WorkbookCommentThreadSnapshot): void {
  writer.string(thread.threadId);
  writer.string(thread.sheetName);
  writer.string(thread.address);
  writer.u32(thread.comments.length);
  thread.comments.forEach((entry) => encodeCommentEntry(writer, entry));
  writer.bool(thread.resolved !== undefined);
  if (thread.resolved !== undefined) {
    writer.bool(thread.resolved);
  }
  writer.bool(thread.resolvedByUserId !== undefined);
  if (thread.resolvedByUserId !== undefined) {
    writer.string(thread.resolvedByUserId);
  }
  writer.bool(thread.resolvedAtUnixMs !== undefined);
  if (thread.resolvedAtUnixMs !== undefined) {
    writer.u32(thread.resolvedAtUnixMs);
  }
}

function decodeCommentThread(reader: BinaryReader): WorkbookCommentThreadSnapshot {
  const thread: WorkbookCommentThreadSnapshot = {
    threadId: reader.string(),
    sheetName: reader.string(),
    address: reader.string(),
    comments: [],
  };
  const count = reader.u32();
  for (let index = 0; index < count; index += 1) {
    thread.comments.push(decodeCommentEntry(reader));
  }
  if (reader.bool()) {
    thread.resolved = reader.bool();
  }
  if (reader.bool()) {
    thread.resolvedByUserId = reader.string();
  }
  if (reader.bool()) {
    thread.resolvedAtUnixMs = reader.u32();
  }
  return thread;
}

function encodeNote(writer: BinaryWriter, note: WorkbookNoteSnapshot): void {
  writer.string(note.sheetName);
  writer.string(note.address);
  writer.string(note.text);
}

function decodeNote(reader: BinaryReader): WorkbookNoteSnapshot {
  return {
    sheetName: reader.string(),
    address: reader.string(),
    text: reader.string(),
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
    case "renameSheet":
      writer.string(op.oldName);
      writer.string(op.newName);
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
    case "setSheetProtection":
      encodeSheetProtection(writer, op.protection);
      return;
    case "clearSheetProtection":
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
    case "setDataValidation":
      encodeDataValidation(writer, op.validation);
      return;
    case "clearDataValidation":
      writer.string(op.sheetName);
      encodeCellRangeRef(writer, op.range);
      return;
    case "upsertConditionalFormat":
      encodeConditionalFormat(writer, op.format);
      return;
    case "deleteConditionalFormat":
      writer.string(op.id);
      writer.string(op.sheetName);
      return;
    case "upsertRangeProtection":
      encodeRangeProtection(writer, op.protection);
      return;
    case "deleteRangeProtection":
      writer.string(op.id);
      writer.string(op.sheetName);
      return;
    case "upsertCommentThread":
      encodeCommentThread(writer, op.thread);
      return;
    case "deleteCommentThread":
      writer.string(op.sheetName);
      writer.string(op.address);
      return;
    case "upsertNote":
      encodeNote(writer, op.note);
      return;
    case "deleteNote":
      writer.string(op.sheetName);
      writer.string(op.address);
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
    case "upsertCellStyle":
      encodeCellStyleRecord(writer, op.style);
      return;
    case "setStyleRange":
      encodeCellRangeRef(writer, op.range);
      writer.string(op.styleId);
      return;
    case "upsertCellNumberFormat":
      encodeCellNumberFormatRecord(writer, op.format);
      return;
    case "setFormatRange":
      encodeCellRangeRef(writer, op.range);
      writer.string(op.formatId);
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
    case 27:
      return {
        kind: "setCalculationSettings",
        settings: {
          mode: decodeCalculationMode(reader),
          compatibilityMode: decodeCompatibilityMode(reader),
        },
      };
    case 28:
      return { kind: "setVolatileContext", context: { recalcEpoch: reader.u32() } };
    case 3:
      return { kind: "upsertSheet", name: reader.string(), order: reader.u32() };
    case 4:
      return { kind: "deleteSheet", name: reader.string() };
    case 37:
      return { kind: "renameSheet", oldName: reader.string(), newName: reader.string() };
    case 29:
      return {
        kind: "insertRows",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        entries: decodeAxisEntries(reader),
      };
    case 30:
      return {
        kind: "deleteRows",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
      };
    case 31:
      return {
        kind: "moveRows",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        target: reader.u32(),
      };
    case 32:
      return {
        kind: "insertColumns",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
        entries: decodeAxisEntries(reader),
      };
    case 33:
      return {
        kind: "deleteColumns",
        sheetName: reader.string(),
        start: reader.u32(),
        count: reader.u32(),
      };
    case 34:
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
    case 46:
      return {
        kind: "setSheetProtection",
        protection: decodeSheetProtection(reader),
      };
    case 47:
      return { kind: "clearSheetProtection", sheetName: reader.string() };
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
    case 38:
      return {
        kind: "setDataValidation",
        validation: decodeDataValidation(reader),
      };
    case 39:
      return {
        kind: "clearDataValidation",
        sheetName: reader.string(),
        range: decodeCellRangeRef(reader),
      };
    case 44:
      return {
        kind: "upsertConditionalFormat",
        format: decodeConditionalFormat(reader),
      };
    case 45:
      return {
        kind: "deleteConditionalFormat",
        id: reader.string(),
        sheetName: reader.string(),
      };
    case 48:
      return {
        kind: "upsertRangeProtection",
        protection: decodeRangeProtection(reader),
      };
    case 49:
      return {
        kind: "deleteRangeProtection",
        id: reader.string(),
        sheetName: reader.string(),
      };
    case 40:
      return {
        kind: "upsertCommentThread",
        thread: decodeCommentThread(reader),
      };
    case 41:
      return {
        kind: "deleteCommentThread",
        sheetName: reader.string(),
        address: reader.string(),
      };
    case 42:
      return {
        kind: "upsertNote",
        note: decodeNote(reader),
      };
    case 43:
      return {
        kind: "deleteNote",
        sheetName: reader.string(),
        address: reader.string(),
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
      return { kind: "upsertCellStyle", style: decodeCellStyleRecord(reader) };
    case 17:
      return {
        kind: "setStyleRange",
        range: decodeCellRangeRef(reader),
        styleId: reader.string(),
      };
    case 35:
      return { kind: "upsertCellNumberFormat", format: decodeCellNumberFormatRecord(reader) };
    case 36:
      return {
        kind: "setFormatRange",
        range: decodeCellRangeRef(reader),
        formatId: reader.string(),
      };
    case 18:
      return { kind: "clearCell", sheetName: reader.string(), address: reader.string() };
    case 19:
      return {
        kind: "upsertDefinedName",
        name: reader.string(),
        value: decodeDefinedNameValue(reader),
      };
    case 20:
      return { kind: "deleteDefinedName", name: reader.string() };
    case 21:
      return { kind: "upsertTable", table: decodeTable(reader) };
    case 22:
      return { kind: "deleteTable", name: reader.string() };
    case 23:
      return {
        kind: "upsertSpillRange",
        sheetName: reader.string(),
        address: reader.string(),
        rows: reader.u32(),
        cols: reader.u32(),
      };
    case 24:
      return {
        kind: "deleteSpillRange",
        sheetName: reader.string(),
        address: reader.string(),
      };
    case 25:
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
    case 26:
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
