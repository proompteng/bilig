import {
  WORKBOOK_SNAPSHOT_CONTENT_TYPE,
  createSnapshotChunkFrames,
  type ProtocolFrame,
  type SnapshotChunkFrame,
} from '@bilig/binary-protocol'
import {
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  type WorkbookImportContentType,
} from '@bilig/agent-api'
import { isWorkbookSnapshot, type WorkbookSnapshot } from '@bilig/protocol'

const snapshotEncoder = new TextEncoder()
const snapshotDecoder = new TextDecoder()
const encodedSnapshotCache = new WeakMap<WorkbookSnapshot, Uint8Array>()
let snapshotPublicationSequence = 0

export const SNAPSHOT_ASSEMBLY_MAX_AGE_MS = 5 * 60_000
export const SNAPSHOT_ASSEMBLY_MAX_CHUNKS = 4_096
const SNAPSHOT_ASSEMBLY_MAX_BYTES = 128 * 1024 * 1024

export interface BrowserSubscriber {
  id: string
  send(frame: ProtocolFrame): void
}

export type BrowserSubscriberRegistry = Map<string, Map<string, BrowserSubscriber>>

export interface SnapshotPublication {
  snapshotId: string
  contentType: typeof WORKBOOK_SNAPSHOT_CONTENT_TYPE
  bytes: Uint8Array
  frames: ReturnType<typeof createSnapshotChunkFrames>
}

interface SnapshotAssembly {
  documentId: string
  snapshotId: string
  cursor: number
  contentType: string
  chunkCount: number
  chunks: Array<Uint8Array | undefined>
  totalByteLength: number
  updatedAtUnixMs: number
}

export interface CompletedSnapshotAssembly {
  documentId: string
  snapshotId: string
  cursor: number
  contentType: string
  bytes: Uint8Array
}

export type SnapshotAssemblyRegistry = Map<string, SnapshotAssembly>

export function attachBrowserSubscriber(
  registry: BrowserSubscriberRegistry,
  documentId: string,
  subscriberId: string,
  send: (frame: ProtocolFrame) => void,
): () => void {
  const subscribers = registry.get(documentId) ?? new Map<string, BrowserSubscriber>()
  subscribers.set(subscriberId, { id: subscriberId, send })
  registry.set(documentId, subscribers)
  return () => {
    const next = registry.get(documentId)
    next?.delete(subscriberId)
    if (next && next.size === 0) {
      registry.delete(documentId)
    }
  }
}

export function broadcastToBrowsers(registry: BrowserSubscriberRegistry, documentId: string, frame: ProtocolFrame): void {
  registry.get(documentId)?.forEach((subscriber) => subscriber.send(frame))
}

export function listBrowserSubscriberIds(registry: BrowserSubscriberRegistry, documentId: string): string[] {
  return [...(registry.get(documentId)?.keys() ?? [])]
}

export function normalizeBaseUrl(value: string): string {
  return value.endsWith('/') ? value.slice(0, -1) : value
}

export function buildBrowserUrl(browserAppBaseUrl: string | undefined, serverUrl: string, documentId: string): string | undefined {
  if (!browserAppBaseUrl) {
    return undefined
  }
  const url = new URL(normalizeBaseUrl(browserAppBaseUrl))
  url.searchParams.set('document', documentId)
  url.searchParams.set('server', serverUrl)
  return url.toString()
}

export function decodeWorkbookBase64(bytesBase64: string): Uint8Array {
  const normalized = bytesBase64.trim()
  if (normalized.length === 0 || normalized.length % 4 !== 0) {
    throw new Error('Workbook upload bytesBase64 must be a non-empty base64 string')
  }
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(normalized)) {
    throw new Error('Workbook upload bytesBase64 contains invalid base64 characters')
  }
  return new Uint8Array(Buffer.from(normalized, 'base64'))
}

export function createImportedDocumentId(contentType?: WorkbookImportContentType): string {
  const random =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function' ? crypto.randomUUID() : Math.random().toString(36).slice(2)
  const prefix =
    contentType === CSV_CONTENT_TYPE
      ? 'csv'
      : contentType === XLSB_CONTENT_TYPE
        ? 'xlsb'
        : contentType === XLSM_CONTENT_TYPE
          ? 'xlsm'
          : contentType === LEGACY_XLS_CONTENT_TYPE
            ? 'xls'
            : 'xlsx'
  return `${prefix}:${random}`
}

function createSnapshotPublicationId(documentId: string): string {
  const createdAtUnixMs = Date.now()
  const entropy =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `${createdAtUnixMs}:${snapshotPublicationSequence++}`
  return `${documentId}:snapshot:${createdAtUnixMs}:${entropy}`
}

export function createSnapshotPublication(documentId: string, cursor: number, snapshot: WorkbookSnapshot): SnapshotPublication {
  const snapshotId = createSnapshotPublicationId(documentId)
  const bytes = encodeWorkbookSnapshot(snapshot)
  return createSnapshotPublicationFromBytes({
    documentId,
    snapshotId,
    cursor,
    contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
    bytes,
  })
}

export function createSnapshotPublicationFromBytes(snapshot: CompletedSnapshotAssembly): SnapshotPublication {
  return {
    snapshotId: snapshot.snapshotId,
    contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
    bytes: snapshot.bytes,
    frames: createSnapshotChunkFrames({
      documentId: snapshot.documentId,
      snapshotId: snapshot.snapshotId,
      cursor: snapshot.cursor,
      contentType: snapshot.contentType,
      bytes: snapshot.bytes,
    }),
  }
}

export function encodeWorkbookSnapshot(snapshot: WorkbookSnapshot): Uint8Array {
  const cached = encodedSnapshotCache.get(snapshot)
  if (cached) {
    return cached
  }
  const bytes = snapshotEncoder.encode(JSON.stringify(snapshot))
  encodedSnapshotCache.set(snapshot, bytes)
  return bytes
}

function pruneExpiredSnapshotAssemblies(registry: SnapshotAssemblyRegistry, nowUnixMs: number, maxAgeMs: number): void {
  registry.forEach((assembly, snapshotId) => {
    if (nowUnixMs - assembly.updatedAtUnixMs > maxAgeMs) {
      registry.delete(snapshotId)
    }
  })
}

export interface SnapshotAssemblyOptions {
  readonly nowUnixMs?: number
  readonly maxAgeMs?: number
  readonly maxChunks?: number
  readonly maxBytes?: number
}

function isValidSnapshotChunkFrame(frame: SnapshotChunkFrame, maxChunks: number, maxBytes: number): boolean {
  return (
    frame.documentId.length > 0 &&
    frame.snapshotId.length > 0 &&
    frame.contentType.length > 0 &&
    Number.isSafeInteger(frame.cursor) &&
    frame.cursor >= 0 &&
    Number.isSafeInteger(frame.chunkCount) &&
    frame.chunkCount > 0 &&
    frame.chunkCount <= maxChunks &&
    Number.isSafeInteger(frame.chunkIndex) &&
    frame.chunkIndex >= 0 &&
    frame.chunkIndex < frame.chunkCount &&
    frame.bytes instanceof Uint8Array &&
    frame.bytes.byteLength <= maxBytes
  )
}

function snapshotChunkMatchesAssembly(assembly: SnapshotAssembly, frame: SnapshotChunkFrame): boolean {
  return (
    assembly.documentId === frame.documentId &&
    assembly.snapshotId === frame.snapshotId &&
    assembly.cursor === frame.cursor &&
    assembly.contentType === frame.contentType &&
    assembly.chunkCount === frame.chunkCount
  )
}

function byteArraysEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) {
    return false
  }
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

export function acceptSnapshotChunk(
  registry: SnapshotAssemblyRegistry,
  frame: SnapshotChunkFrame,
  options: SnapshotAssemblyOptions = {},
): CompletedSnapshotAssembly | null {
  const nowUnixMs = options.nowUnixMs ?? Date.now()
  const maxAgeMs = options.maxAgeMs ?? SNAPSHOT_ASSEMBLY_MAX_AGE_MS
  const maxChunks = options.maxChunks ?? SNAPSHOT_ASSEMBLY_MAX_CHUNKS
  const maxBytes = options.maxBytes ?? SNAPSHOT_ASSEMBLY_MAX_BYTES
  pruneExpiredSnapshotAssemblies(registry, nowUnixMs, maxAgeMs)

  if (!isValidSnapshotChunkFrame(frame, maxChunks, maxBytes)) {
    registry.delete(frame.snapshotId)
    return null
  }

  const existingAssembly = registry.get(frame.snapshotId)
  if (existingAssembly && !snapshotChunkMatchesAssembly(existingAssembly, frame)) {
    registry.delete(frame.snapshotId)
    return null
  }

  const assembly = existingAssembly ?? {
    documentId: frame.documentId,
    snapshotId: frame.snapshotId,
    cursor: frame.cursor,
    contentType: frame.contentType,
    chunkCount: frame.chunkCount,
    chunks: Array.from<Uint8Array | undefined>({ length: frame.chunkCount }),
    totalByteLength: 0,
    updatedAtUnixMs: nowUnixMs,
  }

  const existingChunk = assembly.chunks[frame.chunkIndex]
  if (existingChunk) {
    if (!byteArraysEqual(existingChunk, frame.bytes)) {
      registry.delete(frame.snapshotId)
      return null
    }
  } else {
    if (assembly.totalByteLength + frame.bytes.byteLength > maxBytes) {
      registry.delete(frame.snapshotId)
      return null
    }
    assembly.chunks[frame.chunkIndex] = frame.bytes
    assembly.totalByteLength += frame.bytes.byteLength
  }
  assembly.updatedAtUnixMs = nowUnixMs
  registry.set(frame.snapshotId, assembly)

  if (!assembly.chunks.every((chunk): chunk is Uint8Array => chunk !== undefined)) {
    return null
  }

  const totalLength = assembly.chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0)
  const bytes = new Uint8Array(totalLength)
  let offset = 0
  assembly.chunks.forEach((chunk) => {
    bytes.set(chunk, offset)
    offset += chunk.byteLength
  })
  registry.delete(frame.snapshotId)

  return {
    documentId: assembly.documentId,
    snapshotId: assembly.snapshotId,
    cursor: assembly.cursor,
    contentType: assembly.contentType,
    bytes,
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function hasWorkbookSnapshotCellShape(value: unknown): value is WorkbookSnapshot {
  if (!isWorkbookSnapshot(value)) {
    return false
  }
  const sheets = value['sheets']
  return sheets.every((sheet) => {
    if (!isRecord(sheet) || typeof sheet['name'] !== 'string' || typeof sheet['order'] !== 'number') {
      return false
    }
    const cells = sheet['cells']
    if (!Array.isArray(cells)) {
      return false
    }
    return cells.every((cell) => isRecord(cell) && typeof cell['address'] === 'string')
  })
}

export function decodeWorkbookSnapshotBytes(snapshot: CompletedSnapshotAssembly): WorkbookSnapshot {
  if (snapshot.contentType !== WORKBOOK_SNAPSHOT_CONTENT_TYPE) {
    throw new Error(`Unsupported snapshot content type: ${snapshot.contentType}`)
  }
  const decoded: unknown = JSON.parse(snapshotDecoder.decode(snapshot.bytes))
  if (!hasWorkbookSnapshotCellShape(decoded)) {
    throw new Error('Workbook snapshot payload does not match the expected schema')
  }
  return decoded
}
