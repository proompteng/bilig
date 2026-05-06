import type { WorkbookMacroPayloadSnapshot } from '@bilig/protocol'

const binaryChunkSize = 0x8000

export interface PreservedVbaProjectCodeNames {
  readonly workbookCodeName?: string
  readonly sheetCodeNames?: readonly {
    readonly sheetName: string
    readonly codeName: string
  }[]
}

function encodeBinaryString(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += binaryChunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + binaryChunkSize))
  }
  return binary
}

function decodeBinaryString(binary: string): Uint8Array {
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

function encodeBase64(bytes: Uint8Array): string {
  const btoa = globalThis.btoa
  if (typeof btoa === 'function') {
    return btoa(encodeBinaryString(bytes))
  }
  return Buffer.from(bytes).toString('base64')
}

function decodeBase64(dataBase64: string): Uint8Array {
  const atob = globalThis.atob
  if (typeof atob === 'function') {
    return decodeBinaryString(atob(dataBase64))
  }
  return new Uint8Array(Buffer.from(dataBase64, 'base64'))
}

function normalizeCodeName(value: string | undefined): string | undefined {
  const trimmed = value?.trim()
  return trimmed && trimmed.length > 0 ? trimmed : undefined
}

function normalizeSheetCodeNames(
  sheetCodeNames: PreservedVbaProjectCodeNames['sheetCodeNames'] | undefined,
): WorkbookMacroPayloadSnapshot['sheetCodeNames'] | undefined {
  const normalized =
    sheetCodeNames
      ?.map((entry) => ({
        sheetName: entry.sheetName.trim(),
        codeName: entry.codeName.trim(),
      }))
      .filter((entry) => entry.sheetName.length > 0 && entry.codeName.length > 0) ?? []
  return normalized.length > 0 ? normalized : undefined
}

export function createPreservedVbaProjectPayload(
  bytes: Uint8Array,
  codeNames: PreservedVbaProjectCodeNames = {},
): WorkbookMacroPayloadSnapshot {
  const workbookCodeName = normalizeCodeName(codeNames.workbookCodeName)
  const sheetCodeNames = normalizeSheetCodeNames(codeNames.sheetCodeNames)
  return {
    kind: 'vbaProject',
    storage: 'base64',
    dataBase64: encodeBase64(bytes),
    byteLength: bytes.byteLength,
    preservedWithoutExecution: true,
    ...(workbookCodeName ? { workbookCodeName } : {}),
    ...(sheetCodeNames ? { sheetCodeNames } : {}),
  }
}

export function decodePreservedVbaProjectPayload(payload: WorkbookMacroPayloadSnapshot | undefined): Uint8Array | undefined {
  if (!payload || payload.kind !== 'vbaProject' || payload.storage !== 'base64' || !payload.preservedWithoutExecution) {
    return undefined
  }
  const bytes = decodeBase64(payload.dataBase64)
  return bytes.byteLength === payload.byteLength ? bytes : undefined
}
