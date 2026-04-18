import type { WorkbookSnapshot } from '@bilig/protocol'
import type { RuntimeImage } from './runtime-image.js'

const RUNTIME_IMAGE = Symbol.for('bilig.runtimeImage')
const RUNTIME_SNAPSHOT = Symbol.for('bilig.runtimeSnapshot')

type RuntimeImageCarrier = {
  [RUNTIME_IMAGE]?: RuntimeImage
}

type RuntimeSnapshotCarrier = {
  [RUNTIME_SNAPSHOT]?: WorkbookSnapshot
}

export function attachRuntimeImage<T extends object>(carrier: T, runtimeImage: RuntimeImage): T {
  Object.defineProperty(carrier, RUNTIME_IMAGE, {
    value: runtimeImage,
    configurable: true,
    enumerable: false,
    writable: true,
  })
  return carrier
}

export function readRuntimeImage(carrier: unknown): RuntimeImage | undefined {
  if (!carrier || typeof carrier !== 'object') {
    return undefined
  }
  return (carrier as RuntimeImageCarrier)[RUNTIME_IMAGE]
}

export function attachRuntimeSnapshot<T extends object>(carrier: T, snapshot: WorkbookSnapshot): T {
  Object.defineProperty(carrier, RUNTIME_SNAPSHOT, {
    value: snapshot,
    configurable: true,
    enumerable: false,
    writable: true,
  })
  return carrier
}

export function readRuntimeSnapshot(carrier: unknown): WorkbookSnapshot | undefined {
  if (!carrier || typeof carrier !== 'object') {
    return undefined
  }
  return (carrier as RuntimeSnapshotCarrier)[RUNTIME_SNAPSHOT]
}
