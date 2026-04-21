import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT,
  GRID_SCENE_PACKET_V2_VERSION,
  type GridScenePacketV2,
} from './scene-packet-v2.js'

export type GridScenePacketValidationResult =
  | { readonly ok: true }
  | {
      readonly ok: false
      readonly reason: string
    }

export function validateGridScenePacketV2(packet: GridScenePacketV2): GridScenePacketValidationResult {
  if (packet.magic !== GRID_SCENE_PACKET_V2_MAGIC) {
    return invalid('bad magic')
  }
  if (packet.version !== GRID_SCENE_PACKET_V2_VERSION) {
    return invalid('bad version')
  }
  if (!Number.isInteger(packet.generation) || packet.generation < 0) {
    return invalid('bad generation')
  }
  if (packet.sheetName.length === 0) {
    return invalid('missing sheet name')
  }
  if (
    packet.viewport.rowStart < 0 ||
    packet.viewport.colStart < 0 ||
    packet.viewport.rowEnd >= MAX_ROWS ||
    packet.viewport.colEnd >= MAX_COLS
  ) {
    return invalid('viewport out of bounds')
  }
  if (packet.viewport.rowEnd < packet.viewport.rowStart || packet.viewport.colEnd < packet.viewport.colStart) {
    return invalid('empty viewport')
  }
  if (!isFiniteNonNegative(packet.surfaceSize.width) || !isFiniteNonNegative(packet.surfaceSize.height)) {
    return invalid('bad surface size')
  }
  if (!Number.isInteger(packet.rectCount) || packet.rectCount < 0) {
    return invalid('bad rect count')
  }
  if (!Number.isInteger(packet.fillRectCount) || packet.fillRectCount < 0) {
    return invalid('bad fill rect count')
  }
  if (!Number.isInteger(packet.borderRectCount) || packet.borderRectCount < 0) {
    return invalid('bad border rect count')
  }
  if (packet.fillRectCount + packet.borderRectCount !== packet.rectCount) {
    return invalid('rect count mismatch')
  }
  if (!Number.isInteger(packet.textCount) || packet.textCount < 0) {
    return invalid('bad text count')
  }
  if (packet.rects.length < Math.max(1, packet.rectCount) * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT) {
    return invalid('rect buffer too small')
  }
  if (packet.rectInstances.length < Math.max(1, packet.rectCount) * GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT) {
    return invalid('rect instance buffer too small')
  }
  if (packet.textMetrics.length < Math.max(1, packet.textCount) * GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT) {
    return invalid('text buffer too small')
  }
  for (let index = 0; index < packet.rectCount; index += 1) {
    const offset = index * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT
    if (!isFiniteNumber(readFloat(packet.rects, offset + 0)) || !isFiniteNumber(readFloat(packet.rects, offset + 1))) {
      return invalid('bad rect position')
    }
    if (!isFiniteNonNegative(readFloat(packet.rects, offset + 2)) || !isFiniteNonNegative(readFloat(packet.rects, offset + 3))) {
      return invalid('bad rect size')
    }
    for (let colorOffset = 4; colorOffset < 8; colorOffset += 1) {
      if (!isFiniteNumber(readFloat(packet.rects, offset + colorOffset))) {
        return invalid('bad rect color')
      }
    }
    const instanceOffset = index * GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT
    if (
      !isFiniteNumber(readFloat(packet.rectInstances, instanceOffset + 0)) ||
      !isFiniteNumber(readFloat(packet.rectInstances, instanceOffset + 1))
    ) {
      return invalid('bad rect instance position')
    }
    if (
      !isFiniteNonNegative(readFloat(packet.rectInstances, instanceOffset + 2)) ||
      !isFiniteNonNegative(readFloat(packet.rectInstances, instanceOffset + 3))
    ) {
      return invalid('bad rect instance size')
    }
    for (let colorOffset = 4; colorOffset < 12; colorOffset += 1) {
      if (!isFiniteNumber(readFloat(packet.rectInstances, instanceOffset + colorOffset))) {
        return invalid('bad rect instance color')
      }
    }
    if (
      !isFiniteNonNegative(readFloat(packet.rectInstances, instanceOffset + 16)) ||
      !isFiniteNonNegative(readFloat(packet.rectInstances, instanceOffset + 17)) ||
      !isFiniteNonNegative(readFloat(packet.rectInstances, instanceOffset + 18)) ||
      !isFiniteNonNegative(readFloat(packet.rectInstances, instanceOffset + 19))
    ) {
      return invalid('bad rect instance clip')
    }
  }
  for (let index = 0; index < packet.textCount; index += 1) {
    const offset = index * GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT
    if (!isFiniteNumber(readFloat(packet.textMetrics, offset + 0)) || !isFiniteNumber(readFloat(packet.textMetrics, offset + 1))) {
      return invalid('bad text position')
    }
    if (
      !isFiniteNonNegative(readFloat(packet.textMetrics, offset + 2)) ||
      !isFiniteNonNegative(readFloat(packet.textMetrics, offset + 3))
    ) {
      return invalid('bad text size')
    }
    for (let clipOffset = 4; clipOffset < 8; clipOffset += 1) {
      if (!isFiniteNonNegative(readFloat(packet.textMetrics, offset + clipOffset))) {
        return invalid('bad text clip')
      }
    }
  }
  return { ok: true }
}

function invalid(reason: string): GridScenePacketValidationResult {
  return { ok: false, reason }
}

function readFloat(buffer: Float32Array, index: number): number {
  return buffer[index] ?? Number.NaN
}

function isFiniteNumber(value: number): boolean {
  return Number.isFinite(value)
}

function isFiniteNonNegative(value: number): boolean {
  return Number.isFinite(value) && value >= 0
}
