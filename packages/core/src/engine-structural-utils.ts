import type { StructuralAxisTransform } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'

export function structuralTransformForOp(
  op: Extract<
    EngineOp,
    {
      kind: 'insertRows' | 'deleteRows' | 'moveRows' | 'insertColumns' | 'deleteColumns' | 'moveColumns'
    }
  >,
): StructuralAxisTransform {
  switch (op.kind) {
    case 'insertRows':
      return { kind: 'insert', axis: 'row', start: op.start, count: op.count }
    case 'deleteRows':
      return { kind: 'delete', axis: 'row', start: op.start, count: op.count }
    case 'moveRows':
      return { kind: 'move', axis: 'row', start: op.start, count: op.count, target: op.target }
    case 'insertColumns':
      return { kind: 'insert', axis: 'column', start: op.start, count: op.count }
    case 'deleteColumns':
      return { kind: 'delete', axis: 'column', start: op.start, count: op.count }
    case 'moveColumns':
      return { kind: 'move', axis: 'column', start: op.start, count: op.count, target: op.target }
    default:
      return assertNever(op)
  }
}

export function mapStructuralAxisIndex(index: number, transform: StructuralAxisTransform): number | undefined {
  switch (transform.kind) {
    case 'insert':
      return index >= transform.start ? index + transform.count : index
    case 'delete':
      if (index < transform.start) {
        return index
      }
      if (index >= transform.start + transform.count) {
        return index - transform.count
      }
      return undefined
    case 'move':
      if (transform.target < transform.start) {
        if (index >= transform.target && index < transform.start) {
          return index + transform.count
        }
      } else if (transform.target > transform.start) {
        if (index >= transform.start + transform.count && index < transform.target + transform.count) {
          return index - transform.count
        }
      }
      if (index >= transform.start && index < transform.start + transform.count) {
        return transform.target + (index - transform.start)
      }
      return index
    default:
      return assertNever(transform)
  }
}

export function mapStructuralBoundary(boundary: number, transform: StructuralAxisTransform): number {
  if (boundary <= 0) {
    return 0
  }
  const mapped = mapStructuralAxisIndex(boundary - 1, transform)
  return mapped === undefined ? 0 : mapped + 1
}

function assertNever(value: never): never {
  throw new Error(`Unhandled structural transform case: ${JSON.stringify(value)}`)
}
