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

export function mapStructuralAxisInterval(
  start: number,
  end: number,
  transform: StructuralAxisTransform,
): { start: number; end: number } | undefined {
  switch (transform.kind) {
    case 'insert':
      if (transform.start <= start) {
        return { start: start + transform.count, end: end + transform.count }
      }
      if (transform.start <= end) {
        return { start, end: end + transform.count }
      }
      return { start, end }
    case 'delete': {
      const deleteEnd = transform.start + transform.count - 1
      if (deleteEnd < start) {
        return { start: start - transform.count, end: end - transform.count }
      }
      if (transform.start > end) {
        return { start, end }
      }
      const survivingStart = start < transform.start ? start : deleteEnd + 1
      const survivingEnd = end > deleteEnd ? end : transform.start - 1
      if (survivingStart > survivingEnd) {
        return undefined
      }
      const nextStart = mapStructuralAxisIndex(survivingStart, transform)
      const nextEnd = mapStructuralAxisIndex(survivingEnd, transform)
      return nextStart === undefined || nextEnd === undefined ? undefined : { start: nextStart, end: nextEnd }
    }
    case 'move': {
      const segments =
        transform.target < transform.start
          ? [
              { start: 0, end: transform.target - 1, delta: 0 },
              { start: transform.target, end: transform.start - 1, delta: transform.count },
              {
                start: transform.start,
                end: transform.start + transform.count - 1,
                delta: transform.target - transform.start,
              },
              { start: transform.start + transform.count, end: Number.MAX_SAFE_INTEGER, delta: 0 },
            ]
          : [
              { start: 0, end: transform.start - 1, delta: 0 },
              {
                start: transform.start,
                end: transform.start + transform.count - 1,
                delta: transform.target - transform.start,
              },
              {
                start: transform.start + transform.count,
                end: transform.target + transform.count - 1,
                delta: -transform.count,
              },
              { start: transform.target + transform.count, end: Number.MAX_SAFE_INTEGER, delta: 0 },
            ]
      let nextStart: number | undefined
      let nextEnd: number | undefined
      segments.forEach((segment) => {
        const overlapStart = Math.max(start, segment.start)
        const overlapEnd = Math.min(end, segment.end)
        if (overlapStart > overlapEnd) {
          return
        }
        const mappedStart = overlapStart + segment.delta
        const mappedEnd = overlapEnd + segment.delta
        nextStart = nextStart === undefined ? mappedStart : Math.min(nextStart, mappedStart)
        nextEnd = nextEnd === undefined ? mappedEnd : Math.max(nextEnd, mappedEnd)
      })
      return nextStart === undefined || nextEnd === undefined ? undefined : { start: nextStart, end: nextEnd }
    }
    default:
      return assertNever(transform)
  }
}

export function inverseMapStructuralAxisIndex(index: number, transform: StructuralAxisTransform): number | undefined {
  switch (transform.kind) {
    case 'insert':
      if (index < transform.start) {
        return index
      }
      if (index >= transform.start + transform.count) {
        return index - transform.count
      }
      return undefined
    case 'delete':
      return index >= transform.start ? index + transform.count : index
    case 'move':
      if (transform.target < transform.start) {
        if (index >= transform.target && index < transform.target + transform.count) {
          return transform.start + (index - transform.target)
        }
        if (index >= transform.target + transform.count && index < transform.start + transform.count) {
          return index - transform.count
        }
        return index
      }
      if (transform.target > transform.start) {
        if (index >= transform.target && index < transform.target + transform.count) {
          return transform.start + (index - transform.target)
        }
        if (index >= transform.start && index < transform.target) {
          return index + transform.count
        }
        return index
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
