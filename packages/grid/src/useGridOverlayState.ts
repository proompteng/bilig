import type { Rectangle } from './gridTypes.js'

export function sameBounds(left: Rectangle | undefined, right: Rectangle | undefined): boolean {
  return left?.x === right?.x && left?.y === right?.y && left?.width === right?.width && left?.height === right?.height
}
