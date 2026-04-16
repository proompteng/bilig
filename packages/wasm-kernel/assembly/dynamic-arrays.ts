const DEFAULT_DYNAMIC_ARRAY_SHAPE_CAPACITY = 16

let dynamicArrayShapeRows = new Uint32Array(DEFAULT_DYNAMIC_ARRAY_SHAPE_CAPACITY)
let dynamicArrayShapeCols = new Uint32Array(DEFAULT_DYNAMIC_ARRAY_SHAPE_CAPACITY)
let dynamicArrayShapeValid = new Uint8Array(DEFAULT_DYNAMIC_ARRAY_SHAPE_CAPACITY)

function ensureDynamicArrayShapeCapacity(required: i32): void {
  if (required <= dynamicArrayShapeRows.length) {
    return
  }
  let nextCapacity = dynamicArrayShapeRows.length
  while (nextCapacity < required) {
    nextCapacity *= 2
  }
  const nextRows = new Uint32Array(nextCapacity)
  const nextCols = new Uint32Array(nextCapacity)
  const nextValid = new Uint8Array(nextCapacity)
  for (let index = 0; index < dynamicArrayShapeRows.length; index++) {
    nextRows[index] = dynamicArrayShapeRows[index]
    nextCols[index] = dynamicArrayShapeCols[index]
    nextValid[index] = dynamicArrayShapeValid[index]
  }
  dynamicArrayShapeRows = nextRows
  dynamicArrayShapeCols = nextCols
  dynamicArrayShapeValid = nextValid
}

export function registerTrackedArrayShape(arrayIndex: u32, rows: i32, cols: i32): void {
  const index = <i32>arrayIndex
  ensureDynamicArrayShapeCapacity(index + 1)
  dynamicArrayShapeRows[index] = <u32>rows
  dynamicArrayShapeCols[index] = <u32>cols
  dynamicArrayShapeValid[index] = 1
}

export function getTrackedArrayRows(arrayIndex: u32): i32 {
  const index = <i32>arrayIndex
  if (index < 0 || index >= dynamicArrayShapeValid.length || dynamicArrayShapeValid[index] == 0) {
    return i32.MIN_VALUE
  }
  return <i32>dynamicArrayShapeRows[index]
}

export function getTrackedArrayCols(arrayIndex: u32): i32 {
  const index = <i32>arrayIndex
  if (index < 0 || index >= dynamicArrayShapeValid.length || dynamicArrayShapeValid[index] == 0) {
    return i32.MIN_VALUE
  }
  return <i32>dynamicArrayShapeCols[index]
}
