export function growUint32(buffer: Uint32Array, required: number): Uint32Array {
  let capacity = buffer.length
  while (capacity < required) {
    capacity *= 2
  }
  const next = new Uint32Array(capacity)
  next.set(buffer)
  return next
}

export function appendPackedCellIndex(indices: Uint32Array, cellIndex: number): Uint32Array {
  for (let index = 0; index < indices.length; index += 1) {
    if (indices[index] === cellIndex) {
      return indices
    }
  }
  const next = new Uint32Array(indices.length + 1)
  next.set(indices)
  next[indices.length] = cellIndex
  return next
}
