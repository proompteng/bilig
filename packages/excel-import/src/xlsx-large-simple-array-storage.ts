export function filledUint32Array(length: number, value: number): Uint32Array<ArrayBuffer> {
  const output = new Uint32Array(length)
  output.fill(value)
  return output
}

export function growUint8Array(source: Uint8Array<ArrayBuffer>, nextCapacity: number): Uint8Array<ArrayBuffer> {
  const output = new Uint8Array(nextCapacity)
  output.set(source)
  return output
}

export function growInt8Array(source: Int8Array<ArrayBuffer>, nextCapacity: number): Int8Array<ArrayBuffer> {
  const output = new Int8Array(nextCapacity)
  output.set(source)
  return output
}

export function growUint32Array(source: Uint32Array<ArrayBuffer>, nextCapacity: number, fillValue?: number): Uint32Array<ArrayBuffer> {
  const output = new Uint32Array(nextCapacity)
  output.set(source)
  if (fillValue !== undefined && nextCapacity > source.length) {
    output.fill(fillValue, source.length)
  }
  return output
}

export function growUint16Array(source: Uint16Array<ArrayBuffer>, nextCapacity: number): Uint16Array<ArrayBuffer> {
  const output = new Uint16Array(nextCapacity)
  output.set(source)
  return output
}

export function growInt32Array(source: Int32Array<ArrayBuffer>, nextCapacity: number): Int32Array<ArrayBuffer> {
  const output = new Int32Array(nextCapacity)
  output.set(source)
  return output
}

export function growInt16Array(source: Int16Array<ArrayBuffer>, nextCapacity: number): Int16Array<ArrayBuffer> {
  const output = new Int16Array(nextCapacity)
  output.set(source)
  return output
}

export function growFloat64Array(source: Float64Array<ArrayBuffer>, nextCapacity: number): Float64Array<ArrayBuffer> {
  const output = new Float64Array(nextCapacity)
  output.set(source)
  if (nextCapacity > source.length) {
    output.fill(Number.NaN, source.length)
  }
  return output
}
