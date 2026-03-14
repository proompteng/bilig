export interface EdgeSlice {
  ptr: number;
  len: number;
  cap: number;
}

const EMPTY_SLICE: EdgeSlice = { ptr: -1, len: 0, cap: 0 };

export class EdgeArena {
  private buffer = new Uint32Array(64);
  private freeList: EdgeSlice[] = [];
  private nextPtr = 0;

  reset(): void {
    this.buffer.fill(0);
    this.freeList = [];
    this.nextPtr = 0;
  }

  empty(): EdgeSlice {
    return EMPTY_SLICE;
  }

  alloc(size: number): EdgeSlice {
    if (size <= 0) {
      return EMPTY_SLICE;
    }

    const freeIndex = this.freeList.findIndex((slice) => slice.cap >= size);
    if (freeIndex !== -1) {
      const [slice] = this.freeList.splice(freeIndex, 1);
      return { ptr: slice!.ptr, len: 0, cap: slice!.cap };
    }

    const ptr = this.nextPtr;
    this.ensureCapacity(ptr + size);
    this.nextPtr += size;
    return { ptr, len: 0, cap: size };
  }

  replace(slice: EdgeSlice, nextValues: Uint32Array | readonly number[]): EdgeSlice {
    const values = nextValues instanceof Uint32Array ? nextValues : Uint32Array.from(nextValues);
    if (values.length === 0) {
      this.free(slice);
      return EMPTY_SLICE;
    }

    let target = slice;
    if (target.cap < values.length || target.ptr < 0) {
      this.free(slice);
      target = this.alloc(values.length);
    }

    this.buffer.set(values, target.ptr);
    return {
      ptr: target.ptr,
      len: values.length,
      cap: target.cap
    };
  }

  read(slice: EdgeSlice): Uint32Array {
    if (slice.ptr < 0 || slice.len <= 0) {
      return new Uint32Array();
    }
    return this.buffer.slice(slice.ptr, slice.ptr + slice.len);
  }

  appendUnique(slice: EdgeSlice, value: number): EdgeSlice {
    const values = this.read(slice);
    for (let index = 0; index < values.length; index += 1) {
      if (values[index] === value) {
        return slice;
      }
    }
    const next = new Uint32Array(values.length + 1);
    next.set(values);
    next[values.length] = value;
    return this.replace(slice, next);
  }

  removeValue(slice: EdgeSlice, value: number): EdgeSlice {
    const values = this.read(slice);
    if (values.length === 0) {
      return slice;
    }

    let found = false;
    const next = new Uint32Array(values.length);
    let cursor = 0;
    for (let index = 0; index < values.length; index += 1) {
      const current = values[index]!;
      if (current === value) {
        found = true;
        continue;
      }
      next[cursor] = current;
      cursor += 1;
    }

    if (!found) {
      return slice;
    }
    return this.replace(slice, next.subarray(0, cursor));
  }

  free(slice: EdgeSlice): void {
    if (slice.ptr < 0 || slice.cap <= 0) {
      return;
    }
    this.freeList.push({
      ptr: slice.ptr,
      len: 0,
      cap: slice.cap
    });
  }

  private ensureCapacity(nextSize: number): void {
    if (nextSize <= this.buffer.length) {
      return;
    }
    let capacity = this.buffer.length;
    while (capacity < nextSize) {
      capacity *= 2;
    }
    const next = new Uint32Array(capacity);
    next.set(this.buffer);
    this.buffer = next;
  }
}
