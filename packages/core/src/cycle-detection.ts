type U32 = Uint32Array<ArrayBufferLike>;
type I32 = Int32Array<ArrayBufferLike>;

export interface PackedCycleDetectionResult {
  cycleMembers: U32;
  cycleMemberCount: number;
  cycleGroups: I32;
}

export class CycleDetector {
  private visitEpoch = 1;
  private lastFormulaCount = 0;
  private visitMarks: U32 = new Uint32Array(64);
  private onStackMarks: U32 = new Uint32Array(64);
  private indices: I32 = new Int32Array(64);
  private lowLinks: I32 = new Int32Array(64);
  private formulaList: U32 = new Uint32Array(64);
  private stack: U32 = new Uint32Array(64);
  private component: U32 = new Uint32Array(64);
  private cycleMembers: U32 = new Uint32Array(64);
  private cycleGroups: I32 = createInt32(64, -1);

  detect(
    formulaCellIndices: Iterable<number>,
    maxCellIndexExclusive: number,
    forEachFormulaDependency: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void,
    isFormula: (cellIndex: number) => boolean
  ): PackedCycleDetectionResult {
    this.ensureCapacity(maxCellIndexExclusive);
    this.bumpEpochs();

    for (let index = 0; index < this.lastFormulaCount; index += 1) {
      this.cycleGroups[this.formulaList[index]!] = -1;
    }

    let formulaCount = 0;
    for (const cellIndex of formulaCellIndices) {
      this.formulaList[formulaCount] = cellIndex;
      this.cycleGroups[cellIndex] = -1;
      formulaCount += 1;
    }
    this.lastFormulaCount = formulaCount;

    let nextIndex = 0;
    let stackLength = 0;
    let cycleMemberCount = 0;
    let nextCycleGroupId = 0;

    const strongConnect = (cellIndex: number): void => {
      this.visitMarks[cellIndex] = this.visitEpoch;
      this.indices[cellIndex] = nextIndex;
      this.lowLinks[cellIndex] = nextIndex;
      nextIndex += 1;
      this.stack[stackLength] = cellIndex;
      stackLength += 1;
      this.onStackMarks[cellIndex] = this.visitEpoch;

      forEachFormulaDependency(cellIndex, (dependency) => {
        if (!isFormula(dependency)) {
          return;
        }
        if (this.visitMarks[dependency] !== this.visitEpoch) {
          strongConnect(dependency);
          this.lowLinks[cellIndex] = Math.min(this.lowLinks[cellIndex]!, this.lowLinks[dependency]!);
          return;
        }
        if (this.onStackMarks[dependency] === this.visitEpoch) {
          this.lowLinks[cellIndex] = Math.min(this.lowLinks[cellIndex]!, this.indices[dependency]!);
        }
      });

      if (this.lowLinks[cellIndex] !== this.indices[cellIndex]) {
        return;
      }

      let componentLength = 0;
      while (stackLength > 0) {
        const member = this.stack[stackLength - 1]!;
        stackLength -= 1;
        this.onStackMarks[member] = 0;
        this.component[componentLength] = member;
        componentLength += 1;
        if (member === cellIndex) {
          break;
        }
      }

      let isSelfLoop = false;
      if (componentLength === 1) {
        const member = this.component[0]!;
        forEachFormulaDependency(member, (dependency) => {
          if (dependency === member) {
            isSelfLoop = true;
          }
        });
      }

      if (componentLength > 1 || isSelfLoop) {
        const cycleGroupId = nextCycleGroupId;
        nextCycleGroupId += 1;
        for (let index = 0; index < componentLength; index += 1) {
          const member = this.component[index]!;
          this.cycleGroups[member] = cycleGroupId;
          this.cycleMembers[cycleMemberCount] = member;
          cycleMemberCount += 1;
        }
      }
    };

    for (let index = 0; index < formulaCount; index += 1) {
      const cellIndex = this.formulaList[index]!;
      if (this.visitMarks[cellIndex] === this.visitEpoch) {
        continue;
      }
      strongConnect(cellIndex);
    }

    return {
      cycleMembers: this.cycleMembers,
      cycleMemberCount,
      cycleGroups: this.cycleGroups
    };
  }

  private ensureCapacity(maxCellIndexExclusive: number): void {
    if (maxCellIndexExclusive <= this.visitMarks.length) {
      return;
    }

    let capacity = this.visitMarks.length;
    while (capacity < maxCellIndexExclusive) {
      capacity *= 2;
    }

    this.visitMarks = growUint32(this.visitMarks, capacity);
    this.onStackMarks = growUint32(this.onStackMarks, capacity);
    this.indices = growInt32(this.indices, capacity, 0);
    this.lowLinks = growInt32(this.lowLinks, capacity, 0);
    this.formulaList = growUint32(this.formulaList, capacity);
    this.stack = growUint32(this.stack, capacity);
    this.component = growUint32(this.component, capacity);
    this.cycleMembers = growUint32(this.cycleMembers, capacity);
    this.cycleGroups = growInt32(this.cycleGroups, capacity, -1);
  }

  private bumpEpochs(): void {
    this.visitEpoch += 1;
    if (this.visitEpoch !== 0xffff_ffff) {
      return;
    }
    this.visitEpoch = 1;
    this.visitMarks.fill(0);
    this.onStackMarks.fill(0);
  }
}

function createInt32(size: number, fillValue: number): I32 {
  const buffer = new Int32Array(size);
  buffer.fill(fillValue);
  return buffer as I32;
}

function growUint32(buffer: U32, capacity: number): U32 {
  const next = new Uint32Array(capacity);
  next.set(buffer);
  return next as U32;
}

function growInt32(buffer: I32, capacity: number, fillValue: number): I32 {
  const next = new Int32Array(capacity);
  next.set(buffer);
  next.fill(fillValue, buffer.length);
  return next as I32;
}
