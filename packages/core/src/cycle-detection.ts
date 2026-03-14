export interface CycleDetectionResult {
  inCycle: Set<number>;
  cycleGroups: Map<number, number>;
}

export function detectFormulaCycles(
  formulaCellIndices: Iterable<number>,
  getFormulaDependencies: (cellIndex: number) => Iterable<number>
): CycleDetectionResult {
  const indices = new Map<number, number>();
  const lowLinks = new Map<number, number>();
  const stack: number[] = [];
  const onStack = new Set<number>();
  const inCycle = new Set<number>();
  const cycleGroups = new Map<number, number>();
  const formulaSet = new Set(formulaCellIndices);
  let nextIndex = 0;
  let nextCycleGroupId = 0;

  const strongConnect = (cellIndex: number) => {
    indices.set(cellIndex, nextIndex);
    lowLinks.set(cellIndex, nextIndex);
    nextIndex += 1;
    stack.push(cellIndex);
    onStack.add(cellIndex);

    for (const dependency of getFormulaDependencies(cellIndex)) {
      if (!formulaSet.has(dependency)) {
        continue;
      }
      if (!indices.has(dependency)) {
        strongConnect(dependency);
        lowLinks.set(cellIndex, Math.min(lowLinks.get(cellIndex)!, lowLinks.get(dependency)!));
        continue;
      }
      if (onStack.has(dependency)) {
        lowLinks.set(cellIndex, Math.min(lowLinks.get(cellIndex)!, indices.get(dependency)!));
      }
    }

    if (lowLinks.get(cellIndex) !== indices.get(cellIndex)) {
      return;
    }

    const component: number[] = [];
    while (stack.length > 0) {
      const member = stack.pop()!;
      onStack.delete(member);
      component.push(member);
      if (member === cellIndex) {
        break;
      }
    }

    const isSelfLoop =
      component.length === 1 && [...getFormulaDependencies(component[0]!)].some((dependency) => dependency === component[0]);

    if (component.length > 1 || isSelfLoop) {
      const cycleGroupId = nextCycleGroupId;
      nextCycleGroupId += 1;
      component.forEach((member) => {
        inCycle.add(member);
        cycleGroups.set(member, cycleGroupId);
      });
    }
  };

  formulaSet.forEach((cellIndex) => {
    if (!indices.has(cellIndex)) {
      strongConnect(cellIndex);
    }
  });

  return {
    inCycle,
    cycleGroups
  };
}
