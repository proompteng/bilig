import type { RuntimeFormula, U32 } from '../runtime-state.js'

interface CycleEvaluationFormulaLookup {
  readonly get: (cellIndex: number) => RuntimeFormula | undefined
  readonly has: (cellIndex: number) => boolean
}

export interface CycleEvaluationNode {
  readonly kind: 'formula' | 'cycle'
  readonly order: number
  readonly formulaCellIndices: number[]
}

export function buildCycleEvaluationNodes(args: {
  readonly ordered: readonly number[] | U32
  readonly orderedCount: number
  readonly formulas: CycleEvaluationFormulaLookup
  readonly cycleGroupIds: ArrayLike<number | undefined>
  readonly forEachFormulaDependencyCell: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void
}): readonly CycleEvaluationNode[] | undefined {
  const nodeIndexByKey = new Map<string, number>()
  const nodeIndexByCell = new Map<number, number>()
  const dirtyFormulaOrder = new Map<number, number>()
  const nodes: CycleEvaluationNode[] = []
  let hasCycleMembers = false

  for (let orderedIndex = 0; orderedIndex < args.orderedCount; orderedIndex += 1) {
    const cellIndex = args.ordered[orderedIndex]!
    if (!args.formulas.has(cellIndex)) {
      continue
    }
    dirtyFormulaOrder.set(cellIndex, orderedIndex)
    const cycleGroupId = args.cycleGroupIds[cellIndex] ?? -1
    const kind: CycleEvaluationNode['kind'] = cycleGroupId >= 0 ? 'cycle' : 'formula'
    const key = kind === 'cycle' ? `cycle:${cycleGroupId}` : `formula:${cellIndex}`
    let nodeIndex = nodeIndexByKey.get(key)
    if (nodeIndex === undefined) {
      nodeIndex = nodes.length
      nodeIndexByKey.set(key, nodeIndex)
      nodes.push({
        kind,
        order: orderedIndex,
        formulaCellIndices: [],
      })
    }
    if (kind === 'cycle') {
      hasCycleMembers = true
    }
    nodes[nodeIndex]!.formulaCellIndices.push(cellIndex)
    nodeIndexByCell.set(cellIndex, nodeIndex)
  }

  if (!hasCycleMembers) {
    return undefined
  }

  nodes.forEach((node) => {
    node.formulaCellIndices.sort((left, right) => (dirtyFormulaOrder.get(left) ?? 0) - (dirtyFormulaOrder.get(right) ?? 0))
  })

  const outgoing = nodes.map(() => new Set<number>())
  const indegree = new Int32Array(nodes.length)
  for (let nodeIndex = 0; nodeIndex < nodes.length; nodeIndex += 1) {
    const node = nodes[nodeIndex]!
    for (let formulaIndex = 0; formulaIndex < node.formulaCellIndices.length; formulaIndex += 1) {
      const formula = args.formulas.get(node.formulaCellIndices[formulaIndex]!)
      if (!formula) {
        continue
      }
      args.forEachFormulaDependencyCell(formula.cellIndex, (dependencyCellIndex) => {
        const dependencyNodeIndex = nodeIndexByCell.get(dependencyCellIndex)
        if (dependencyNodeIndex === undefined || dependencyNodeIndex === nodeIndex || outgoing[dependencyNodeIndex]!.has(nodeIndex)) {
          return
        }
        outgoing[dependencyNodeIndex]?.add(nodeIndex)
        indegree[nodeIndex] = (indegree[nodeIndex] ?? 0) + 1
      })
    }
  }

  const ready: number[] = []
  const visited = new Uint8Array(nodes.length)
  for (let nodeIndex = 0; nodeIndex < nodes.length; nodeIndex += 1) {
    if (indegree[nodeIndex] === 0) {
      ready.push(nodeIndex)
    }
  }
  ready.sort((left, right) => nodes[left]!.order - nodes[right]!.order)

  const orderedNodes: CycleEvaluationNode[] = []
  while (ready.length > 0) {
    const nodeIndex = ready.shift()!
    if (visited[nodeIndex] === 1) {
      continue
    }
    visited[nodeIndex] = 1
    orderedNodes.push(nodes[nodeIndex]!)
    for (const dependentNodeIndex of outgoing[nodeIndex]!) {
      indegree[dependentNodeIndex] = (indegree[dependentNodeIndex] ?? 0) - 1
      if (indegree[dependentNodeIndex] === 0) {
        ready.push(dependentNodeIndex)
      }
    }
    ready.sort((left, right) => nodes[left]!.order - nodes[right]!.order)
  }

  if (orderedNodes.length === nodes.length) {
    return orderedNodes
  }

  const remaining = nodes.filter((_node, nodeIndex) => visited[nodeIndex] === 0).toSorted((left, right) => left.order - right.order)
  return [...orderedNodes, ...remaining]
}
