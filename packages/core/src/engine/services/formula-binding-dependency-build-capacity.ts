import { growUint32 } from '../../engine-buffer-utils.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

export function ensureFormulaBindingDependencyBuildCapacity(
  args: CreateEngineFormulaBindingServiceArgs,
  cellCapacity: number,
  dependencyCapacity: number,
  symbolicRefCapacity = 0,
  symbolicRangeCapacity = 0,
): void {
  if (cellCapacity > args.getDependencyBuildSeen().length) {
    args.setDependencyBuildSeen(growUint32(args.getDependencyBuildSeen(), cellCapacity))
  }
  if (cellCapacity > args.getDependencyBuildCells().length) {
    args.setDependencyBuildCells(growUint32(args.getDependencyBuildCells(), cellCapacity))
  }
  if (dependencyCapacity > args.getDependencyBuildEntities().length) {
    args.setDependencyBuildEntities(growUint32(args.getDependencyBuildEntities(), dependencyCapacity))
  }
  if (dependencyCapacity > args.getDependencyBuildRanges().length) {
    args.setDependencyBuildRanges(growUint32(args.getDependencyBuildRanges(), dependencyCapacity))
  }
  if (dependencyCapacity > args.getDependencyBuildNewRanges().length) {
    args.setDependencyBuildNewRanges(growUint32(args.getDependencyBuildNewRanges(), dependencyCapacity))
  }
  if (symbolicRefCapacity > args.getSymbolicRefBindings().length) {
    args.setSymbolicRefBindings(growUint32(args.getSymbolicRefBindings(), symbolicRefCapacity))
  }
  if (symbolicRangeCapacity > args.getSymbolicRangeBindings().length) {
    args.setSymbolicRangeBindings(growUint32(args.getSymbolicRangeBindings(), symbolicRangeCapacity))
  }
}
