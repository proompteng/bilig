# WorkPaper HyperFormula Targeted Reread

Date: `2026-04-13`

Status: `captured prior-art notes for the remaining red lanes`

## Purpose

This note records the concrete HyperFormula patterns that still matter for the remaining WorkPaper
red lanes. It is not a general audit. It is the narrow reread that informed the current delivery
order.

Files reread in the local HyperFormula checkout:

- `/Users/gregkonush/github.com/hyperformula/src/Operations.ts`
- `/Users/gregkonush/github.com/hyperformula/src/DependencyGraph/DependencyGraph.ts`
- `/Users/gregkonush/github.com/hyperformula/src/UndoRedo.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnBinarySearch.ts`
- `/Users/gregkonush/github.com/hyperformula/src/parser/ParserWithCaching.ts`
- `/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts`
- `/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts`
- `/Users/gregkonush/github.com/hyperformula/src/DependencyGraph/RangeVertex.ts`
- `/Users/gregkonush/github.com/hyperformula/src/interpreter/CriterionFunctionCompute.ts`
- `/Users/gregkonush/github.com/hyperformula/src/interpreter/plugin/NumericAggregationPlugin.ts`

## Structural Transforms

What HyperFormula does:

- row and column add, remove, and move are first-class engine transforms
- address mapping, dependency graph state, range state, and lookup state are updated as part of the
  same transform
- structural undo and redo reuse the same transform model

What that means for WorkPaper:

- the remaining structural misses are not evaluator math problems
- `StructuralTransformService` has to own in-place range retargeting and transform-owned dependency
  maintenance
- structural undo and redo cannot be an afterthought

## Exact Lookup Vs Approximate Lookup

What HyperFormula does:

- `ColumnIndex` owns exact lookup with per-value row buckets
- exact lookup refresh is narrow and value-specific
- approximate sorted lookup is a separate path centered on binary search, not a heavier secondary
  index

What that means for WorkPaper:

- `ExactColumnIndexService` and `SortedColumnSearchService` should stay split
- exact after-write work should keep shrinking toward raw bucket maintenance
- approximate-after-write should get narrower, not more elaborate

## Parser And Template Reuse

What HyperFormula does:

- parser caching is based on normalized token streams relative to base address
- repeated row-shifted formulas amortize parse and dependency work instead of recompiling cell by
  cell
- build wiring installs parser and lookup services before graph construction

What that means for WorkPaper:

- `FormulaTemplateNormalizationService` is the correct subsystem
- the remaining row-template problem is prepared binding family reuse, not more compiled-plan
  deduplication alone
- snapshot rebuild will stay red until build normalization gets cheaper

## Range Reuse

What HyperFormula does:

- `RangeVertex` owns range-level caches
- dependent cache invalidation is explicit
- both criteria functions and numeric aggregates reuse smaller cached ranges when possible

What that means for WorkPaper:

- `CriterionRangeCacheService` was the correct cut and should remain range-owned
- `RangeAggregateCacheService` still needs smaller-range extension semantics for
  `aggregate-overlapping-sliding-window`

## Practical Takeaways

The reread confirmed the correct remaining order:

1. `StructuralTransformService`
2. `FormulaTemplateNormalizationService`
3. `RebuildExecutionPolicy`
4. `SortedColumnSearchService`
5. `ExactColumnIndexService`
6. `SuspendedBulkMutationLane`
7. `RangeAggregateCacheService`

The reread also confirmed what not to do:

- do not collapse exact and approximate lookup into one abstraction
- do not try to solve structural misses with evaluator micro-optimizations
- do not treat parser-template build as only a compile-cache problem
- do not add generic caches that are not owned by range or lookup entities
