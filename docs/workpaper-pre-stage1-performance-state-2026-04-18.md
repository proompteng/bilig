# WorkPaper Pre-Stage-1 Performance State

Date: `2026-04-18`

Commit: `1b9e9270`

Branch: `main`

Status: `baseline captured before Stage 1 durable column ownership refactor`

## Purpose

This document records the current performance state of `bilig2` after the Stage 0 scorecard work and
before the Stage 1 lookup-owner refactor.

It is intended to be the comparison point for the next architecture step:

- durable column ownership
- exact and approximate lookup state owned by one persistent column owner
- after-write lookup families as the primary benchmark gate

## Commands Run

The baseline below was captured on the current tree with these commands:

- `pnpm bench:workpaper:competitive`
- `pnpm bench:contracts`

Additional read-only detail was collected from the same tree for the contract metrics shape.

## Competitive Scorecard

Current broad scorecard after Stage 0 reporting changes:

- eligible comparable workloads: `34`
- `WorkPaper` wins: `15`
- `HyperFormula` wins: `19`
- excluded families from broad scorecard:
  - `config-toggle`
  - `dynamic-array`

Important interpretation rule:

- family geomean is only meaningful alongside win counts
- `meanSpeedupGeomean` is a magnitude metric, not a directional leader flag
- broad scorecard leadership should be read from `WorkPaper` wins vs `HyperFormula` wins, not geomean alone

## Family Summary

| Family | Eligible | Comparable | WorkPaper Wins | HyperFormula Wins | Geomean |
| --- | --- | ---: | ---: | ---: | ---: |
| `build` | yes | 5 | 2 | 3 | `2.925x` |
| `rebuild` | yes | 1 | 1 | 0 | `2.265x` |
| `runtime-restore` | yes | 1 | 0 | 1 | `2.017x` |
| `config-toggle` | no | 1 | 1 | 0 | `767.438x` |
| `dirty-execution` | yes | 4 | 3 | 1 | `1.392x` |
| `batch-edit` | yes | 5 | 3 | 2 | `1.416x` |
| `structural-rows` | yes | 3 | 0 | 3 | `2.159x` |
| `structural-columns` | yes | 3 | 0 | 3 | `10.130x` |
| `range-read` | yes | 1 | 0 | 1 | `1.797x` |
| `overlapping-aggregate` | yes | 1 | 1 | 0 | `2.428x` |
| `sliding-window-aggregate` | yes | 1 | 0 | 1 | `4.803x` |
| `conditional-aggregation` | yes | 2 | 2 | 0 | `2.237x` |
| `lookup-exact` | yes | 2 | 2 | 0 | `1.947x` |
| `lookup-after-write` | yes | 2 | 0 | 2 | `1.220x` |
| `lookup-approximate` | yes | 1 | 0 | 1 | `1.167x` |
| `lookup-approximate-after-write` | yes | 1 | 0 | 1 | `24.159x` |
| `lookup-text` | yes | 1 | 1 | 0 | `3.204x` |
| `dynamic-array` | no | 0 | 0 | 0 | `n/a` |

## Current Red Lanes

The biggest current broad-suite losses remain concentrated in the same owner-mismatch families.

Worst current comparable workloads:

| Workload | Faster Engine | Speedup | WorkPaper Mean | HyperFormula Mean |
| --- | --- | ---: | ---: | ---: |
| `structural-insert-columns` | `HyperFormula` | `79.252x` | `23.494 ms` | `0.296 ms` |
| `lookup-approximate-sorted-after-column-write` | `HyperFormula` | `24.159x` | `1.346 ms` | `0.056 ms` |
| `structural-delete-columns` | `HyperFormula` | `6.082x` | `41.403 ms` | `6.808 ms` |
| `aggregate-overlapping-sliding-window` | `HyperFormula` | `4.803x` | `0.300 ms` | `0.062 ms` |
| `build-parser-cache-row-templates` | `HyperFormula` | `4.759x` | `151.532 ms` | `31.839 ms` |
| `partial-recompute-mixed-frontier` | `HyperFormula` | `2.627x` | `11.394 ms` | `4.338 ms` |
| `structural-delete-rows` | `HyperFormula` | `2.482x` | `11.141 ms` | `4.490 ms` |
| `batch-edit-single-column-with-undo` | `HyperFormula` | `2.336x` | `2.704 ms` | `1.157 ms` |

The architectural reading is unchanged:

- structural rows and columns are still red as families
- after-write lookup is still red as a family
- approximate-after-write lookup is still catastrophically red
- runtime restore is still red
- sliding-window aggregate remains red

## Why Stage 1 Next

The next phase should still be durable column ownership because the current scorecard shows:

- `lookup-exact` is already green
- `lookup-after-write` is still `0 / 2`
- `lookup-approximate-after-write` is still `0 / 1`

That is the clearest “owner is still wrong after mutation” signal in the current suite.

The target workloads for Stage 1 remain:

- `lookup-with-column-index-after-column-write`
- `lookup-with-column-index-after-batch-write`
- `lookup-approximate-sorted-after-column-write`

## Contract Baseline

Current contract benchmark state on this tree:

- `pnpm bench:contracts` passed on rerun
- the current truthful contract split now distinguishes:
  - `cold-build`
  - `dirty-execution`
  - `render-commit`
  - `runtime-restore`
  - `patch-emission`

Selected current contract metrics:

| Benchmark | Metric | Current p95 | Budget | Status |
| --- | --- | ---: | ---: | --- |
| `load100k` | `elapsedMs` | `362.07 ms` | `1500 ms` | pass |
| `load250k` | `elapsedMs` | `918.56 ms` | `1500 ms` | pass |
| `edit10k` | `elapsedMs` | `14.66 ms` | `120 ms` | pass |
| `edit10k` | `recalcMs` | `7.94 ms` | `120 ms` | pass |
| `rangeAggregates10k` | `elapsedMs` | `16.09 ms` | `120 ms` | pass |
| `rangeAggregates10k` | `recalcMs` | `7.89 ms` | `100 ms` | pass |
| `topologyEdit10k` | `elapsedMs` | `30.66 ms` | `80 ms` | pass |
| `topologyEdit10k` | `recalcMs` | `6.63 ms` | `80 ms` | pass |
| `renderCommit10k` | `elapsedMs` | `78.39 ms` | `80 ms` | pass, but near threshold |
| `workerWarmStart100k` | `elapsedMs` | `7.84 ms` | `500 ms` | pass |
| `workerWarmStart250k` | `elapsedMs` | `17.42 ms` | `700 ms` | pass |
| `workerVisibleEdit10k` | `visiblePatchMs` | `4.92 ms` | `16 ms` | pass |
| `workerReconnectCatchUp100Pending` | `catchUpMs` | `376.06 ms` | `2000 ms` | pass |

Important nuance:

- `workerVisibleEdit10k.commitMs` is not contract-gated today, but its current p95 is `143.60 ms`
- `renderCommit10k` is currently the tightest budgeted lane, and it is close to the threshold

## Contract Noise Note

During this capture session:

- one earlier `pnpm bench:contracts` run failed with:
  - `renderCommit10k p95 = 227.54 ms`
  - budget = `80.00 ms`
- a fresh rerun on the same tree passed with:
  - `renderCommit10k p95 = 78.39 ms`

That means the render-commit contract lane is currently noisy enough to deserve attention even
though the final baseline pass succeeded.

The practical reading:

- treat `renderCommit10k` as a fragile near-threshold contract
- avoid interpreting a single failing run as the whole current state
- keep that lane under observation during the Stage 1 lookup-owner work even though it is not the primary target

## Engine Counter Coverage

The Stage 0 counter plumbing is visible in current competitive output, but it is not yet universal
across every workload helper path.

Current competitive result coverage:

- workloads with `engineCounters` emitted: `14`
- workloads without `engineCounters` emitted: `22`

Workloads currently emitting counters:

- `build-dense-literals`
- `build-mixed-content`
- `build-many-sheets`
- `single-edit-chain`
- `single-edit-fanout`
- `single-formula-edit-recalc`
- `batch-edit-single-column`
- `batch-edit-multi-column`
- `range-read-dense`
- `lookup-no-column-index`
- `lookup-with-column-index`
- `lookup-approximate-sorted`
- `lookup-text-exact`
- `dynamic-array-filter`

Important current limitation:

- many additional-workload helper paths still emit legacy benchmark payloads without `engineCounters`
- that means the current benchmark truth work is materially better than before, but still incomplete as a fully universal instrumentation layer

This should not block Stage 1, but it should be remembered when interpreting counter-based deltas
for structural, restore, and after-write helper workloads.

## Baseline Interpretation

Before the Stage 1 refactor, the repo is in this state:

- the scorecard is materially more truthful than before
- the broad suite is still red overall: `15` wins vs `19` losses on eligible comparable workloads
- the strongest remaining red families are still:
  - `structural-columns`
  - `structural-rows`
  - `lookup-after-write`
  - `lookup-approximate-after-write`
  - `runtime-restore`
  - `sliding-window-aggregate`
- contract benches are mostly healthy, but `renderCommit10k` is close enough to budget that it can flap

That is the performance state we are carrying into Stage 1.
