# fast-check Best Practices
## Date: 2026-04-16
## Status: active guidance

## Why this document exists

`bilig` already uses `fast-check` heavily for engine, formula, server, browser, and boundary fuzzing.

This document captures the best practices that should govern future usage in this repo, based on:

- the official `fast-check` docs
- the upstream `fast-check` source in `/Users/gregkonush/github.com/fast-check`
- what has already worked well in `bilig`
- the concrete mistakes and fixes we have already gone through

It is intentionally practical. It is not a generic introduction to property-based testing.

## Upstream references reviewed

- `https://fast-check.dev/docs/advanced/race-conditions/`
- `https://fast-check.dev/docs/advanced/model-based-testing/`
- `https://fast-check.dev/docs/migration-guide/from-3.x-to-4.x/`
- `https://fast-check.dev/docs/configuration/timeouts/`
- `/Users/gregkonush/github.com/fast-check/packages/fast-check/src/check/model/ModelRunner.ts`
- `/Users/gregkonush/github.com/fast-check/packages/fast-check/src/arbitrary/_internals/interfaces/Scheduler.ts`
- `/Users/gregkonush/github.com/fast-check/packages/fast-check/src/arbitrary/commands.ts`
- `/Users/gregkonush/github.com/fast-check/packages/fast-check/test/e2e/ReplayCommands.spec.ts`
- `/Users/gregkonush/github.com/fast-check/packages/vitest/README.md`

## Core principles

### 1. Use the smallest property that defends one real guarantee

A property should fail with one obvious explanation.

Good:

- projection parity
- pending journal convergence
- request/response routing parity
- structural undo parity

Bad:

- one giant property that mixes storage, sync, viewport, and UI focus

If a failure could plausibly belong to three systems, the property is too broad.

### 2. Prefer semantic invariants over implementation counters

The property should assert what must stay true for the product, not what happened internally.

Good:

- final snapshot equals replayed snapshot
- pending mutations match the model
- delivered responses match requested operations

Bad:

- callback count equals N
- helper X was called before helper Y

### 3. Keep the model simpler than the real system

Upstream explicitly warns against making the model a carbon copy of the implementation.

The model should track only the semantic facts needed for the guarantee:

- authoritative values
- pending statuses
- expected order
- expected visible selection/range

Do not duplicate the engine, runtime, or storage layer in the test.

## Property suites

### Use `runProperty` for stateless or lightly stateful invariants

Use this for:

- parser/canonicalization invariants
- codec roundtrips
- projection equality
- import/export parity

### Use `runModelProperty` for command-based state machines

Use this when:

- actions have preconditions
- undo/redo matters
- command history matters
- the model evolves step by step

In this repo that includes engine history and metadata suites.

### Use `runScheduledProperty` for race and ordering boundaries

Use this only where ordering is the actual product risk:

- sync relay
- request/response transport
- local store mutation races
- runtime local/authoritative interleavings

Do not use scheduled fuzz just because a test is asynchronous.

## Commands and replay

### Always preserve `replayPath` for `fc.commands(...)`

For command-based properties, `seed` and `path` are not enough for exact replay.

When using `fc.commands(...)`, the exact replay requires `replayPath`.

In `bilig`, the harness must:

- capture `replayPath` from command failures
- store it in artifacts and replay fixtures
- pass it back into `fc.commands(..., { replayPath })` during replay

If this is missing, command replay is approximate, not exact.

### Keep commands readable

Commands should have:

- a real `check`
- a small `run`
- a meaningful `toString()`

Do not hide a huge state machine inside one command.

## Scheduler usage

### Prefer `waitFor(...)` or `waitIdle()`

Current upstream guidance is:

- prefer `waitFor(promise)` when you know what completion you are waiting for
- prefer `waitIdle()` when you want the scheduler to drain all reachable scheduled work

### Avoid `waitAll()`

`waitAll()` is deprecated upstream.

Do not add new `waitAll()` usage.

When updating older code:

- replace `waitAll()` with `waitIdle()` when draining all reachable scheduled work
- replace `waitAll()` with `waitFor(Promise.all(...))` when waiting for specific scheduled operations

### Schedule the async boundary, not the whole test

Good:

- wrap the asynchronous API with `scheduleFunction`
- or schedule independent actions/promises with labels

Bad:

- schedule giant outer wrappers that themselves control unrelated sync logic
- create scheduler deadlocks by awaiting scheduled promises before the scheduler is allowed to release them

### Do not depend on unrelated scheduled tasks inside the model

Upstream is explicit here: model logic should not need other scheduled tasks to finish in order to make progress.

The model may trigger scheduled work, but it should not be written as “wait for some other random scheduled thing and then continue.”

## Timeouts and budgets

### Use runner-level budgets for fuzz lanes

For repo fuzz lanes, prefer `interruptAfterTimeLimit` and controlled `numRuns` budgets.

This is the right level for:

- `default`
- `main`
- `nightly`
- `replay`

### Do not use timeouts to hide semantic failures

Never:

- pad timeouts to make CI green
- lower `numRuns` just to avoid failures
- convert a failing invariant into a looser one to avoid timeouts

Timeout changes are only valid when:

- the property shape is already correct
- the suite is demonstrably bounded
- the budget belongs to the lane policy, not to bug concealment

### Keep scheduled suites cheaper than pure properties

Scheduled suites are expensive by nature.

Their budget should be smaller than generic property suites for the same lane.

That is not cheating. That is acknowledging the real cost model of scheduler-driven interleavings.

## Arbitrary design

### Generate meaningful states, not random noise

Prefer:

- realistic addresses
- realistic workbook shapes
- bounded but semantically rich transitions
- generator bias toward known-dangerous boundaries

Avoid:

- giant arbitrary spaces with almost no semantic value
- “random everything” generators that mostly produce boring no-ops

### Bound complexity intentionally

Bounds are allowed when they are product-relevant.

Good:

- a fixed number of tracked cells in a runtime journal property
- bounded action counts in scheduled suites
- targeted address sets for selection geometry

Bad:

- tiny bounds chosen only to avoid failures
- arbitrary `maxLength` values that weaken the guarantee without product reason

## Replay artifacts

### Every high-value fuzz failure should end as a durable replay

The lifecycle should be:

1. fuzz finds a failure
2. artifact captures it
3. replay fixture is promoted
4. deterministic regression is added when the bug deserves one

### Prefer committed replay fixtures over seed-only folklore

Seeds are useful.

Committed replays are better.

They survive:

- runner changes
- ordering changes
- helper refactors

## Vitest integration

### Raw `fast-check` plus `vitest` is acceptable

This repo uses raw `fast-check` inside `vitest`, which is fine.

`@fast-check/vitest` is useful, but it is optional. The repo does not need to migrate just to be “modern.”

### Keep browser fuzz isolated

Browser fuzz should be:

- explicitly tagged
- lane-controlled
- focused on browser authority invariants

Do not let browser fuzz become the first line of defense for engine semantics.

## Repo-specific guidance for `bilig`

### Preferred split

- `packages/core`: semantic engine truth
- `packages/formula`: parser/translation/evaluation truth
- `apps/bilig`: sync/projection/server truth
- `apps/web`: runtime/projection/browser truth
- boundary packages like `worker-transport` and `storage-browser`: direct package-owned guarantees

### Preferred lane shape

- `default`: high-signal semantic core plus the most critical boundaries
- `main`: full maintained fuzz surface
- `nightly`: broader and more expensive exploration
- `replay`: exact fixture reproduction

### New work should follow these patterns

- command-based suites must support `replayPath`
- scheduler suites must use `waitFor` or `waitIdle`
- browser fuzz must handle stack startup races explicitly
- replay corpora should grow alongside new fuzz surfaces

## Anti-patterns

Do not:

- use `waitAll()` in new scheduler tests
- claim command replay is exact without `replayPath`
- let the model mirror the implementation
- use timeout inflation as a correctness fix
- write one enormous property that hides subsystem ownership
- make browser fuzz responsible for core engine semantics

## Practical checklist

Before adding a new fuzz suite, ask:

1. What exact product guarantee does this suite defend?
2. Should this be `runProperty`, `runModelProperty`, or `runScheduledProperty`?
3. Is the model simpler than the real system?
4. Are the arbitraries generating meaningful states?
5. If this uses `fc.commands(...)`, is `replayPath` supported?
6. If this uses scheduler, am I using `waitFor` or `waitIdle` correctly?
7. If this fails in CI, can I promote it into a durable replay fixture?

If the answer to any of those is “no,” the suite design is not finished yet.
