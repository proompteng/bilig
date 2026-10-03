# Public APIs

## Current state

- `@bilig/core` now implements the documented range-mutation, undo/redo, selection, and sync-state surface.
- `SelectionState` is additive: existing callers can keep using `sheetName` and `address`, while newer callers can use `anchorAddress`, `range`, and `editMode`.
- `connectSyncClient` is live. `apps/web` is already worker-first by default, but the remote sync service is not a closed durable worksheet backend yet.
- `@bilig/binary-protocol` is already a real wire protocol for sync frames, but the authoritative replicated workbook op family is still narrower than the full local engine surface.
- `@bilig/agent-api` is currently a binary envelope around JSON payloads, not yet a fully typed binary request/response/event schema.

## Stable packages

`@bilig/workpaper` is the current external npm adoption path for WorkPaper
calculation and Excel workbook import/export. Its `@bilig/workpaper/xlsx`
subpath exposes the repository XLSX importer/exporter without requiring a second
public npm package. The other package surfaces listed here are stable repository
package boundaries; not every package name is provisioned on npm yet.

- `@bilig/core`
- `@bilig/workpaper`
- `@bilig/excel-import`
- `@bilig/formula`
- `@bilig/wasm-kernel`
- `@bilig/workbook`
- `@bilig/grid`
- `@bilig/binary-protocol`
- `@bilig/worker-transport`
- `@bilig/agent-api`
- `@bilig/storage-server`
- `@bilig/excel-fixtures`

## Agent-first workbook surface

`@bilig/workbook` is the generic public package for consumer-defined workbook
models. It does not ship business-model templates and does not depend on
`@bilig/core`, `@bilig/workpaper`, `@bilig/agent-api`, `zod`, or `effect`.

`@bilig/workbook` defines generic, inspectable workbook intent for agents and
runtimes. It does not depend on hardcoded business models or human spreadsheet
UI assumptions.

The stable contract has four layers.

Package identity is intentionally explicit:

| Package            | Choose when                                                                                   | Do not use for                                     |
| ------------------ | --------------------------------------------------------------------------------------------- | -------------------------------------------------- |
| `@bilig/workbook`  | Defining generic agent intent, refs, formulas, checks, plan data, schemas, and proof handoff. | Calculating formulas or owning workbook state.     |
| `@bilig/workpaper` | Running workbook tools, MCP, or product workflows around persisted WorkPaper state.           | Designing a reusable model API for other runtimes. |
| `@bilig/workpaper` | Owning workbook state inside Node with formula recalculation and import/export.               | Publishing generic agent intent contracts.         |
| `@bilig/core`      | Implementing calculation or mutation internals.                                               | Consumer-facing agent model definitions.           |

The root export remains the ordinary agent path: models, refs, checks, formulas,
plans, runtime proof, command results, schemas, and low-level ops. For agents
that want smaller import maps, the package also publishes layered subpaths:
`@bilig/workbook/model`, `@bilig/workbook/prepare`,
`@bilig/workbook/find`, `@bilig/workbook/check`, `@bilig/workbook/formula`,
`@bilig/workbook/verify`, `@bilig/workbook/runtime`,
`@bilig/workbook/command`, `@bilig/workbook/features`,
`@bilig/workbook/testing`, and `@bilig/workbook/schema`.

### Agent model contract

Consumers define models. Bilig does not ship hardcoded revenue, quote, forecast,
prepaid, or other business models in this package.

- `defineModel` freezes model metadata and action manifests.
- `find` returns generic workbook refs through `findTable`, `findColumn`,
  `findRange`, `findName`, and `findRows`.
- `checks` declares generic facts the runtime must prove.
- `find`, `check`, and `formula` are frozen helper namespaces, so consumers
  cannot mutate the imported public tool surface after an agent inspects it.
- `actions` build portable workbook intent with `writeFormula`, `writeValue`,
  `format`, `clear`, or a guarded low-level op.
- `inspectModel` returns a frozen model manifest without running find, checks,
  or actions.
- `checkWorkbookModelDescription` validates a transported `describeModel`
  manifest, including constrained action input descriptions, without loading
  consumer model code.
- `planWorkbookAction` returns a frozen result wrapper containing either a
  frozen plan or a frozen structured failure, including `invalid_model` when the
  supplied model manifest is not data-safe.
- `prepareWorkbookAction` is the canonical preflight for agents that want the
  full handoff in one call: it plans, verifies, emits JSON-safe plan data,
  computes `planId`, and describes runtime requirements without importing an
  engine.
- `verifyPlan` proves static consistency without importing or starting an
  engine.
- `describeModel`, `describeRef`, `describePlan`, `describePlanResult`, and
  `describeRunResult` return frozen JSON-safe objects for logs, approvals,
  tools, and tests.
- `runWorkbookPlan` and `runWorkbookAction` return frozen run results, so status,
  changes, checks, errors, apply proof, undo refs, and unverified proof notes
  stay stable after inspection.
- `toPlanData` returns JSON-safe plan data for runtime handoff.
- `checkPlanData` returns structured path-based issues for transported plan
  payloads before hydration.
- `hydratePlanData` restores frozen refs and helper methods from transported
  plan data without reconstructing the consumer's private `refs` shape.
- `verifyPlanData` proves a JSON-transported plan description after it has lost
  local helper methods and consumer-private `refs` object shape.
- `checkInput`, `checkPlanData`, `verifyPlan`, `verifyModel`, and
  `verifyWorkbookReadbacks` return frozen verdict containers, generated issue
  arrays, and proof checks, so an inspected result stays stable.
- `checkWorkbookReadbackProof(data)` validates transported `{ checks,
readbacks }` proof in one call and returns frozen verified proof or stable
  readback issues. `isWorkbookReadbackProof(data)` is the boolean guard over the
  same proof boundary.
- `workbookJsonSchemas`, `workbookJsonSchemaHashes`, and package `fixtures/`
  expose the public ref, model-manifest, plan, runtime-requirements, command,
  run-result, and readback-proof contract for agents and non-TypeScript
  consumers without adding a schema framework dependency.
- Feature, command, receipt, result, requirements, and adapter check APIs return
  frozen verdict containers too.

The model contract is intentionally data-first. Refs are frozen. Helper methods
such as `table.column("Amount")` and `rows.column("Amount")` are non-enumerable.
Action input is plain JSON, cloned into the plan, and described with small
metadata rather than a schema dependency.
Action manifests are frozen null-prototype maps, and planning only runs own
manifest actions. Generic action names such as `toString`, `constructor`, or
`__proto__` can be explicit own actions, but inherited prototype fields are not
part of the executable surface.
Imported helper namespaces such as `find`, `check`, and `formula` are frozen,
and factory-created check/find helper objects are frozen too. Formula helper
arrays are frozen even for direct refs or literal operands.
For transport, `toWorkbookRefData` turns any ref into plain data,
`isWorkbookRefData` validates that data, `collectWorkbookRefData` discovers refs
without requiring hidden methods, and `hydrateWorkbookRef` /
`hydrateWorkbookRefs` restores the ergonomic helpers when local code needs them
again.
Ref discovery and ref hydration only inspect enumerable own data properties.
Accessor properties are ignored instead of invoked, so consumer-defined getters
cannot run during planning, verification, logging, or transport hydration.
Array entries follow the same data-only rule, and ref cloning copies known ref
fields rather than spreading extra enumerable properties.
Selector creation is data-only too. `findTable`, `findColumn`, `findRange`, and
`findRows` reject accessor-backed option fields, row predicates, and header
entries before any hidden getter can run.
For full action handoff, `toPlanData(plan)` returns executable JSON-safe plan
data. Runtimes can call `hydratePlanData(data)`, or pass that data directly to
`describeRuntimeRequirements(data)` and `runWorkbookPlan(data, adapter)`.
Hydrated transported plans expose `refs: { refsUsed }`, keeping execution
generic instead of depending on the consumer's model-shaped `refs` object.
Agents that receive untrusted JSON can call `checkPlanData(data)` first to get
boring `{ status, issues }` diagnostics such as `commands[0]` or `refsUsed[2]`
instead of catching a generic hydration failure. Transported plan arrays are
also data-only: holes, non-enumerable entries, or accessor-backed entries are
rejected without invoking getters before hydration or execution. Valid
`checkPlanData` results return canonical plan data, so caller-owned scratch
fields cannot change `workbookPlanId` or the payload later hydrated by runtime
code.
The plan root and nested plan entries such as commands, changes, checks,
formula labels, and expectations must be record-shaped payloads; arrays with
attached fields are rejected even when the attached fields look like valid plan
data.

### Runtime adapter contract

Runtimes execute plans. `@bilig/workbook` only defines the handoff.

- `describeRuntimeRequirements(plan)` tells an adapter which commands must be
  applied, which targets must be read back, and which checks need runtime proof.
- `checkRuntimeRequirements(data)` returns structured path-based issues for
  transported runtime requirement payloads before adapter validation.
- `runWorkbookPlan(planOrData, adapter)` accepts either the in-memory plan or
  transported plan data, returns path-based `invalid_plan_data` errors for
  malformed transported data, refuses to apply invalid plans, calls the adapter,
  verifies readback expectations, verifies generic checks, and returns a boring
  `WorkbookRunResult`.
- `checkRuntimeAdapter(planOrRequirements, adapter)` checks that the adapter has
  the required `apply`, `read`, and `verifyChecks` methods before mutation.
- `checkWorkbookRunAdapter(planOrData, adapter)` and
  `assertWorkbookRunAdapter(planOrData, adapter)` live under
  `@bilig/workbook/testing`. They run the adapter in strict mode and validate
  the JSON-safe run description without depending on a test runner or engine.
- `adapter.apply(plan)` owns runtime mutation and may return `previewOps`,
  `appliedOps`, JSON-safe `proof`, and an undo ref.
- `adapter.read(targets, plan)` returns semantic readbacks for checks such as
  `valueEquals` and `formulaEquals`.
- `adapter.verifyChecks(checks, plan)` may only change check `status` or add
  JSON-safe `proof`; it cannot rewrite the check contract.

Adapter-returned apply results, undo refs, apply errors, and check verifier
output are accepted from own payload fields only. Prototype-inherited fields do
not satisfy runtime proof. Apply-result objects and verifier check objects must
be record-shaped payloads, not arrays with attached fields. Adapter-returned op
arrays and verifier proof must be plain data properties; runtime evidence arrays
must contain own enumerable data entries. Holes, non-enumerable entries, and
accessor-backed fields are rejected before cloning or preview/apply comparison,
including non-enumerable fields that op guards would otherwise read directly.

When `previewOps` and `appliedOps` are both present, `runWorkbookPlan` reports
whether runtime apply matched preview. When they are missing, the result reports
an unverified apply fact. Agents that need fail-closed execution can call
`runWorkbookPlan(plan, adapter, { requireApplyProof: true })`, or use
`{ strict: true }` to require checks before mutation, apply proof, plan-id
proof, base/applied revision proof, no unverified apply facts, concrete command
ops or explicit no-op proof, resolved-ref proof, and passed-check proof with one
option.
Plans with no apply requirements skip mutation entirely: a readback-only or
check-only model can run with only `read` or `verifyChecks`, and the result has
no `apply` summary because nothing was supposed to mutate.
Failed results include `changed` as a concrete answer, not an implied absence:
`changed: []` means no runtime apply succeeded before the failure, while
post-apply proof failures keep the planned change summaries and adapter-provided
undo metadata. If an adapter returns a failed apply with `appliedOps: []` and no
undo metadata, `changed` stays empty because the runtime proved no ops were
applied.

Readback-backed checks attach proof to passed checks. A result can therefore
show the intended action, bound refs, planned commands and ops, adapter
requirements, preview ops, applied ops, preview/apply match, readback values or
formulas, check statuses, proof objects, undo metadata, and remaining unverified
facts without relying on rendered spreadsheet state.

### Feature command handoff

Runtimes may expose workbook extensions as commands, projection interceptors,
and UI contributions. The public package still stays data-only.

Use `@bilig/workbook/features` for this advanced runtime extension surface.
Feature plugin manifests and UI contribution helpers are intentionally not on
the root import path. Ordinary agent models should stay on `defineModel`,
`prepareWorkbookAction`, `runWorkbookPlan`, checks, and formulas.

- `defineWorkbookFeaturePlugin` freezes extension metadata.
- `checkWorkbookFeaturePlugin(data)` validates consumer-provided feature
  manifests before registration, including nested command input-description and
  UI contribution metadata paths. Feature manifests, command descriptors,
  projection interceptors, and UI contributions must be object-record data, not
  class/custom-prototype instances.
- `checkWorkbookCommandRequest(data)` validates transported command requests
  before dispatch and returns path issues such as `featureId`, `commandId`,
  `category`, `mode`, and nested JSON payload paths such as `input.rows[1]`.
- `normalizeWorkbookCommandRequest(data)` returns a frozen request after the
  same validation.
- `checkWorkbookCommandBundle(data)` validates an ordered runtime handoff
  before execution. It requires `targetRevision`, `idempotencyKey`, and
  non-empty `commands`; rejects unknown command kinds and duplicate supplied
  command ids; canonicalizes declared touched ranges; rejects mutation requests
  and ops unless they set
  `destructive: true`; requires touched ranges for destructive commands when
  `scope.maxTouchedCells` is present; semantically counts touched cells and
  rejects bundles over that limit; and returns a `WorkbookCommandResult` with
  normalized touched ranges. JSON Schema covers the transported shape for this
  field, while the checker is the authority for range cardinality.
- `normalizeWorkbookCommandBundle(data)` returns the frozen bundle after the
  same validation.
- `workbookCommandResultForReceipts(bundle, receipts, { revision, undo })`
  converts runtime receipt evidence into an applied, previewed, rejected, or
  noop `WorkbookCommandResult`. It validates receipt count and request identity,
  rejects receipt changed ranges outside declared command `touchedRanges`,
  aggregates changed ranges, records preview/apply match status, and carries undo
  metadata without importing the engine.
- `checkWorkbookCommandResult(data)` and
  `normalizeWorkbookCommandResult(data)` validate transported command results
  before an agent trusts them. Accepted results are only pre-runtime handoff
  acknowledgements; settled proof fields such as receipts, changed ranges,
  revision, undo, and errors are allowed only on receipt-backed runtime result
  statuses.
- `checkWorkbookCommandResultForBundle(bundle, data)` recomputes result
  `status`, `matched`, `changedRanges`, and `errors` from receipts, and rejects
  receipt changed ranges outside the command scopes declared by the bundle. Stored
  command proof cannot carry a hand-edited summary that contradicts runtime
  evidence.
- `checkWorkbookCommandReceipt(data)` validates transported receipt evidence
  before an agent trusts runtime extension output, including nested `proof` and
  `metadata` payload paths.
- Every feature-handoff validator returns a frozen verdict container with frozen
  issue arrays.
- `normalizeWorkbookCommandReceipt(receipt)` and
  `workbookCommandReceiptOpsMatch(receipt)` give agents boring receipt proof
  after preview or apply. Receipt ops are normalized into frozen data and
  compared by canonical op content, so property insertion order cannot create a
  false mismatch and invalid op arrays cannot report a trusted match.
- Feature manifests, command requests, and command receipts are validated from
  own payload fields only; prototype-inherited fields cannot satisfy the public
  transport contract. Manifest and receipt arrays must contain own enumerable
  data entries; receipt ops, undo ops, changed ranges, and errors must be data
  properties. Receipt changed ranges are canonicalized through the same workbook
  range normalizer used by command scopes. Holes, non-enumerable entries,
  invalid range addresses, and accessor-backed receipt proof are rejected before
  freezing or preview/apply comparison.
- Receipt statuses are semantic: previewed receipts cannot include applied
  proof, applied receipts cannot include errors and must carry applied
  evidence, rejected receipts cannot claim changed workbook proof, and noop
  receipts cannot claim changed ranges or ops.
- Frozen vocabularies such as `workbookCommandCategories`,
  `workbookCommandExecutionModes`, and `workbookCommandReceiptStatuses` let tool
  builders expose exact supported values without a schema framework. Advanced
  feature-extension vocabularies live under `@bilig/workbook/features`.

Feature command handoff does not move execution into `@bilig/workbook`. The
runtime owns command semantics. The package only normalizes the manifest,
validates transported requests and bundles, and describes receipt evidence.
The current `@bilig/agent-api` workbook review flow converts its app-specific
`WorkbookAgentCommandBundle` into this generic `WorkbookCommandBundle` before
preview and before authoritative apply. That keeps the rich app command set
outside the public package while forcing the runtime handoff through
revision-bound, idempotent, destructive, range-normalized data.
After authoritative apply, `apps/bilig` requires a validated
`WorkbookCommandResult` for the exact accepted bundle and applied revision before
it writes an execution record. That result is persisted as `command_result_json`,
so later agents can inspect the same generic proof without replaying human
spreadsheet state.
Low-level `op` commands use deterministic `workbook-op` receipt identity through
`workbookOpCommandReceiptIdentity` and `workbookOpCommandReceipt`; arbitrary
receipt ids are rejected when a result is checked against its bundle.

### Escape hatches

The preferred path is the small model API. Escape hatches stay explicit:

- `findRange` is available when a consumer really has a concrete range.
- `formula.raw(source, { inputs })` is available for formulas outside the helper
  set while keeping dependencies inspectable. Add `labels: [{ name, ref }]`
  when raw formula text uses custom ref tokens.
- `workbook.addOp(op, { target?, message? })` carries the existing low-level
  workbook operation language through the same plan, description, and
  verification flow.

Escape hatches do not make the package domain-specific. They keep the public API
generic while still letting advanced runtimes use the lower-level workbook
language.
Low-level op guards accept only plain own-field payloads. Prototype-inherited
op fields, nested ranges, and batch clocks do not satisfy `isWorkbookOp` or
`isEngineOpBatch`. Accessor-backed required fields, nested fields, and op-array
entries are rejected from descriptors without invoking getters.

Domain examples such as revenue, forecast, and quote approval belong under
consumer app or WorkPaper examples, not inside `@bilig/workbook` itself.

Full export surface:

- `defineModel`
- `buildWorkbookActionPlan`
- `planWorkbookAction`
- `prepareWorkbookAction`
- `inspectModel`
- `collectWorkbookRefs`
- `collectWorkbookRefData`
- `findTable`, `findColumn`, `findRange`, `findName`, and `findRows`
- `find`
- `workbookRefKinds`
- `isWorkbookRefKind`
- `checkWorkbookRef`
- `isWorkbookRef`
- `checkWorkbookRefData`
- `isWorkbookRefData`
- `toWorkbookRefData`
- `hydrateWorkbookRef`
- `hydrateWorkbookRefs`
- `workbookRowOperators`
- `workbookRowOperatorValueTypes`
- `isWorkbookRowOperator`
- `isWorkbookRowValueCompatible`
- `check`
- `describeModel`
- `checkWorkbookModelDescription`
- `isWorkbookModelDescription`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `describeRunResult`
- `checkWorkbookRunResultDescription`
- `isWorkbookRunResultDescription`
- `describeRuntimeRequirements`
- `checkRuntimeRequirements`
- `checkRuntimeAdapter`
- `toPlanData`
- `isPlanData`
- `checkPlanData`
- `hydratePlanData`
- `verifyPlan`
- `verifyPlanData`
- `verifyModel`
- `checkInput`
- `workbookActionInputDescriptionKinds`
- `isWorkbookActionInputDescriptionKind`
- `isWorkbookActionInputDescription`
- `isWorkbookActionInput`
- `builtInWorkbookCheckKinds`
- `isBuiltInWorkbookCheckKind`
- `runWorkbookPlan`
- `runWorkbookAction`
- `workbookPlanId`
- `workbookActionCommandDigest`
- `verifyWorkbookReadbacks`
- `checkWorkbookReadbackProof`
- `isWorkbookReadbackProof`
- `normalizeWorkbookActionInputDescription`
- `workbookRunErrorCodes`
- `isWorkbookRunErrorCode`
- `workbookRuntimeRequirementKinds`
- `isWorkbookRuntimeRequirementKind`
- `workbookRuntimeCapabilities`
- `isWorkbookRuntimeCapability`
- `checkWorkbookCommandRequest`
- `normalizeWorkbookCommandRequest`
- `isWorkbookCommandRequest`
- `checkWorkbookCommandBundle`
- `normalizeWorkbookCommandBundle`
- `isWorkbookCommandBundle`
- `workbookCommandResultFor`
- `workbookCommandResultForReceipts`
- `workbookOpCommandReceiptIdentity`
- `workbookOpCommandReceipt`
- `workbookOpCommandFeatureId`
- `checkWorkbookCommandResult`
- `checkWorkbookCommandResultForBundle`
- `normalizeWorkbookCommandResult`
- `isWorkbookCommandResult`
- `isWorkbookCommandResultForBundle`
- `checkWorkbookCommandReceipt`
- `normalizeWorkbookCommandReceipt`
- `isWorkbookCommandReceipt`
- `workbookCommandReceiptOpsMatch`
- `workbookCommandCategories`
- `isWorkbookCommandCategory`
- `workbookCommandExecutionModes`
- `isWorkbookCommandExecutionMode`
- `workbookCommandReceiptStatuses`
- `isWorkbookCommandReceiptStatus`
- `workbookCommandBundleCommandKinds`
- `isWorkbookCommandBundleCommandKind`
- `workbookJsonSchemaVersion`, `workbookJsonSchemaNames`, `workbookJsonSchemas`, `workbookJsonSchemaHashes`, `workbookJsonSchemaBundleHash`, and `workbookJsonSchemaHash`
- `formula`
- `workbook.addOp(op, { target?, message? })` inside model actions
- `findTable`, `findColumn`, `findRange`, `findName`, and `findRows` through the model workbook context and as top-level helpers
- `check.exists`, `check.noFormulaErrors`, `check.valueEquals`, `check.formulaEquals`, and `check.custom` through the model workbook context and as top-level helpers
- `WorkbookModel`, `WorkbookAction`, `WorkbookActionConfig`, `WorkbookActionDefinition`, `WorkbookActionContext`, `WorkbookCheckContext`, `WorkbookFindWorkbook`, `WorkbookCheckWorkbook`, `WorkbookActionWorkbook`, `WorkbookModelWorkbook`, `WorkbookFindNamespace`, `WorkbookRef`, `WorkbookRefData`, `WorkbookRefKind`, `WorkbookRefIssueCode`, `WorkbookRefIssue`, `WorkbookRefCheckResult`, `WorkbookRefDataIssueCode`, `WorkbookRefDataIssue`, `WorkbookRefDataCheckResult`, `WorkbookBaseRefData`, `WorkbookRangeRef`, `WorkbookRangeRefData`, `WorkbookNameRef`, `WorkbookNameRefData`, `WorkbookTableRef`, `WorkbookTableRefData`, `WorkbookColumnRef`, `WorkbookColumnRefData`, `WorkbookRowsRef`, `WorkbookRowsRefData`, `WorkbookRowOperator`, `WorkbookRowValueType`, `WorkbookActionInput`, `WorkbookActionInputDescription`, `WorkbookActionInputDescriptionKind`, `WorkbookActionInputIssueCode`, `WorkbookActionInputIssue`, `WorkbookActionInputCheckResult`, `WorkbookActionInspection`, `WorkbookAddOpOptions`, `WorkbookActionPlanResult`, `WorkbookActionPreparation`, `WorkbookModelDescription`, `WorkbookModelDescriptionIssueCode`, `WorkbookModelDescriptionIssue`, `WorkbookModelDescriptionCheckResult`, `WorkbookRefDescription`, `WorkbookActionPlanDescription`, `WorkbookPlanData`, `WorkbookPlanId`, `WorkbookPlanDataRefs`, `WorkbookPlanDataIssueCode`, `WorkbookPlanDataIssue`, `WorkbookPlanDataCheckResult`, `WorkbookExecutablePlan`, `WorkbookActionPlanResultDescription`, `WorkbookRunResultDescription`, `WorkbookUndoRefDescription`, `WorkbookRunApplySummaryDescription`, `WorkbookRunUnverifiedDescription`, `WorkbookRuntimeRequirements`, `WorkbookRuntimeRequirement`, `WorkbookRuntimeRequirementKind`, `WorkbookRuntimeRequirementsIssueCode`, `WorkbookRuntimeRequirementsIssue`, `WorkbookRuntimeRequirementsCheckResult`, `WorkbookRuntimeCapability`, `WorkbookRuntimeAdapterIssueCode`, `WorkbookRuntimeAdapterMethod`, `WorkbookRuntimeAdapterIssue`, `WorkbookRuntimeAdapterCheckResult`, `WorkbookRuntimeAdapterCandidate`, `WorkbookJsonSchema`, `WorkbookJsonSchemaName`, `WorkbookJsonSchemaValue`, `WorkbookJsonSchemaScalar`, `WorkbookPlanVerification`, `WorkbookPlanIssue`, `WorkbookModelVerification`, `WorkbookModelActionVerification`, `WorkbookModelVerificationOptions`, `WorkbookRunAdapter`, `WorkbookRunApplyResult`, `WorkbookRunOptions`, `WorkbookRunApplySummary`, `WorkbookCommandResolvedRefs`, `WorkbookResolvedRefData`, `WorkbookResolvedRefValue`, `WorkbookRunUnverified`, `WorkbookRunUnverifiedKind`, `WorkbookRunReadback`, `WorkbookReadbackProof`, `WorkbookReadbackProofCheckResult`, `WorkbookReadbackVerification`, `WorkbookReadbackIssue`, `WorkbookReadbackIssueCode`, `WorkbookCheckExpectation`, `WorkbookCheckExpectationDescription`, `WorkbookBuiltInCheckKind`, `WorkbookCustomCheckOptions`, `WorkbookReadbackCheckOptions`, `WorkbookFormulaExpression`, `WorkbookFormulaLabel`, `WorkbookFormulaLabelDescription`, `WorkbookRawFormulaOptions`, `WorkbookRunResult`, `WorkbookRunError`, `WorkbookRunErrorCode`, and `WorkbookCheckResult`
- `WorkbookFeatureId`, `WorkbookCommandCategory`, `WorkbookCommandExecutionMode`, `WorkbookCommandReceiptStatus`, `WorkbookCommandBundleCommandKind`, `WorkbookCommandResultStatus`, `WorkbookOpCommandReceiptIdentity`, `WorkbookOpCommandReceiptOptions`, `WorkbookCommandRequest`, `WorkbookCommandRequestIssueCode`, `WorkbookCommandRequestIssue`, `WorkbookCommandRequestCheckResult`, `WorkbookCommandBundleScope`, `WorkbookCommandBundleCommand`, `WorkbookCommandBundle`, `WorkbookCommandResult`, `WorkbookCommandResultIssueCode`, `WorkbookCommandResultIssue`, `WorkbookCommandResultCheckResult`, `WorkbookCommandBundleIssueCode`, `WorkbookCommandBundleIssue`, `WorkbookCommandBundleCheckResult`, `WorkbookCommandReceipt`, `WorkbookCommandReceiptIssueCode`, `WorkbookCommandReceiptIssue`, and `WorkbookCommandReceiptCheckResult`
- the advanced `@bilig/workbook/features` subpath additionally exports `defineWorkbookFeaturePlugin`, `checkWorkbookFeaturePlugin`, `workbookProjectionInterceptorPoints`, `isWorkbookProjectionInterceptorPoint`, `workbookUiContributionSlots`, `isWorkbookUiContributionSlot`, `WorkbookFeatureLifecycleContext`, `WorkbookCommandDescriptor`, `WorkbookProjectionInterceptorPoint`, `WorkbookUiContributionSlot`, `WorkbookFeaturePluginIssueCode`, `WorkbookFeaturePluginIssue`, `WorkbookFeaturePluginCheckResult`, `WorkbookCellDisplayProjection`, `WorkbookCellStyleProjection`, `WorkbookRangeChromeProjection`, `WorkbookRowVisibilityProjection`, `WorkbookCommandMetadataProjection`, `WorkbookProjectionContext`, `WorkbookProjectionInterceptorRegistration`, `WorkbookUiContribution`, `WorkbookFeatureRegistration`, and `WorkbookFeaturePlugin`
- the existing low-level operation language: `WorkbookOp`, `WorkbookTxn`, `EngineOp`, and `EngineOpBatch`

The package builds portable workbook intent and concrete low-level ops when the
target is already known. Formula helpers use `@bilig/formula` for parsing and
normalization. Actual calculation and authoritative execution stay in
`@bilig/core` and `apps/bilig`.

`defineModel` returns frozen, normalized model metadata. Model and action names
must be non-empty and already trimmed, while descriptions and input metadata are
trimmed into model-owned frozen copies so the manifest an agent inspected cannot
be mutated later by the caller. The original consumer config remains
caller-owned data; `defineModel` does not freeze or rewrite it.
Action-object manifests only read own `run`, `description`, and `input`
properties. Prototype-inherited metadata is ignored, and an inherited `run`
function is rejected, so agent-visible manifests stay plain and explicit.
Accessor-backed model config, action-map entries, and action-object metadata are
rejected without invoking hidden getters.

Model actions can accept plain JSON-safe input through
`planWorkbookAction(model, actionName, input)` and
`buildWorkbookActionPlan(model, actionName, input)`. Agents that want the whole
pre-runtime handoff should use `prepareWorkbookAction(model, actionName, input)`
instead of manually stitching together planning, verification, plan-data, plan-id,
and runtime-requirement calls. The input is cloned,
canonicalized with stable object-key order, recorded on the plan, and exposed to
the action context as `input`. Supported values are strings, finite numbers,
booleans, `null`, arrays without holes, and plain objects. `@bilig/workbook`
does not hardcode consumer domains, but it does provide a small generic input
description language so agents can form safe tool calls without importing the
consumer's private code. `checkInput(description, value)` checks that generic
metadata and returns a plain `{ status, input, issues }` result.
Action names are checked as data before model lookup. Non-string, empty, or
whitespace-padded runtime action names fail with `invalid_action_name` at
`path: "actionName"` and are not coerced through caller-owned `toString`,
`valueOf`, or `Symbol.toPrimitive`.
Omitted input is valid unless the top-level description sets `required: true`;
required omissions return `missing_required_input` instead of pretending
`undefined` is a malformed JSON payload. `planWorkbookAction` runs that same
check before `find`, `checks`, or action code when an action declares input
metadata. Failed action input checks keep the stable input issue `path` and
`issueCode` on each run error, so agents can branch on structured diagnostics
without parsing messages. JSON-safety failures preserve nested paths like
`input.items[2].amount` instead of collapsing every issue to `input`.
Action helper calls also validate their own output boundary during planning:
write, clear, format, and low-level-op helpers reject malformed targets,
non-literal values, invalid format/add-op options, class/custom-prototype option
roots, and accessor-backed op payloads before a plan is returned.
Check helper calls validate the same way: bad targets, readback options, custom
check options, class/custom-prototype option roots, sparse ref arrays, and
accessor-backed check payloads fail before the check is recorded.
Formula helpers reject sparse, accessor-backed, or class/custom-prototype raw
options, explicit input arrays, label arrays, and function argument arrays before
returning formula intent.
Ref transport helpers return frozen plain data and frozen arrays, so inspected
refs stay stable across agent handoff and verification. Transported ref nodes
and nested selector records must be object-record data, not
class/custom-prototype instances.
`verifyPlan` treats public plan handoff as data too: malformed, sparse, or
accessor-backed plan objects return an `invalid_plan` issue instead of executing
hidden properties or throwing at the caller.
Normalized payloads preserve consumer-owned JSON keys as own data properties,
including names like `__proto__` and `constructor`, so transported tool payloads
cannot mutate prototypes or disappear during canonicalization.
Action input payloads and input-description metadata must be enumerable own data
properties. Accessors are rejected without invocation, which keeps tool payload
validation inspectable and prevents hidden consumer code from running during
planning.
Feature extension metadata follows the same plain-data rule.
`defineWorkbookFeaturePlugin`, `checkWorkbookFeaturePlugin`, and
`normalizeWorkbookCommandDescriptor` reject custom-prototype records and
accessor-backed command descriptor fields before feature manifests reach runtime
registration.
`verifyModel(model, { inputs })` supplies per-action inputs for
whole-model verification. Invalid, array-backed, or accessor-backed model
manifests return an `invalid_model` verdict with no actions instead of
throwing, so agents can audit unknown model objects without wrapping the
verifier in their own try/catch.
Verification options use the same data boundary: `inputs` and each action input
must be object-record roots with own data properties. Accessor-backed or
class/custom-prototype verification inputs return structured
`invalid_action_input` results without invoking getters.
The verdict containers and generated issue arrays from `checkInput`,
`checkPlanData`, `verifyPlan`, and `verifyModel` are frozen, so an agent can
inspect them once and pass them onward without later mutation changing the
meaning.
The frozen
`workbookActionInputDescriptionKinds` list plus
`isWorkbookActionInputDescriptionKind(value)`,
`isWorkbookActionInputDescription(value)`, and `isWorkbookActionInput(value)` let
agents validate metadata and JSON-safe tool payloads without importing a schema
framework or string-matching ad hoc kinds.
The published JSON schemas are expected to reject the same representative
invalid payloads as the TypeScript checkers, including row predicate value types
and low-level op commands that omit explicit destructive confirmation.

When agents need to know what an action expects before running workbook code,
actions may be declared as action objects:

```ts
actions: {
  write: {
    description: 'Write a consumer-provided value',
    input: {
      kind: 'object',
      fields: {
        value: { kind: 'number', required: true },
      },
    },
    run({ refs, workbook, input }) {
      if (typeof input !== 'object' || input === null || Array.isArray(input)) {
        throw new Error('input object required')
      }
      const value = input.value
      if (typeof value !== 'number') {
        throw new Error('numeric value required')
      }
      workbook.writeValue(refs.output, value)
    },
  },
}
```

This is descriptive metadata, not a schema framework. Input descriptions use
plain JSON kinds: `json`, `object`, `array`, `string`, `number`, `boolean`, and
`null`. `object` descriptions may list `fields` and set
`additionalProperties: false`; `array` descriptions may list `items`,
`minItems`, and `maxItems`; `number` descriptions may set `min` and `max`;
`string` descriptions may set `minLength`, `maxLength`, and `pattern`; any
description can mark itself `required`, list allowed `values`, and publish a
checked `default` plus `examples`. These names stay intentionally boring so
consumer-defined models can describe parameters without pulling in `zod`,
`effect`, or a Bilig business model package. By default `checkInput` allows
unknown object fields; set `additionalProperties: false` when the action wants a
closed input shape.
`normalizeWorkbookActionInputDescription` trims text, rejects malformed metadata,
freezes the result, and keeps `@bilig/workbook` independent from `zod`, `effect`,
and model-specific validators.

Known single-cell `workbook.format(ref, { numberFormat })` actions compile to
concrete `setCellFormat` ops, including `numberFormat: null` for explicit
format clears. Style patches remain high-level intent until the runtime resolves
style ids.
When a consumer needs a workbook operation that is already covered by the
low-level operation language, model actions can call
`workbook.addOp(op, { target?, message? })`. The op is guarded with
`isWorkbookOp`, cloned into `plan.ops`, and kept as a command so agents can
inspect the handoff without pulling in `@bilig/core`. When a `target` is
supplied for an address or range op, `verifyPlan` checks that the op touches the
same range. For op kinds without an inferable range, `target` is descriptive for
logs and approvals rather than proof of affected cells.
Raw `WorkbookOp` and `EngineOpBatch` guards trust own fields only, including
nested cell ranges and batch clocks, so a runtime can reject prototype-shaped
payloads before hydration or execution. They also reject accessor-backed
required fields, nested fields, and op-array entries without running getters.

Formula helpers keep referenced workbook inputs and formula labels separate from
formula text. Planned `writeFormula` commands expose `inputs` plus `labels`,
where each label maps a token in the formula string to the workbook ref it
represents. Agents can inspect and verify that handoff without relying on human
UI coordinates, hidden JS helper calls, or reverse-parsing placeholder names.
For formulas outside the small helper set, `formula.raw(source, { inputs })`
keeps arbitrary formula text generic while preserving explicit workbook
dependencies for inspection and verification. Use
`formula.raw(source, { inputs, labels })` when the raw formula uses custom ref
tokens. These are declared dependencies and token mappings, not a
parser-discovered model.
Plan verification parses formula text and validates label usage against formula
reference tokens, not substring matches, so labels inside longer identifiers or
quoted string literals do not count as dependency proof.
Formula operands intentionally reject bare strings. Consumers use `formula.raw`
for formula source and `formula.text` for spreadsheet string literals, keeping
agent-authored formulas explicit instead of overloading a string as either code,
a label, a named range, or user text.

Action plans also freeze the consumer-defined `refs` container graph and expose
`refsUsed`, a flat deduped list of workbook refs found inside that object. This
keeps custom models generic while still letting agents inspect what the model
resolved.
The same generic refs are available outside model callbacks through top-level
`findTable`, `findColumn`, `findRange`, `findName`, and `findRows` helpers, or
through the frozen `find` namespace with short aliases such as
`find.table(...)`, `find.range(...)`, and `find.rows(...)`.
The frozen `workbookRefKinds`, `workbookRowOperators`, and
`workbookRowOperatorValueTypes` data, plus `isWorkbookRefKind`,
`isWorkbookRef`, `isWorkbookRowOperator`, and `isWorkbookRowValueCompatible`,
expose the same selector contract as data so agent tools can validate refs and
row predicates without copying string unions or pulling in a schema framework.
These selector helpers trim text, canonicalize cell addresses, and reject empty
or malformed selectors before the runtime handoff. That keeps bad agent intent
out of the plan instead of letting an invalid address, blank column, duplicate
header, invalid row operator, non-finite predicate value, or invalid
operator/value pair fail later inside an engine adapter.
The same operator/value compatibility is enforced for transported row refs, so
`isWorkbookRefData`, `collectWorkbookRefData`, `toWorkbookRefData`, and
`hydrateWorkbookRef` do not accept row predicates that `findRows` would reject.
Table header selectors are an all-of match, not a coordinate or position
contract. Headers are case-sensitive after trimming, stored in sorted order, and
deduped so agents get stable ids for the same generic table intent.
Row selectors keep the value contract simple: `eq` and `neq` accept any JSON
literal, `contains` and `startsWith` accept strings, and ordered comparisons
accept numbers or strings.
`findRows` refs include their predicate value in the stable id, so distinct
consumer-defined row predicates remain distinct during agent inspection and
dedupe. Labels stay simple and readable for logs.
Refs are frozen data objects. Helper methods such as `table.column()` and
`rows.column()` remain available for ergonomics, but they are non-enumerable so
object-key inspection and JSON descriptions stay data-first.
For table-backed row selectors, `rows.column("Amount")` targets that column only
inside the matching rows. Core/app runtime adapters can resolve those generic
refs into exact cells for writes, formats, clears, checks, and row-wise formula
input alignment without adding hardcoded workbook models.
The same planned checks are available outside model callbacks through top-level
`check.exists(ref)`, `check.noFormulaErrors(ref)`,
`check.valueEquals(ref, value)`, `check.formulaEquals(ref, formula)`, and
`check.custom({ kind, message, target, refs })` helpers. Custom checks let
consumers carry their own invariants without adding hardcoded business models to
the package. `target` names the main ref, and `refs` names supporting refs so
agents can describe and verify the full invariant contract.
Custom check kinds cannot reuse built-in names. Tool builders can inspect the
frozen `builtInWorkbookCheckKinds` list or call `isBuiltInWorkbookCheckKind`
before accepting consumer-defined check metadata.
Readback checks add machine-readable expectations to the same generic check
channel: `valueEquals` stores the expected literal value, and `formulaEquals`
stores normalized formula text plus explicit formula input refs and labels.
Runtime code can evaluate those expectations after applying the plan, while
agents can inspect the proof target without relying on visual spreadsheet state.

Model callback phases expose only the API needed for that phase. `find` gets
find helpers only, `checks` gets find helpers and planned-check helpers, and
actions get find helpers, checks, and mutation-planning methods. Discovery and
proof declaration cannot accidentally write workbook intent before the action
phase.
Checks returned from a consumer `checks()` callback are also treated as a data
boundary. Returned arrays must contain own enumerable data entries, and returned
check fields must be own data properties; sparse arrays and accessor-backed
checks fail planning without invoking hidden getters.

`describeModel` returns a JSON-safe model manifest with the model name, optional
model description, sorted action names, per-action descriptions, optional input
descriptions, and whether model-level checks exist. It does not run `find`,
checks, or actions. `checkWorkbookModelDescription(data)` validates the
transported manifest, rejects sparse/accessor-backed entries, verifies that
`actionDetails` exactly match `actions`, and normalizes constrained input
descriptions with the same rules as `checkInput`.
Model inspection reads own data properties only. Model roots must be object
records, not arrays or class/custom-prototype objects with attached fields.
Action maps and action objects may have prototypes, but only own action entries
and own action metadata are read; accessor-backed model names, descriptions,
action maps, and action metadata are rejected without invoking hidden getters.
For agent logs, approvals, tests, and runtime handoff, `describeRef` and
`describePlan` produce JSON-safe descriptions of refs and action plans. The
descriptions preserve generic action input and workbook intent while removing
consumer-private `refs` object shape and helper methods.
Description outputs are frozen all the way down, including nested refs, command
descriptions, check proof, apply proof, undo ops, and error arrays, so an agent
can inspect them once without later caller-side mutation changing the same
object.
Those descriptions are real transport data, not screenshots or UI state:
`isWorkbookRefData` validates ref payloads, `hydrateWorkbookRefs` restores
local helper methods after JSON transport, and `toPlanData` makes a full plan
JSON-safe for handoff. `isPlanData` validates that payload, `checkPlanData`
explains invalid payloads with stable paths, `hydratePlanData` restores frozen
refs and helpers, and `verifyPlanData` verifies transported plan data using only
`refsUsed`, commands, ops, changes, and checks. Invalid transported action input
and check proof keep nested JSON paths such as `input.rows[1]` and
`checks[0].proof.when`. Plan-data validation treats transport payloads as own
data only, so prototype-inherited fields cannot make an invalid payload look
valid. Valid `checkPlanData` results canonicalize the plan before returning it,
so extra enumerable scratch fields are ignored consistently by ids, hydration,
requirements, and execution.
Plan roots and nested plan entries are record-shaped payloads, not arrays with
attached fields. Runtime evidence arrays are still arrays; plan objects are not.
Plans are frozen handoff objects: action input, refs, refs used, commands,
concrete ops, changed summaries, and checks cannot be rewritten after planning.
That lets an agent inspect a plan once and pass the same intent to an adapter
without caller-side metadata drift.
`checkWorkbookRef(ref)` validates live ergonomic refs before transport, including
the helper functions that make table and row refs usable. `checkWorkbookRefData(data)`
validates transported range/name/table/column/rows refs with stable path issues
before hydration or persisted proof use. Ref data roots, nested ranges, row
predicates, nested table refs, and nested rows refs must be object-record data.
`isWorkbookRef` and `isWorkbookRefData` are the boolean guards over those same
boundaries.
Planning validates model manifest data before reading action metadata or running
model code. Accessor-backed model names, action maps, action entries, or action
metadata return structured `invalid_model` errors without invoking hidden
getters.
`inspectModel` and `planWorkbookAction` return frozen wrappers too, so action
manifests, planned results, and failure results have the same inspect-once
behavior as the nested plan objects.
`describePlanResult` applies the same description layer to either planned or
failed action planning results.
`describeRunResult` applies the same JSON-safe description layer after
execution, preserving `done`/`failed` status, changed summaries, checks, errors,
apply proof, command receipts, command-bound `noop` proof, undo ops, and unverified proof notes while
removing ref helper functions from the public result. The returned run
description is frozen before it crosses the agent boundary.
`checkWorkbookRunResultDescription` validates persisted run-result descriptions
with stable path issues before a sync event or later agent treats them as proof.
For persisted command no-op proof, the checker requires empty
`previewOps`/`appliedOps`, matching receipt `commandKind` and `commandDigest`,
matching `effect.kind`, and full `effect.op` data for low-level `op` commands.
`isWorkbookRunResultDescription` is the boolean guard over the same boundary.
The published run-result JSON schema mirrors those proof fields, so
non-TypeScript agents do not have to treat apply or undo proof as opaque blobs.
`describeRuntimeRequirements(plan)` gives agents a JSON-safe adapter checklist
for the same plan or transported plan data: which generic commands must be applied, which readbacks are
needed, and which checks need proof. It stays generic, with capabilities such
as `writeFormula`, `writeValue`, `format`, `clear`, `applyOp`, `read`, and
`verifyCheck`. Command-derived concrete single-cell ops are not repeated as
extra `applyOp` requirements, while explicit or manually assembled ops still
appear as `applyOp`. It does not import the engine.
`checkRuntimeRequirements(data)` validates that checklist after JSON transport
and returns boring `{ status, issues }` diagnostics with paths such as
`requirements[0].capability` or `requirements[2].refs[0]`. That lets agents
reject malformed adapter handoff data before checking runtime methods or
starting mutation. Runtime requirement validation also ignores inherited fields;
the adapter checklist has to be present as explicit payload data. Requirement
arrays and nested ref arrays are data-only too: holes, non-enumerable entries,
or accessor-backed entries are rejected without invoking getters. Valid runtime
requirements are returned as frozen normalized data with caller-owned extra
fields stripped before an agent or adapter trusts the checklist.
The requirements root object, each requirement entry, and the adapter object
must be record-shaped payloads; arrays with attached properties are rejected
even when those properties look like valid runtime handoff fields.

`verifyPlan` gives agents a runtime-free consistency check before handoff. It
first checks that the plan handoff itself is own data, then flags invalid action
input, unresolved command targets, unresolved formula inputs, missing formula
labels, formula labels that do not point at resolved refs or do not appear in
formula reference tokens, duplicate resolved refs, refs used that are not
discoverable from `refs`, unparsable formulas, and missing concrete ops for
write, clear, and number-format commands whose target is already known as a
single cell. Custom check targets and supporting refs must also resolve through
`refsUsed`. Formula readback expectation inputs and labels must resolve through
`refsUsed`, and
formula expectation text must be parseable.
Checks must start as `planned`; consumer model code cannot mark a check passed
or failed before runtime proof.
Low-level `addOp` commands must contain valid `WorkbookOp` values, must still
appear in `plan.ops`, and must match their declared `target` when the op exposes
a concrete address or range.
`verifyModel` applies the same planning and verification flow to every action
in a consumer-defined model, returning one JSON-safe model-level verdict.
If the model manifest itself is invalid, the verdict is `invalid` with an
`invalid_model` error and an empty action list.
Successfully planned actions include their runtime requirements in the same
result, so an agent can inspect the planned intent, static proof, and adapter
handoff checklist without stitching multiple API calls together.
`verifyWorkbookReadbacks` freezes the returned runtime proof verdict, generated
checks, and readback issues. Formula/value readback proof is normalized before it
is attached to passed checks.
`checkWorkbookReadbackProof(data)` is the transported proof boundary for the
same evidence. It validates own data fields, normalizes readbacks, verifies the
checks, and returns either frozen verified proof or stable issue codes.
Readback proof objects, check objects, expectations, and formula labels must be
record-shaped payloads, not arrays with attached fields.
`checkRuntimeAdapter(planOrRequirements, adapter)` compares that checklist to an
adapter shape and returns a plain valid/invalid result with malformed
requirements or missing capability issues. It accepts a live plan, transported
plan data, or the output of `describeRuntimeRequirements`; invalid transported
plan data returns `invalid_requirements` issues instead of throwing. Its verdict
and issue arrays are frozen.
`workbookPlanId(planOrData)` returns the stable id for the generic plan data an
adapter is asked to apply, so runtime proof can be tied back to the exact model,
action, refs, commands, ops, changes, and checks that were planned. Transported
plan data is canonicalized through `checkPlanData` before hashing.
`workbookActionCommandDigest(command)` returns the stable digest for one planned
high-level command. Apply adapters can return `commandReceipts` with command
index, command kind, command digest, preview ops, applied ops, resolved-ref
proof, and optional `formulaLabels` proof, so an agent can inspect which planned
command produced which materialized operations and which workbook refs the
runtime materialized. `resolvedRefs` are deliberately concrete: a receipt for a
symbolic name/table/column/rows command must bind that planned ref to one or
more concrete range refs before strict proof can pass.
`@bilig/workbook` rejects stale digests, duplicate or missing command indexes,
receipt preview/apply mismatches, receipts whose ops do not match the planned
command's concrete workbook op, and receipts whose flattened ops do not match
the apply-level preview or applied ops. With `requireApplyProof: true`, a plan
with high-level commands fails closed unless those command receipts are present.
With `requireResolvedRefs: true`, a plan with ref-targeting commands fails
closed unless those command receipts also bind planned refs to concrete range
refs.
With `strict: true`, each command receipt must also prove concrete applied ops
or command-bound no-op proof with `source`, `evidence`, `opCount: 0`,
`commandKind`, `commandDigest`, and command-effect `effect` data; ref-targeting
commands must also prove resolved concrete refs, and low-level `op` commands
must prove the full planned op, not only the op kind. Range refs must match the
planned range exactly, and symbolic refs must materialize to concrete ranges
before their ops are accepted. The persisted run-result description checker
keeps that proof inspectable by rejecting no-op descriptions whose proof fields
do not match the enclosing command receipt.
Strict runs also require
mutating plans to declare checks before `adapter.apply` is called, require
`baseRevision` and `revision` on apply proof, fail closed on
`expectedBaseRevision` mismatches, reject unverified apply summaries, and reject
passed checks that do not carry proof.
For generic feature-command results, `@bilig/workbook` also derives the settled
result status, matched flag, changed ranges, and errors from receipts before a
bundle-bound result can pass validation.
`runWorkbookPlan(planOrData, adapter)` and
`runWorkbookAction(model, actionName, adapter, input)` add a transport-neutral
apply-and-prove loop on top of the same contracts. The adapter receives the full
plan, or a hydrated transported plan with `refs: { refsUsed }`. Malformed
transported plan data returns `status: "failed"` with `invalid_plan_data`
errors and never reaches `adapter.apply`. When the requirements include an apply
capability, the adapter applies it through
whatever runtime the consumer owns and may return `previewOps`, `appliedOps`,
apply proof, undo metadata, and semantic readbacks for the expectation targets.
When the plan only requires readbacks or generic check proof, `@bilig/workbook`
skips mutation and runs those proof steps directly. `@bilig/workbook` compares
preview ops to applied ops when both are present, compares readbacks against
`valueEquals` and `formulaEquals` checks, and returns a boring
`WorkbookRunResult`. Passed readback-backed checks include JSON-safe proof such as
`{ source: "readback", value }` or
`{ source: "readback", formula, expectedFormula, materializedFormula }`, so the
result says what the runtime actually read back and how symbolic formula labels
were materialized. Formula readbacks can include `formulaLabels`; those labels
are applied by parsed formula tokens through `@bilig/formula`, not by substring
replacement. If static verification fails,
the apply adapter is not called. If the adapter is missing a required method for
the plan, the apply adapter is not called and the run fails with
`adapter_missing_capability`. If preview/apply ops mismatch, the run fails with
`apply_mismatch`. If `requireApplyProof` is true and the adapter omits preview
or applied ops, omits command receipts for a command-based plan, or returns
command receipts with no concrete applied ops, the run fails with
`apply_not_verified`. In strict mode, missing or stale resolved-ref proof for a
ref-targeting command also fails with `apply_not_verified`. Strict mode and
`requireNoUnverified` reject missing preview/applied op pairs before readback. If
the adapter returns a stale `planId`, the run fails before readback or check
proof; if `requirePlanId` is true and the adapter omits it, the run fails with
`plan_not_verified`. `requireRevision` and strict mode require `baseRevision`
and `revision`; `expectedBaseRevision` rejects stale applies before readback.
`requireCheckProof` and strict mode reject passed checks without proof. Run
options and adapter-conformance options must be plain objects with known keys and
own data properties only; custom-prototype objects, unknown option names, and
accessor-backed options fail with `invalid_run_options`. Adapter methods are
also read as own data properties only; accessor-backed runtime methods are
treated as missing capabilities without invoking getters. If a readback
expectation is missing or mismatched after a reader runs, the run fails with
deterministic codes such as `readback_missing`, `value_mismatch`, or
`formula_mismatch`.
The in-repo `@bilig/core` adapter is the reference execution path for this
contract. It returns plan-bound apply proof, command receipts, concrete applied
ops, base/applied revisions, resolved target/input refs, and core-owned proof for
generic check verification. The monolith app accepts transported
`WorkbookPlanData` through `workbook.applyWorkbookPlanData`, runs it with
`{ strict: true }` and the current expected base revision, persists the original
plan together with the materialized applied ops for replay and the frozen
run-result description for later inspection, and applies the run undo ops if
post-apply proof fails. That keeps `@bilig/workbook` generic while giving agents
a real app-owned execution path for their own models.
Returned run results are frozen before they cross the public API boundary,
including nested changed summaries, checks, errors, apply summaries, undo refs,
and unverified proof notes. That lets an agent inspect a run once and keep the
same proof object for logging, approval, or retry decisions without defensive
copying.
Failures after an adapter reports `status: "applied"` preserve `changed` and
`undo`, so an agent can still inspect what was applied and how to reverse it
when a later proof step rejects the run.
Failed apply results only preserve `changed` when the adapter reports applied
ops or undo metadata; an explicit empty `appliedOps` array with no undo keeps
`changed: []`.
Runtime readbacks must match the requested target set exactly; surplus
readbacks fail with `readback_unexpected`, and duplicate targets fail with
`readback_duplicate`. Formula readbacks are parsed with `@bilig/formula` and
canonicalized into no-leading-`=` proof, so agents do not have to special-case
runtime formatting differences such as a leading equals sign, whitespace, or
redundant parentheses.
Readback and check proof objects must be record-shaped data; arrays with
attached `target`, `value`, `formula`, or check fields are rejected before they
can become proof.
Run errors use the stable `WorkbookRunErrorCode` union. Agents and adapters can
inspect the frozen `workbookRunErrorCodes` list or call
`isWorkbookRunErrorCode(value)` before branching on a code. Runtime adapters
should use `apply_failed` for apply exceptions and `runtime_rejected` for
intentional runtime refusal with a specific message instead of inventing
model-specific public error codes. Errors may also include a stable `path` and
`issueCode` when a generic checker can identify the exact rejected input or
adapter issue.
`adapter.apply` only applies the plan and may return apply proof plus an undo
ref; it cannot drop, replace, or prove checks. Returning `status: "applied"`
with non-empty `errors` is rejected as `runtime_rejected`. If no apply proof is
required and the adapter omits preview or applied ops, the run can still finish
but reports an `unverified` apply fact instead of pretending preview/apply match
was proven.
Apply result fields, nested runtime errors, and undo metadata are sanitized from
own fields before they reach `WorkbookRunResult`. Accessor-backed preview ops,
applied ops, undo ops, runtime errors, and verifier proof are rejected without
invoking getters.
If an adapter returns `status: "failed"` with `appliedOps` or `undo`, the failed
run preserves the planned change summaries and undo ref because the runtime is
signaling that mutation evidence exists even though the apply step rejected the
overall run.
Adapters can also expose `verifyChecks(checks, plan)` for generic proof of
non-readback checks such as existence checks, formula-error checks, and
consumer-defined invariants. `verifyChecks` returns the same checks in the same
order and may only change `status` or add JSON-safe `proof`. Malformed output,
contract changes, invalid proof, and unsupported verifier mutations fail with
`invalid_check_verification`; thrown verifier errors fail with
`check_verification_failed`, and failed checks become `check_failed` run errors.
Accepted verifier output is sanitized before it reaches `WorkbookRunResult`. If
a check remains `planned` after readback and adapter verification,
`runWorkbookPlan` returns `failed` with `check_not_verified`; `status: "done"`
does not hide unproven checks.
Verifier output is also own-field only: inherited `status`, `kind`, `message`,
target, refs, expectation, or proof data is ignored and cannot prove a check.
`@bilig/core` provides `createWorkbookRunAdapter(engine)` for the canonical
engine handoff. It materializes generic `plan.commands` into engine operations,
including range and table-column writes, falls back to explicit `plan.ops` for
low-level plans, reads single-cell `valueEquals` and `formulaEquals` targets,
and verifies generic `exists` and `noFormulaErrors` checks. When the engine
applies a plan, the adapter returns matching `previewOps` and `appliedOps`, plus
the stable plan id, per-command receipts, resolved-ref proof, and JSON-safe
apply proof. When the
engine captures an undo transaction, the adapter
returns a portable `undo.ops` ref using the same workbook operation language.
Consumer-defined business meaning stays in the model; the core adapter only
proves generic workbook facts.

## Core engine surface

The canonical engine surface includes:

- `createWorkbookRunAdapter`
- `createSheet`
- `deleteSheet`
- `setCellValue`
- `setCellFormula`
- `setCellFormat`
- `clearCell`
- `setRangeValues`
- `setRangeFormulas`
- `clearRange`
- `fillRange`
- `copyRange`
- `pasteRange`
- `setSelection`
- `undo`
- `redo`
- `getCell`
- `getDependencies`
- `getDependents`
- `explainCell`
- `exportSnapshot`
- `importSnapshot`
- `exportReplicaSnapshot`
- `importReplicaSnapshot`
- `applyRemoteBatch`
- `subscribe`
- `subscribeBatches`
- `connectSyncClient`
- `disconnectSyncClient`
- `getSyncState`

## WorkPaper surface

`@bilig/workpaper` exposes `WorkPaper`, a HyperFormula-style headless workbook API on top
of `@bilig/core`:

- `WorkPaper.buildEmpty`
- `WorkPaper.buildFromArray`
- `WorkPaper.buildFromSheets`
- `WorkPaper.buildFromSnapshot`
- workbook reads for cells, ranges, sheets, and named expressions
- display-value and formula-diagnostic reads for user-facing errors and
  structured financial validation details
- workbook mutations for cells, rows, columns, sheets, clipboard, and history
- `batch`, `suspendEvaluation`, `resumeEvaluation`, `undo`, and `redo`
- formula helpers such as address parsing, normalization, validation, and scratch evaluation
- static language and custom-function registration
- HyperFormula-style positional events through `on`, `once`, and `off`
- additive structured events through `onDetailed`, `onceDetailed`, and `offDetailed`
- stable internal adapter getters: `graph`, `rangeMapping`, `arrayMapping`,
  `sheetMapping`, `addressMapping`, `dependencyGraph`, `evaluator`,
  `columnSearch`, and `lazilyTransformingAstService`

`WorkPaper` is the canonical top-level contract.

## Excel Import Surface

`@bilig/workpaper/xlsx` exposes the CSV/XLSX boundary for WorkPaper consumers:

- `importXlsx(bytes, fileName)`
- `importXlsxFile(path, fileName?, options?)`
- `importCsv(text, fileName, options?)`
- `importWorkbookFile(bytes, fileName, contentType, options?)`
- `exportXlsx(snapshot)`
- `exportWorkPaperXlsxToFileAsync(workbook, outputPath)`
- `CSV_CONTENT_TYPE`
- `XLSX_CONTENT_TYPE`

CSV import auto-detects comma, semicolon, and tab delimiters. For locale-specific
accounting exports, pass `{ delimiter: ";", decimalSeparator: "," }`.

```sh
pnpm add @bilig/workpaper
```

```ts
import { WorkPaper } from '@bilig/workpaper'
import { exportWorkPaperXlsxToFileAsync, importXlsxFile } from '@bilig/workpaper/xlsx'

const imported = importXlsxFile('model.xlsx')
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  useColumnIndex: true,
})

const firstSheetName = imported.snapshot.sheets[0]?.name
const firstSheet = firstSheetName === undefined ? undefined : workbook.getSheetId(firstSheetName)
if (firstSheet === undefined) throw new Error('Workbook has no sheets')

workbook.setCellContents({ sheet: firstSheet, row: 1, col: 1 }, 150_000)
await exportWorkPaperXlsxToFileAsync(workbook, 'model-edited.xlsx')
workbook.dispose()
```

Use `importXlsxFile(...).snapshot` with `WorkPaper.buildFromSnapshot()` when a
large XLSX already lives on disk and a consumer needs Excel workbook metadata
such as defined names, tables, and translated structured references.
`WorkPaper.buildFromSheets()` remains a metadata-free array/sheet constructor.
For large imported workbooks with scalar literal edits,
`exportWorkPaperXlsxToFileAsync()` can patch the original XLSX package from its
preserved source reader and write to disk without forcing a full in-memory
export snapshot. Use `importXlsx(bytes, fileName)` for already-materialized
bytes, and use `workbook.exportSnapshot()` with `exportXlsx()` for generated
workbooks or edits that cannot be represented as source-preserving scalar
patches.

### Core types added in the current tranche

- `CellRangeRef`
- `SelectionRange`
- `SelectionEditMode`
- `SyncState`

## Binary protocol

`@bilig/binary-protocol` exposes:

- `PROTOCOL_VERSION`
- `encodeFrame(frame): Uint8Array`
- `decodeFrame(bytes): ProtocolFrame`

The current frame families are:

- `hello`
- `appendBatch`
- `ack`
- `snapshotChunk`
- `cursorWatermark`
- `heartbeat`
- `error`

The protocol surface is already real for sync and snapshot traffic. The remaining architecture gap is not “do we have binary frames?” but “does the authoritative workbook mutation language fully match what the local engine can represent?”

## Worker transport

`@bilig/worker-transport` exposes:

- `createWorkerEngineHost(engine, port)`
- `createWorkerEngineClient({ port })`

The first tranche already supports:

- method invocation
- engine events
- outbound batch subscriptions

## Target state

- the core API stays source-compatible while the worker-first runtime and durable backend land under it
- the binary protocol becomes the canonical wire format for browser sync, backend relay, and agent APIs
- the agent API moves from JSON payload bodies to typed binary request/response/event frames

## Exit gate

- every API documented here exists in code
- typecheck passes across all packages that import these APIs
- direct engine tests cover range mutation, history, selection state, and sync-state behavior
- the docs no longer claim a stable interface that is missing in the repo

## Agent API

`@bilig/agent-api` exposes:

- `AgentRequest`
- `AgentResponse`
- `AgentEvent`
- `encodeAgentFrame`
- `decodeAgentFrame`
- `encodeStdioMessage`
- `decodeStdioMessages`

Today the package defines typed TypeScript request/response/event unions, but the transport payload is still serialized with `JSON.stringify(...)` inside the binary frame envelope. The target state is a true typed binary schema shared by stdio and remote network usage.
