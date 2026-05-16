# Changelog

All notable changes to `@bilig/headless` will be documented in this file.

This package is released as part of the aligned bilig library package set.

## 0.16.6

- Release type: patch
- Previous libraries tag: libraries-v0.16.5
- Manual override: no

## Fixes

- fix(docs): sync public evidence for headless 0.16.5 (fcdc237b)
- perf(core): reuse rectangular aggregate templates (48c96ac2)
- perf(headless): batch serialized formula paste (bad422dd)
- fix(grid): stabilize table clears and metadata (86c5747e)
- fix(headless): support Node 22 runtime installs (3ef59ac4)
- perf(headless): preserve direct scalar formula bindings (dfc18a00)
- perf(headless): reduce matrix plan allocation (a252c443)
- fix(headless): validate workpaper mcp cli args (7f7a4300)
- perf(core): fast-translate aggregate templates (c27455c2)
- perf(core): reduce region subscription key churn (3c97533e)
- fix(test): share strict bench tolerance parsing (120ad96b)
- fix(smoke): validate workpaper stage flag (9159f541)
- perf(headless): speed tracked formula edits (3694f5da)
- perf(core): reduce kernel sync defer allocation (67956ff1)
- fix(release): validate publish env flags (81dc392f)
- perf(core): reuse direct scalar closure indices (c4ce54d9)

## Internal runtime changes

- docs(growth): add quote approval proof page (42e4444a)
- refactor(core): isolate operation service test hooks (e8d0980c)
- docs(growth): surface npm provenance trust path (2a945bda)
- ci(security): publish openssf scorecard results (f6c80a2b)
- ci(security): add codeql and dependabot (b1c8541f)
- ci(security): constrain workflow token permissions (e830185d)
- ci(security): pin workflow and image dependencies (42b111ff)
- docs(growth): add proof-time bookmark path (234117c7)
- docs(mcp): add runnable stdio transcript smoke (e43dffff)

## 0.16.5

- Release type: patch
- Previous libraries tag: libraries-v0.16.4
- Manual override: no

## Fixes

- fix(docs): refresh headless performance evidence (81271868)
- perf(headless): streamline matrix dimension updates (91f9c619)

## Internal runtime changes

- chore(headless): refresh package footprint (71ce5c20)
- docs(trust): surface security and support policies (970a8a7e)
- docs(adoption): add production readiness checklist (29df03fd)

## 0.16.4

- Release type: patch
- Previous libraries tag: libraries-v0.16.3
- Manual override: no

## Fixes

- fix(docs): sync public evidence for headless 0.16.3 (a2c20561)

## Internal runtime changes

- refactor(formula): isolate workday builtins (6a0d0c33)

## 0.16.3

- Release type: patch
- Previous libraries tag: libraries-v0.16.2
- Manual override: no

## Fixes

- perf(core): streamline direct scalar delta bookkeeping (a8ec9251)
- fix(release): align headless public evidence version (87e8b785)
- fix(grid): sharpen spreadsheet font rendering (77ac870b)
- fix(docs): restore headless footprint artifact (3d0d2c6f)

## Internal runtime changes

- docs(growth): record star spike evidence (943a6aa7)
- refactor(excel-import): isolate cell value parsing (9abfdc8b)

## 0.16.2

- Release type: patch
- Previous libraries tag: libraries-v0.16.1
- Manual override: no

## Fixes

- fix(core): anchor sheet range metadata (077bfbe1)

## 0.16.1

- Release type: patch
- Previous libraries tag: libraries-v0.16.0
- Manual override: no

## Fixes

- perf(headless): template-bind fresh append formulas (b6ec9a28)
- fix(web): drop stale assistant context retries (194eefa9)
- perf(headless): skip unchanged aggregate retargets on tail append (73c1e4e5)
- perf(headless): accelerate cross-sheet direct formulas (808668fe)
- perf(headless): skip column dependency checks without subscribers (700f5835)
- perf(headless): fast path rectangular aggregate clears (03d153dc)
- perf(headless): tighten formula replacement propagation (d10d55dc)
- perf(headless): avoid combined matrix refs (926778d1)
- fix(core): bind fresh formulas with defined names safely (604f0117)
- perf(headless): combine fresh aggregate scans (c2a78b5d)
- fix(ci): build headless before footprint probe (ceefb2dc)
- fix(headless): refresh package footprint evidence (c20bcbf0)
- perf(headless): collapse matrix dimension refreshes (e845484b)
- fix(ci): build runtime types before release metadata push (027c019c)
- perf(headless): preserve scalar formula dimensions (4ce2e20b)
- perf(headless): skip fresh-cell spill cleanup (f6eabba6)
- fix(ci): skip stale runtime release plans (0addaa62)

## Internal runtime changes

- docs(headless): gate cold-start package footprint (18b52b2c)
- docs(headless): refresh package footprint (b0ea64d0)
- refactor: address workbook technical debt (67a282c7)
- refactor(core): split fresh aggregate mutation helpers (98165d20)
- refactor(core): split dynamic scalar binding helpers (91502b78)
- refactor(core): split recalc evaluation state helpers (01887c09)
- refactor(core): split live kernel sync state (fc1aefa3)
- refactor(core): centralize formula binding cell flags (a65ca70a)
- refactor(excel-import): split xlsx style value helpers (29f42415)
- docs(discovery): restore proven headless positioning (34b4c9b4)
- refactor(core): split initial prefix aggregate evaluation (21f0d675)
- refactor(wasm): split statistics rank dispatch (03f4d59e)
- refactor(wasm): split vm output string arena (63c4260c)
- refactor(core): split direct scalar slice tracking (bb560eaa)
- refactor(core): centralize formula binding effect errors (fde61550)
- refactor(core): centralize mutation op records (aa997db6)
- refactor(core): isolate full invalidation emission (d3b16839)
- docs(headless): add formula recalculation discovery pages (03fa7286)
- refactor(wasm): remove duplicate concat writer (b5cf1678)
- docs(headless): add screenshot automation boundary article (6b31a4c8)
- refactor(core): isolate recalc iteration settings (93bd88d4)
- refactor(core): centralize kernel sync literal events (9ecdb56a)
- refactor(core): route workbook protection through metadata service (8059c417)
- refactor(core): isolate direct criteria ast helpers (a8dda3a7)
- refactor(headless): isolate history snapshot cloning (85191d7b)
- refactor(core): isolate recalc event emission (4fd2365a)
- refactor(core): centralize structural axis edits (83a27fd9)

## 0.16.0

- Release type: minor
- Previous libraries tag: libraries-v0.15.1
- Manual override: no

## Features

- feat(examples): add quote approval workpaper api (6b4dd7ea)

## 0.15.1

- Release type: patch
- Previous libraries tag: libraries-v0.15.0
- Manual override: no

## Fixes

- perf(headless): skip fresh append range probes (8db053b1)

## 0.15.0

- Release type: minor
- Previous libraries tag: libraries-v0.14.29
- Manual override: no

## Features

- feat(headless): add file-backed workpaper mcp mode (b51992fb)

## 0.14.29

- Release type: patch
- Previous libraries tag: libraries-v0.14.28
- Manual override: no

## Fixes

- perf(headless): fast path rectangular row sums (fe9a70b1)

## Internal runtime changes

- docs(headless): compress package positioning (f86a6b3e)

## 0.14.28

- Release type: patch
- Previous libraries tag: libraries-v0.14.27
- Manual override: no

## Fixes

- perf(headless): streamline direct aggregate binding (91424682)

## 0.14.27

- Release type: patch
- Previous libraries tag: libraries-v0.14.26
- Manual override: no

## Fixes

- perf(headless): skip independent aggregate topo repair (2e739ebc)

## Internal runtime changes

- docs(growth): refresh conversion snapshot (e9ce18e9)
- docs(growth): hide campaign docs from public path (c2a44b5b)
- docs(evidence): sync public benchmark claims (6eef08a0)

## 0.14.26

- Release type: patch
- Previous libraries tag: libraries-v0.14.25
- Manual override: no

## Fixes

- perf(headless): expand competitive workbook benchmarks (9f63cbc7)

## Internal runtime changes

- ci(release): tolerate github mirror race (204afce8)

## 0.14.25

- Release type: patch
- Previous libraries tag: libraries-v0.14.24
- Manual override: yes

## Internal runtime changes

- docs(readme): tighten headless positioning (09d12302)

## 0.14.24

- Release type: patch
- Previous libraries tag: libraries-v0.14.23
- Manual override: no

## Fixes

- perf(headless): fast path exact index match (7ea2aaf1)

## Internal runtime changes

- docs(site): rebuild landing hero (9ccc2f13)
- test(bench): broaden headless competitive workloads (a9b3d5d8)

## 0.14.23

- Release type: patch
- Previous libraries tag: libraries-v0.14.22
- Manual override: no

## Fixes

- perf(core): cache normalized range lookups (0ca9a7c0)

## 0.14.22

- Release type: patch
- Previous libraries tag: libraries-v0.14.21
- Manual override: no

## Fixes

- perf(headless): compact initial formula load refs (65f3d446)

## 0.14.21

- Release type: patch
- Previous libraries tag: libraries-v0.14.20
- Manual override: no

## Fixes

- perf(workbook): harden headless mutation fast paths (0a806e6c)

## 0.14.20

- Release type: patch
- Previous libraries tag: libraries-v0.14.14
- Manual override: no

## Fixes

- fix(headless): resolve real workbook import regressions (5ca9d46a)
- fix(ci): keep workbook worker release budget green (9f2dbfe5)
- fix(headless): restore dynamic spills from documents (2c57d688)
- perf(workbook): harden headless engine leadership gates (b0e399e9)
- fix(core): support literal leaf formula fast path (ebc51a7d)
- fix(ci): build protocol before release metadata push (1694f56f)
- fix(ci): tolerate mirrored github release push (ae7bcdce)
- fix(workbook): stabilize grid editing (a561edf4)
- fix(headless): bulk restore imported axis metadata (255f1a4a)

## Internal runtime changes

- chore(release): align runtime package versions (08e090c1)
- refactor(headless): split corpus verification helpers (0ffadc1b)
- chore(release): runtime packages v0.14.15 (073d1933)
- chore(release): runtime packages v0.14.16 (fa683aaa)
- chore(release): runtime packages v0.14.17 (1625a1e8)
- chore(release): runtime packages v0.14.18 (b7e712fd)
- chore(release): runtime packages v0.14.19 (f8805b43)

## 0.14.19

- Release type: patch
- Previous libraries tag: libraries-v0.14.14
- Manual override: no

## Fixes

- fix(headless): resolve real workbook import regressions (5ca9d46a)
- fix(ci): keep workbook worker release budget green (9f2dbfe5)
- fix(headless): restore dynamic spills from documents (2c57d688)
- perf(workbook): harden headless engine leadership gates (b0e399e9)
- fix(core): support literal leaf formula fast path (ebc51a7d)
- fix(ci): build protocol before release metadata push (1694f56f)
- fix(ci): tolerate mirrored github release push (ae7bcdce)
- fix(workbook): stabilize grid editing (a561edf4)
- fix(headless): bulk restore imported axis metadata (255f1a4a)

## Internal runtime changes

- chore(release): align runtime package versions (08e090c1)
- refactor(headless): split corpus verification helpers (0ffadc1b)
- chore(release): runtime packages v0.14.15 (073d1933)
- chore(release): runtime packages v0.14.16 (fa683aaa)
- chore(release): runtime packages v0.14.17 (1625a1e8)
- chore(release): runtime packages v0.14.18 (b7e712fd)

## 0.14.18

- Release type: patch
- Previous libraries tag: libraries-v0.14.14
- Manual override: no

## Fixes

- fix(headless): resolve real workbook import regressions (5ca9d46a)
- fix(ci): keep workbook worker release budget green (9f2dbfe5)
- fix(headless): restore dynamic spills from documents (2c57d688)
- perf(workbook): harden headless engine leadership gates (b0e399e9)
- fix(core): support literal leaf formula fast path (ebc51a7d)
- fix(ci): build protocol before release metadata push (1694f56f)
- fix(ci): tolerate mirrored github release push (ae7bcdce)
- fix(workbook): stabilize grid editing (a561edf4)
- fix(headless): bulk restore imported axis metadata (255f1a4a)

## Internal runtime changes

- chore(release): align runtime package versions (08e090c1)
- refactor(headless): split corpus verification helpers (0ffadc1b)
- chore(release): runtime packages v0.14.15 (073d1933)
- chore(release): runtime packages v0.14.16 (fa683aaa)
- chore(release): runtime packages v0.14.17 (1625a1e8)

## 0.14.17

- Release type: patch
- Previous libraries tag: libraries-v0.14.14
- Manual override: no

## Fixes

- fix(headless): resolve real workbook import regressions (5ca9d46a)
- fix(ci): keep workbook worker release budget green (9f2dbfe5)
- fix(headless): restore dynamic spills from documents (2c57d688)
- perf(workbook): harden headless engine leadership gates (b0e399e9)
- fix(core): support literal leaf formula fast path (ebc51a7d)
- fix(ci): build protocol before release metadata push (1694f56f)
- fix(ci): tolerate mirrored github release push (ae7bcdce)

## Internal runtime changes

- chore(release): align runtime package versions (08e090c1)
- refactor(headless): split corpus verification helpers (0ffadc1b)
- chore(release): runtime packages v0.14.15 (073d1933)
- chore(release): runtime packages v0.14.16 (fa683aaa)

## 0.14.16

- Release type: patch
- Previous libraries tag: libraries-v0.14.14
- Manual override: no

## Fixes

- fix(headless): resolve real workbook import regressions (5ca9d46a)
- fix(ci): keep workbook worker release budget green (9f2dbfe5)
- fix(headless): restore dynamic spills from documents (2c57d688)
- perf(workbook): harden headless engine leadership gates (b0e399e9)
- fix(core): support literal leaf formula fast path (ebc51a7d)
- fix(ci): build protocol before release metadata push (1694f56f)
- fix(ci): tolerate mirrored github release push (ae7bcdce)

## Internal runtime changes

- chore(release): align runtime package versions (08e090c1)
- refactor(headless): split corpus verification helpers (0ffadc1b)
- chore(release): runtime packages v0.14.15 (073d1933)

## 0.14.15

- Release type: patch
- Previous libraries tag: libraries-v0.14.14
- Manual override: no

## Fixes

- fix(headless): resolve real workbook import regressions (5ca9d46a)
- fix(ci): keep workbook worker release budget green (9f2dbfe5)
- fix(headless): restore dynamic spills from documents (2c57d688)
- perf(workbook): harden headless engine leadership gates (b0e399e9)
- fix(core): support literal leaf formula fast path (ebc51a7d)
- fix(ci): build protocol before release metadata push (1694f56f)

## Internal runtime changes

- chore(release): align runtime package versions (08e090c1)
- refactor(headless): split corpus verification helpers (0ffadc1b)

## 0.14.14

- Release type: patch
- Previous libraries tag: libraries-v0.14.13
- Manual override: no

### Fixes

- fix(assistant): throttle rendered context sync (92edb870)
- fix(headless): publish xlsx subpath (5c6f6b76)

## 0.1.95

- Release type: patch
- Previous libraries tag: none
- Manual override: yes

## 0.1.1

- Publish a packed tarball so npm registry manifests resolve internal bilig dependencies correctly.

## 0.1.2

- Align the headless library package set onto a single publish version for npm consumers.
