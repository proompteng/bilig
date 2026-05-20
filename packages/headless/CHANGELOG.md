# Changelog

All notable changes to `@bilig/headless` will be documented in this file.

This package is released as part of the aligned bilig library package set.

## 0.40.21

- Release type: patch
- Previous libraries tag: libraries-v0.40.20
- Manual override: no

## Fixes
- perf(xlsx): intern streamed worksheet metadata (2afd7a34)

## 0.40.20

- Release type: patch
- Previous libraries tag: libraries-v0.40.19
- Manual override: no

## Fixes
- perf(xlsx): avoid lazy metadata cell expansion (0c2d693b)

## 0.40.19

- Release type: patch
- Previous libraries tag: libraries-v0.40.18
- Manual override: no

## Fixes
- perf(xlsx): materialize compressed style ranges (1873f142)
- perf(core): fast-bind restored direct scalar formulas (72afb11c)

## 0.40.18

- Release type: patch
- Previous libraries tag: libraries-v0.40.17
- Manual override: no

## Fixes
- perf(xlsx): scope shared strings per sheet (ef6dc2b2)

## 0.40.17

- Release type: patch
- Previous libraries tag: libraries-v0.40.16
- Manual override: no

## Fixes
- perf(xlsx): avoid eager shared-string arena copies (9e0b47aa)
- perf(core): summarize far shifted aggregate pages (f90f01a5)

## 0.40.16

- Release type: patch
- Previous libraries tag: libraries-v0.40.15
- Manual override: no

## Fixes
- fix(xlsx): fall back on untyped streamed metadata (2ee22194)

## 0.40.15

- Release type: patch
- Previous libraries tag: libraries-v0.40.14
- Manual override: no

## Fixes
- perf(core): bulk bind fresh direct scalar runs (047b854a)

## 0.40.14

- Release type: patch
- Previous libraries tag: libraries-v0.40.13
- Manual override: no

## Fixes
- perf(core): bucket compound exact criteria aggregates (62cd7e72)

## 0.40.13

- Release type: patch
- Previous libraries tag: libraries-v0.40.12
- Manual override: no

## Fixes
- fix(xlsx): compact near-dense import coordinates (d15533c7)
- perf(core): restore prepared runtime family runs (4312390b)

## 0.40.12

- Release type: patch
- Previous libraries tag: libraries-v0.40.11
- Manual override: no

## Fixes
- perf(core): replay large formula family runs (03488e2d)
- fix(xlsx): keep rich shared strings lazy (d7acef49)

## 0.40.11

- Release type: patch
- Previous libraries tag: libraries-v0.40.10
- Manual override: no

## Fixes
- fix(xlsx): pack shared string arena storage (dd790620)

## 0.40.10

- Release type: patch
- Previous libraries tag: libraries-v0.40.9
- Manual override: no

## Fixes
- fix(xlsx): reduce dense arena overgrowth (06d3a15d)

## 0.40.9

- Release type: patch
- Previous libraries tag: libraries-v0.40.8
- Manual override: no

## Fixes
- fix(xlsx): defer large style coordinates (cdffc3ae)

## 0.40.8

- Release type: patch
- Previous libraries tag: libraries-v0.40.7
- Manual override: no

## Fixes
- fix(xlsx): stream implicit worksheet refs (935e2b33)

## 0.40.7

- Release type: patch
- Previous libraries tag: libraries-v0.40.6
- Manual override: no

## Fixes
- fix(xlsx): release dense import buffers (5d8bee40)

## 0.40.6

- Release type: patch
- Previous libraries tag: libraries-v0.40.5
- Manual override: no

## Fixes
- fix(xlsx): reduce lazy snapshot materialization memory (07ee4e89)

## 0.40.5

- Release type: patch
- Previous libraries tag: libraries-v0.40.4
- Manual override: no

## Fixes
- fix(xlsx): stream OLE control artifacts (44b46ee8)

## 0.40.4

- Release type: patch
- Previous libraries tag: libraries-v0.40.3
- Manual override: no

## Fixes
- fix(xlsx): fail fast on corrupt zip entries (3b9d58d4)
- fix(xlsx): stream data validation metadata (917ac2e6)

## 0.40.3

- Release type: patch
- Previous libraries tag: libraries-v0.40.2
- Manual override: no

## Fixes
- fix(xlsx): add file-backed import source path (8addb2e1)

## 0.40.2

- Release type: patch
- Previous libraries tag: libraries-v0.40.1
- Manual override: no

## Fixes
- fix(xlsx): stream complex import artifacts (f7469d8c)

## 0.40.1

- Release type: patch
- Previous libraries tag: libraries-v0.40.0
- Manual override: no

## Fixes
- fix(xlsx): release large import arena scratch (717b6a8a)

## 0.40.0

- Release type: minor
- Previous libraries tag: libraries-v0.39.0
- Manual override: no

## Features
- feat(excel-import): reduce xlsx import memory (6c04e41d)
- feat(excel-import): compact dense runtime coordinates (052c5349)

## Fixes
- fix(core): route metadata formulas through js parity (4d72b3de)
- perf(core): index compound exact criteria aggregates (0e2bb2dc)
- fix(xlsx): bound large workbook import builds (74210d5e)
- fix(xlsx): preserve unsupported chart drawings (892322e6)
- fix(excel-import): preserve compact import fidelity (a0340a33)
- fix(xlsx): stream cached formula imports (415e0cb1)
- fix(xlsx): warn on data table formula imports (f9cae29e)
- fix(xlsx): reduce fallback import memory (0b3105e6)

## Internal runtime changes
- test(core): add semantic invariant gate (80c73e48)
- docs(agent): add recalc skill metadata (d3aa7850)
- refactor(excel-import): use shared workbook semantics in tests (38f49c24)

## 0.39.0

- Release type: minor
- Previous libraries tag: libraries-v0.38.3
- Manual override: no

## Features
- feat(core): centralize workbook semantic projection (e7d9434d)

## Internal runtime changes
- ci(release): use current npm for runtime assets (7e0a788d)
- docs(agent): sync workpaper discovery release docs (2c1d76a8)

## 0.38.3

- Release type: patch
- Previous libraries tag: libraries-v0.38.2
- Manual override: no

## Fixes
- fix(headless): make WorkPaper config rebuild rollback atomic (5306047f)
- perf(formula): accelerate mixed criteria predicates (063d9edc)
- fix(core): preserve formula binding timeout failures (c9dc9f48)
- perf(core): widen exact criteria aggregate buckets (1596d969)
- perf(formula): batch native lookup recalc (c58d4722)
- fix(core): enforce operation evaluation budgets (332ef129)

## Internal runtime changes
- docs(npm): canonicalize scoped runtime packages (4d9202f0)
- test(headless): cover split WorkPaper surface base (4d7dc45c)
- test(headless): allow whole-column criteria regression on CI (6b7ccebd)
- ci(release): skip local hooks for runtime publish pushes (4cebaf17)

## 0.38.2

- Release type: patch
- Previous libraries tag: libraries-v0.38.1
- Manual override: no

## Fixes

- perf(core): restore runtime formula family runs (662ab10f)

## 0.38.1

- Release type: patch
- Previous libraries tag: libraries-v0.38.0
- Manual override: no

## Fixes

- fix(web): centralize projected local delta authority (48c77b09)
- perf(formula): preallocate scalar delta closure buffers (5018ad98)

## Internal runtime changes

- chore(format): normalize headless docs (5b08b1ff)

## 0.38.0

- Release type: minor
- Previous libraries tag: libraries-v0.37.2
- Manual override: no

## Features

- feat(runtime): add scoped Bilig npm packages (b2b1a825)

## Fixes

- perf(core): chunk initial direct scalar runs (5cef046e)

## 0.37.2

- Release type: patch
- Previous libraries tag: libraries-v0.37.1
- Manual override: no

## Fixes

- fix(zero): share persisted value guards (10bab669)
- perf(formula): tighten scalar row-pair batch writes (586e30cb)
- perf(core): avoid reparsing initial formula templates (a7d70e4f)

## 0.37.1

- Release type: patch
- Previous libraries tag: libraries-v0.37.0
- Manual override: no

## Fixes

- perf(core): share written column tracking (4a22f9fb)
- perf(formula): route large ifs aggregates through native predicate (30cf1116)
- perf(core): skip supported formula cache parses (2e44b5a5)
- perf(formula): trust direct scalar closure deltas (3b6bbe25)
- perf(formula): tighten scalar column batch writes (51b72473)

## Internal runtime changes

- docs(growth): route sheetjs users to named package (bd0987f8)

## 0.37.0

- Release type: minor
- Previous libraries tag: libraries-v0.36.2
- Manual override: no

## Features

- feat(xlsx): expose sheetjs recalc command (ad9ad52f)

## Fixes

- perf(formula): add native predicate criteria aggregation (e5aabb7b)

## 0.36.2

- Release type: patch
- Previous libraries tag: libraries-v0.36.1
- Manual override: no

## Fixes

- fix(docs): route sheetjs users to live xlsx package (e8bfef83)

## 0.36.1

- Release type: patch
- Previous libraries tag: libraries-v0.36.0
- Manual override: no

## Fixes

- perf(formula): widen native over-limit initialization (6198ed6f)
- perf(core): speed wide dense cell allocation (369b2f5e)
- fix(release): skip unprovisioned runtime packages (9d486e56)

## 0.36.0

- Release type: minor
- Previous libraries tag: libraries-v0.35.1
- Manual override: no

## Features

- feat(recalc): add sheetjs formula recalc package (b27bbd1a)

## Fixes

- perf(core): trim formula initialization bookkeeping (3e58a650)

## 0.35.1

- Release type: patch
- Previous libraries tag: libraries-v0.35.0
- Manual override: no

## Fixes

- perf(formula): native anchored prefix initialization (da6b943d)

## Internal runtime changes

- docs(growth): target sheetjs formula readback traffic (55faaa0b)

## 0.35.0

- Release type: minor
- Previous libraries tag: libraries-v0.34.1
- Manual override: no

## Features

- feat(recalc): cover incumbent xlsx formula bridges (bb5da689)

## Fixes

- perf(formula): retune native direct scalar initialization (f5b076ac)
- perf(headless): skip scalar inspection compiles (fc973438)

## 0.34.1

- Release type: patch
- Previous libraries tag: libraries-v0.34.0
- Manual override: no

## Fixes

- perf(formula): promote xlookup spill returns (060be2c8)

## 0.34.0

- Release type: minor
- Previous libraries tag: libraries-v0.33.1
- Manual override: no

## Features

- feat(recalc): ship package-native formula proof CLIs (b9d6bfd8)

## Fixes

- perf(core): reuse safe inline initial formulas (3944959a)

## 0.33.1

- Release type: patch
- Previous libraries tag: libraries-v0.33.0
- Manual override: no

## Fixes

- perf(formula): promote xlookup approximate matching (25bd729f)

## 0.33.0

- Release type: minor
- Previous libraries tag: libraries-v0.32.9
- Manual override: no

## Features

- feat(bilig-workpaper): ship agent-ready npm entrypoints (9da4ee33)

## 0.32.9

- Release type: patch
- Previous libraries tag: libraries-v0.32.8
- Manual override: no

## Fixes

- perf(core): avoid string keys in formula family init (028e5084)

## 0.32.8

- Release type: patch
- Previous libraries tag: libraries-v0.32.7
- Manual override: no

## Fixes

- perf(formula): add native row-chain scalar init (2ff802d5)

## 0.32.7

- Release type: patch
- Previous libraries tag: libraries-v0.32.6
- Manual override: no

## Fixes

- fix(release): sync static discovery references (0eaed367)
- fix(xlsx-formula-recalc): surface high-traffic recalc entrypoint (d7f76fde)

## 0.32.6

- Release type: patch
- Previous libraries tag: libraries-v0.32.5
- Manual override: no

## Fixes

- perf(formula): write native scalar init into kernel store (c08356c1)
- fix(docs): sync 0.32.5 public agent links (a80d2c7d)

## 0.32.5

- Release type: patch
- Previous libraries tag: libraries-v0.32.4
- Manual override: no

## Fixes

- fix(release): sync agent discovery docs (26007971)

## 0.32.4

- Release type: patch
- Previous libraries tag: libraries-v0.32.3
- Manual override: no

## Fixes

- fix(formula): align headless error semantics (b1c774a9)
- fix(core): ignore stale direct formula deltas (4e6bf441)

## Internal runtime changes

- docs(discovery): sync 0.32.3 agent surfaces (5bdb4132)

## 0.32.3

- Release type: patch
- Previous libraries tag: libraries-v0.32.2
- Manual override: no

## Fixes

- perf(headless): fast-path public literal batches (80edd87e)

## Internal runtime changes

- docs(discovery): sync 0.32.2 agent surfaces (291aa4f2)
- docs(growth): route evaluators to recalc packages (fdde26f3)

## 0.32.2

- Release type: patch
- Previous libraries tag: libraries-v0.32.1
- Manual override: no

## Fixes

- perf(formula): batch direct scalar recalc natively (ecd34ab9)

## Internal runtime changes

- docs(discovery): sync 0.32.1 agent surfaces (f297f822)

## 0.32.1

- Release type: patch
- Previous libraries tag: libraries-v0.32.0
- Manual override: no

## Fixes

- perf(headless): narrow core startup imports (f3ddf1ab)
- fix(package): use publishable workpaper package name (478a0039)

## Internal runtime changes

- docs(discovery): sync 0.32 agent surfaces (eee388d2)

## 0.32.0

- Release type: minor
- Previous libraries tag: libraries-v0.31.1
- Manual override: no

## Features

- feat(package): add unscoped bilig runtime package (e3fd2c02)

## Fixes

- perf(formula): reduce matched criteria aggregates natively (95da4cf9)

## Internal runtime changes

- docs(discovery): sync 0.31.1 agent surfaces (37df83a0)
- docs(discovery): sync 0.31.1 agent surfaces (5ec4853e)

## 0.31.1

- Release type: patch
- Previous libraries tag: libraries-v0.31.0
- Manual override: no

## Fixes

- perf(core): cache runtime restore string ids (87276c1e)

## Internal runtime changes

- docs(discovery): sync 0.31.0 agent surfaces (03bfead1)

## 0.31.0

- Release type: minor
- Previous libraries tag: libraries-v0.30.2
- Manual override: no

## Features

- feat(package): add exceljs formula recalc adapter (e24ca045)

## Fixes

- perf(headless): narrow custom function adapter import (5d04174c)
- fix(agent): throttle passive context churn (0febed2b)

## Internal runtime changes

- docs(discovery): sync 0.30.2 agent surfaces (fc4a6030)

## 0.30.2

- Release type: patch
- Previous libraries tag: libraries-v0.30.1
- Manual override: no

## Fixes

- fix(grid): stabilize editor terminal shortcuts (68e0d32f)

## 0.30.1

- Release type: patch
- Previous libraries tag: libraries-v0.30.0
- Manual override: no

## Fixes

- fix(formula): match excel log error semantics (15d01a85)
- fix(package): resolve xlsx recalc workspace imports (d0485880)

## Internal runtime changes

- docs(discovery): sync runtime package 0.30.0 (9de932db)

## 0.30.0

- Release type: minor
- Previous libraries tag: libraries-v0.29.0
- Manual override: no

## Features

- feat(package): add xlsx formula recalc npm entrypoint (08ac4689)
- feat(formula): batch native direct scalar initialization (03df35e9)
- feat(formula): add native aggregate matrix batches (b25d12ce)

## Fixes

- fix(grid): validate visible fill coverage by geometry (39f36931)
- perf(core): streamline clean direct scalar deltas (016b2153)
- fix(xlsx-formula-recalc): inherit workspace aliases (7876c92b)

## Internal runtime changes

- docs(discovery): sync runtime package 0.29.0 (17dabbe1)
- chore(format): normalize xlsx formula readme (b02af111)

## 0.29.0

- Release type: minor
- Previous libraries tag: libraries-v0.28.2
- Manual override: yes

## Features

- feat(package): add unscoped bilig npm entrypoint (c43c5b69)

## Internal runtime changes

- docs(discovery): sync runtime package 0.28.2 (26662b3c)

## 0.28.2

- Release type: patch
- Previous libraries tag: libraries-v0.28.1
- Manual override: no

## Fixes

- perf(core): group exact criteria aggregates (6d6a2d0c)
- perf(excel-import): speed style-only blank stripping (eb8b1f2d)

## Internal runtime changes

- docs(discovery): sync runtime package 0.28.1 (091ab9fd)

## 0.28.1

- Release type: patch
- Previous libraries tag: libraries-v0.28.0
- Manual override: no

## Fixes

- fix(corpus): expand recent workbook verification (8ae4d8dc)

## 0.28.0

- Release type: minor
- Previous libraries tag: libraries-v0.27.0
- Manual override: no

## Features

- feat(agent): add openai agents workpaper tools (cd9bb8d0)

## Internal runtime changes

- docs(discovery): sync runtime package 0.27.0 (37c6d76a)
- docs(agent): harden mcp server card schemas (9e7c1c5d)

## 0.27.0

- Release type: minor
- Previous libraries tag: libraries-v0.26.1
- Manual override: no

## Features

- feat(create-workpaper): add agent starter (87fb5c74)

## Internal runtime changes

- docs(discovery): sync runtime package 0.26.1 (9c2eefa0)

## 0.26.1

- Release type: patch
- Previous libraries tag: libraries-v0.26.0
- Manual override: no

## Fixes

- perf(headless): split mcp exports from main entry (32a5dc8b)
- fix(release): upload headless mcpb assets (40288dd4)

## Internal runtime changes

- docs(discovery): sync runtime package 0.26.0 (09ec7e80)

## 0.26.0

- Release type: minor
- Previous libraries tag: libraries-v0.25.7
- Manual override: no

## Features

- feat(headless): add mcp challenge cli (a0f3eba8)

## Internal runtime changes

- docs(discovery): sync headless 0.25.7 references (9cd5e8c4)
- docs(mcp): refresh registry distribution proof (b2d6f700)

## 0.25.7

- Release type: patch
- Previous libraries tag: libraries-v0.25.6
- Manual override: no

## Fixes

- perf(core): trust scalar template translation (191bc3af)

## Internal runtime changes

- docs(discovery): sync headless 0.25.6 references (96a3cbe7)

## 0.25.6

- Release type: patch
- Previous libraries tag: libraries-v0.25.5
- Manual override: no

## Fixes

- fix(formula): preserve cached formula parity (0fd14b37)

## Internal runtime changes

- docs(mcp): add Smithery install surface (aeddfc39)

## 0.25.5

- Release type: patch
- Previous libraries tag: libraries-v0.25.4
- Manual override: no

## Fixes

- perf(core): skip redundant csv import recalcs (b1254e8a)

## Internal runtime changes

- docs(discovery): sync headless 0.25.4 references (a73453a9)

## 0.25.4

- Release type: patch
- Previous libraries tag: libraries-v0.25.3
- Manual override: no

## Fixes

- perf(excel-import): enable formula import restore sidecar (8313f2d7)

## Internal runtime changes

- docs(discovery): sync headless 0.25.3 references (ea015ad6)
- docs(agent): publish Claude Desktop MCPB install path (6d25dad4)

## 0.25.3

- Release type: patch
- Previous libraries tag: libraries-v0.25.2
- Manual override: no

## Fixes

- perf(excel-import): add import restore coordinate fast path (0788025e)

## Internal runtime changes

- docs(discovery): sync headless 0.25.2 references (f3196fb4)
- test(headless): stabilize guarded sumifs budget (e33507fc)

## 0.25.2

- Release type: patch
- Previous libraries tag: libraries-v0.25.1
- Manual override: no

## Fixes

- perf(core): precompile csv numeric parsing (9c4e5af1)

## Internal runtime changes

- docs(discovery): sync headless 0.25.1 surfaces (fa0ab38c)

## 0.25.1

- Release type: patch
- Previous libraries tag: libraries-v0.25.0
- Manual override: no

## Fixes

- fix(headless): harden recent workbook parity (df0b9d88)

## Internal runtime changes

- docs(discovery): sync headless 0.25.0 surfaces (6f372228)

## 0.25.0

- Release type: minor
- Previous libraries tag: libraries-v0.24.5
- Manual override: no

## Features

- feat(headless): add agent workbook challenge cli (bf7c2f81)

## 0.24.5

- Release type: patch
- Previous libraries tag: libraries-v0.24.4
- Manual override: no

## Fixes

- perf(headless): reduce physical range write lookups (2170ba56)

## Internal runtime changes

- docs(headless): sync discovery package version (a17e8163)

## 0.24.4

- Release type: patch
- Previous libraries tag: libraries-v0.24.3
- Manual override: no

## Fixes

- fix(workbook): stabilize grid typography and focus ownership (96abe1e7)

## Internal runtime changes

- chore(ci): update github action pins (b33db000)

## 0.24.3

- Release type: patch
- Previous libraries tag: libraries-v0.24.2
- Manual override: no

## Fixes

- perf(headless): expand workpaper fast paths (836ac431)
- fix(core): preserve structural insert entries in sync batches (1ede985b)
- fix(core): clean up rebased fast paths (aa86ac6f)
- perf(core): speed trusted template restore (9740a6a2)

## Internal runtime changes

- refactor(headless): split oversized runtime files (01d09e16)
- docs(headless): refresh published commands for 0.24.2 (7d1d5fb9)
- docs(agent): add workbook challenge (7e7cd271)
- test(core): handle protection rejections in fuzz (5996cfd6)
- test(headless): include fast path surface parity (a7b493b6)

## 0.24.2

- Release type: patch
- Previous libraries tag: libraries-v0.24.1
- Manual override: no

## Fixes

- fix(mcp): publish hosted endpoint metadata (82595db1)

## 0.24.1

- Release type: patch
- Previous libraries tag: libraries-v0.24.0
- Manual override: no

## Fixes

- fix(workbook): harden edit and tile clear races (6135fe6c)

## 0.24.0

- Release type: minor
- Previous libraries tag: libraries-v0.23.4
- Manual override: no

## Features

- feat(mcp): add remote workpaper endpoint (a2349d8e)

## 0.23.4

- Release type: patch
- Previous libraries tag: libraries-v0.23.3
- Manual override: no

## Fixes

- fix(workbook): preserve tile presentation mutations (a234caaf)

## Internal runtime changes

- docs(headless): add skill registry metadata (d35669b6)
- docs(agent): pin headless npm exec commands (fb68ce57)
- docs(agent): harden public skill command guidance (5f9f5ec7)
- docs(agent): refresh discovery surfaces (1c527b72)

## 0.23.3

- Release type: patch
- Previous libraries tag: libraries-v0.23.2
- Manual override: no

## Fixes

- fix(formula): preserve 3d structural range metadata (529cb889)

## 0.23.2

- Release type: patch
- Previous libraries tag: libraries-v0.23.1
- Manual override: no

## Fixes

- fix(formula): translate 3d range references (af35362b)

## 0.23.1

- Release type: patch
- Previous libraries tag: libraries-v0.23.0
- Manual override: no

## Fixes

- fix(formula): propagate criteria aggregate ref errors (47f63810)

## Internal runtime changes

- docs(mcp): track registry refresh lag (3f70e797)

## 0.23.0

- Release type: minor
- Previous libraries tag: libraries-v0.22.2
- Manual override: no

## Features

- feat(mcp): expose workpaper prompts and resources (79a4d16a)

## Fixes

- fix(mcp): keep server metadata publishable (5545f395)

## 0.22.2

- Release type: patch
- Previous libraries tag: libraries-v0.22.1
- Manual override: no

## Fixes

- fix(excel-import): avoid chartsheet worksheet path fallback (278f862b)

## Internal runtime changes

- ci(runtime): stop direct GitHub release pushes (c7cc8682)
- docs(agent): publish agent discovery manifest (3ab9232a)

## 0.22.1

- Release type: patch
- Previous libraries tag: libraries-v0.22.0
- Manual override: no

## Fixes

- fix(docs): expose raw agent skill endpoints (ec73bf17)

## Internal runtime changes

- docs(agent): publish agent discovery pack (93f0e0bf)

## 0.22.0

- Release type: minor
- Previous libraries tag: libraries-v0.21.1
- Manual override: no

## Features

- feat(headless): ship formula clinic cli (7aef71db)

## Internal runtime changes

- docs(headless): align package agent notes with mcp init (d837d0ca)
- ci(runtime): keep mirror package checks green (e4d7fd96)
- docs(mcp): mark glama release live (f34a7d40)
- ci(runtime): skip release planning on fetch failure (0b386bcf)
- ci(runtime): allow mirror release dispatch (1067113d)

## 0.21.1

- Release type: patch
- Previous libraries tag: libraries-v0.21.0
- Manual override: no

## Fixes

- fix(headless): preserve external formula caches (f67ed038)

## Internal runtime changes

- docs(mcp): promote one-command workpaper init (c0b29802)

## 0.21.0

- Release type: minor
- Previous libraries tag: libraries-v0.20.0
- Manual override: no

## Features

- feat(headless): initialize demo workpaper for mcp (4120d324)

## Internal runtime changes

- refactor(excel): split pivot export writer (7ef3640a)
- docs(discovery): track mcp registry publish lag (afd56bce)

## 0.20.0

- Release type: minor
- Previous libraries tag: libraries-v0.19.3
- Manual override: no

## Features

- feat(import): add Excel formula and pivot semantics (82da4b78)

## 0.19.3

- Release type: patch
- Previous libraries tag: libraries-v0.19.2
- Manual override: no

## Fixes

- fix(formula): support whole-axis xlookup ranges (f8ecaf81)

## Internal runtime changes

- docs(growth): fix cloned example commands (9b282f6d)

## 0.19.2

- Release type: patch
- Previous libraries tag: libraries-v0.19.1
- Manual override: no

## Fixes

- fix(formula): exclude non-text wildcard criteria matches (07b9303f)

## 0.19.1

- Release type: patch
- Previous libraries tag: libraries-v0.19.0
- Manual override: no

## Fixes

- fix(formula): coerce blank indirect references (4805458e)

## 0.19.0

- Release type: minor
- Previous libraries tag: libraries-v0.18.29
- Manual override: no

## Features

- feat(headless): add mcp output schemas (d9528032)

## Internal runtime changes

- docs(growth): refresh mcp registry evidence (71eef9cf)

## 0.18.29

- Release type: patch
- Previous libraries tag: libraries-v0.18.28
- Manual override: no

## Fixes

- perf(engine): enable column indexes by default (2b91d1dd)
- fix(corpus): harden recent workbook headless gate (f1519cea)

## Internal runtime changes

- docs(growth): add agent handoff prompt (75eec5c8)

## 0.18.28

- Release type: patch
- Previous libraries tag: libraries-v0.18.27
- Manual override: yes

## Internal runtime changes

- refactor(core): unify lookup write planning (e3037d00)
- docs(growth): refresh v0.18.27 registry evidence (80a124e3)
- docs(growth): add headless agent handbook (564846bb)
- docs(headless): publish agent package notes (3544d8c5)

## 0.18.27

- Release type: patch
- Previous libraries tag: libraries-v0.18.26
- Manual override: no

## Fixes

- fix(core): invoke lambda defined names (1531f8a0)

## Internal runtime changes

- refactor(core): harden mutation and lookup tracking (75f4d020)

## 0.18.26

- Release type: patch
- Previous libraries tag: libraries-v0.18.25
- Manual override: no

## Fixes

- perf(core): narrow structural and lookup hot paths (f68d42bc)

## Internal runtime changes

- refactor(core): isolate mutation inverse ops (79172a88)
- refactor(core): isolate batch cell value mutations (35392b17)
- docs(growth): refresh v0.18.25 registry evidence (5b56b67f)
- refactor(formula): isolate lookup match opcodes (3c7f66ac)
- refactor(core): isolate batch formula mutations (34e5aeb9)

## 0.18.25

- Release type: patch
- Previous libraries tag: libraries-v0.18.24
- Manual override: no

## Fixes

- perf(core): defer fresh logical cell indexes (69fa5800)

## Internal runtime changes

- docs(growth): refresh mcp directory follow-up state (2ba76bb2)
- refactor(core): isolate clear cell mutation flow (6f49726c)
- docs(growth): refresh registry evidence for v0.18.24 (97ad0e47)
- refactor(core): isolate literal cell mutation flow (d3577054)
- docs(growth): align package discovery keywords (44a85351)
- refactor(core): isolate formula cell mutation flow (6c0e0a6c)

## 0.18.24

- Release type: patch
- Previous libraries tag: libraries-v0.18.23
- Manual override: no

## Fixes

- perf(core): bulk-restore dense runtime images (afe392dd)
- fix(core): coalesce fragmented style rectangles (46041b3d)

## Internal runtime changes

- refactor(core): split direct scalar column fast paths (ca0da009)
- refactor(core): isolate structural formula impacts (3d91bbb2)
- docs(growth): refresh discovery conversion evidence (96fab221)
- docs(growth): add mcprepository listing evidence (eedc7c42)
- refactor(formula): isolate binder dependencies (36718d95)

## 0.18.23

- Release type: patch
- Previous libraries tag: libraries-v0.18.22
- Manual override: no

## Fixes

- perf(headless): fast-load dense numeric sheets (526723d4)

## Internal runtime changes

- chore(release): format headless changelog (55dc9215)

## 0.18.22

- Release type: patch
- Previous libraries tag: libraries-v0.18.21
- Manual override: no

## Fixes

- perf(core): fuse fresh aggregate matrix writes (1c6bf730)

## Internal runtime changes

- docs(growth): sync published package evidence (e6e7d288)
- refactor(wasm): centralize lookup candidate comparison (572a091d)
- refactor(core): split formula binding controllers (7c23dca2)

## 0.18.21

- Release type: patch
- Previous libraries tag: libraries-v0.18.20
- Manual override: no

## Fixes

- perf(core): reserve fresh aggregate formula blocks (015b806b)
- fix(mcp): report package version over stdio (ccb321ec)

## Internal runtime changes

- refactor(docs): split discovery trust gate (663e2fa7)

## 0.18.20

- Release type: patch
- Previous libraries tag: libraries-v0.18.19
- Manual override: no

## Fixes

- fix(mcp): expose file-backed tools to directory scanners (352bccb9)

## 0.18.19

- Release type: patch
- Previous libraries tag: libraries-v0.18.18
- Manual override: no

## Fixes

- fix(headless): materialize dense load rectangles safely (82e77db5)
- fix(zero): centralize schema bootstrap (ace3ec96)

## 0.18.18

- Release type: patch
- Previous libraries tag: libraries-v0.18.17
- Manual override: no

## Fixes

- fix(wasm-kernel): support array sumproduct operands (5db44f40)
- fix(wasm-kernel): vectorize unary negation (5f6a5a6c)
- perf(headless): accelerate dense initialization and fresh aggregate formulas (ca1114e0)

## Internal runtime changes

- docs(mcp): add directory scanner docker target (b134572d)

## 0.18.17

- Release type: patch
- Previous libraries tag: libraries-v0.18.16
- Manual override: no

## Fixes

- fix(wasm-kernel): align criteria aggregate array semantics (92c71f48)

## Internal runtime changes

- docs(mcp): add file-backed transcript proof (5c554c46)

## 0.18.16

- Release type: patch
- Previous libraries tag: libraries-v0.18.15
- Manual override: no

## Fixes

- fix(wasm-kernel): align table lookup array semantics (224bfd85)

## 0.18.15

- Release type: patch
- Previous libraries tag: libraries-v0.18.14
- Manual override: no

## Fixes

- fix(wasm-kernel): align lookup array semantics (ae9b4ba5)

## 0.18.14

- Release type: patch
- Previous libraries tag: libraries-v0.18.13
- Manual override: no

## Fixes

- fix(wasm-kernel): support array metadata fast paths (f3d7ad3a)

## 0.18.13

- Release type: patch
- Previous libraries tag: libraries-v0.18.12
- Manual override: no

## Fixes

- fix(wasm-kernel): preserve dynamic array cell values (018d4535)

## 0.18.12

- Release type: patch
- Previous libraries tag: libraries-v0.18.11
- Manual override: no

## Fixes

- fix(headless): improve npm discovery metadata (f15da89d)

## 0.18.11

- Release type: patch
- Previous libraries tag: libraries-v0.18.10
- Manual override: no

## Fixes

- fix(release): sync npm evidence after version bumps (da5a9fb2)
- fix(release): build before footprint sync (f293c629)

## 0.18.10

- Release type: patch
- Previous libraries tag: libraries-v0.18.9
- Manual override: no

## Fixes

- fix(headless): refresh npm evaluator copy (bf35ad7d)

## Internal runtime changes

- chore(docs): refresh runtime evidence for 0.18.9 (56cd815e)
- docs(growth): add plain-language bilig fit guide (207f8484)

## 0.18.9

- Release type: patch
- Previous libraries tag: libraries-v0.18.8
- Manual override: no

## Fixes

- fix(runtime): align starter package version (d8c0770c)

## Internal runtime changes

- ci(runtime): publish starter in common package workflow (580b9741)

## 0.18.8

- Release type: patch
- Previous libraries tag: libraries-v0.18.7
- Manual override: no

## Fixes

- perf(headless): bypass reducer for lazy formula edits (771ebbda)

## 0.18.7

- Release type: patch
- Previous libraries tag: libraries-v0.18.6
- Manual override: no

## Fixes

- perf(core): correct direct formula replacement metrics (f702ceaa)

## Internal runtime changes

- docs(growth): add agent xlsx recalculation page (ba9c766e)

## 0.18.6

- Release type: patch
- Previous libraries tag: libraries-v0.18.5
- Manual override: no

## Fixes

- fix(grid): harden visual fidelity and stale clears (9c604a80)

## Internal runtime changes

- docs(growth): add runnable xlsx proof (a3308f5d)

## 0.18.5

- Release type: patch
- Previous libraries tag: libraries-v0.18.4
- Manual override: no

## Fixes

- perf(core): skip formula-only aggregate input coverage (772c2cf7)

## Internal runtime changes

- docs(growth): add formula clinic report script (7de01748)

## 0.18.4

- Release type: patch
- Previous libraries tag: libraries-v0.18.3
- Manual override: no

## Fixes

- perf(headless): skip no-value structural reductions (2457cd77)

## Internal runtime changes

- chore(release): refresh headless public evidence (f3cc500e)
- docs(growth): add formula bug clinic (878b31db)

## 0.18.3

- Release type: patch
- Previous libraries tag: libraries-v0.18.2
- Manual override: no

## Fixes

- perf(core): defer cold structural formula families (b48b0c21)

## 0.18.2

- Release type: patch
- Previous libraries tag: libraries-v0.18.1
- Manual override: no

## Fixes

- fix(wasm-kernel): align dynamic array semantics (0710cd42)

## 0.18.1

- Release type: patch
- Previous libraries tag: libraries-v0.18.0
- Manual override: no

## Fixes

- fix(formula): align flatten array semantics (5d7725b3)

## 0.18.0

- Release type: minor
- Previous libraries tag: libraries-v0.17.1
- Manual override: no

## Features

- feat(community): add workbook fixture submission path (73798e44)

## Fixes

- perf(headless): cache row-literal formula templates (ddbfad3a)
- fix(formula): preserve ref errors in headless corpus (06549cf5)
- perf(headless): defer tail-append change detachment (cd286a0d)
- perf(core): tighten direct scalar delta hot path (021d59eb)
- perf(core): add primitive fresh cell attach path (caee68e0)
- fix(headless): keep npm keyword metadata compressed (9a355e7e)
- fix(formula): make lookup search modes authoritative (e84bdbbe)
- fix(formula): harden text scalar builtins (bb1c0d6e)
- fix(formula): respect quoted text format literals (372f66bb)
- perf(core): skip reverse edge scans for fresh formulas (aad986ee)
- fix(formula): pad ragged stack arrays (56aa2acb)

## Internal runtime changes

- docs(evidence): align documentation with current artifacts (96fd0a54)
- docs(community): link workbook fixture discussion (2b1ec511)
- docs(mcp): sharpen formula recalculation positioning (b7099696)
- build(create): move starter to scoped npm package (21dcbb26)

## 0.17.1

- Release type: patch
- Previous libraries tag: libraries-v0.17.0
- Manual override: no

## Fixes

- perf(headless): fast path dense mixed sheet loads (66e8d5e4)
- perf(headless): batch runtime snapshot column restores (d242a0f2)

## Internal runtime changes

- ci(create-workpaper): add npm publish gate (b24127b6)
- docs(create-workpaper): avoid unpublished starter command (9435acfd)
- docs(mcp): compare spreadsheet server choices (223a182f)
- chore(headless): sharpen npm discovery keywords (72038f18)

## 0.17.0

- Release type: minor
- Previous libraries tag: libraries-v0.16.28
- Manual override: no

## Features

- feat(create-workpaper): add one-command starter (bef96b48)

## 0.16.28

- Release type: patch
- Previous libraries tag: libraries-v0.16.27
- Manual override: no

## Fixes

- perf(headless): skip scalar formula dependency rebinding (20a5dea9)

## Internal runtime changes

- docs(growth): refresh public evidence (8f800b72)

## 0.16.27

- Release type: patch
- Previous libraries tag: libraries-v0.16.26
- Manual override: no

## Fixes

- fix: harden recalc completion and structural undo (4c3a9300)

## Internal runtime changes

- docs(headless): clarify Excel formula compatibility (c552da9b)
- docs(growth): humanize Show HN copy (62cd354a)

## 0.16.26

- Release type: patch
- Previous libraries tag: libraries-v0.16.25
- Manual override: no

## Fixes

- perf(headless): fast path literal fanout payloads (def91ef6)

## Internal runtime changes

- docs(growth): refresh public evidence for 0.16.25 (5d8efa02)
- docs(growth): tighten maintainer note and trust checks (6fbe664d)
- docs(growth): cover ExcelJS shared formula recalculation (ccf23559)
- docs(growth): target Excel calculation engine searches (b4517b13)

## 0.16.25

- Release type: patch
- Previous libraries tag: libraries-v0.16.24
- Manual override: no

## Fixes

- perf(headless): skip exact lookup batch recalc (fc6da3f3)

## Internal runtime changes

- docs(growth): sharpen maintainer launch copy (36f40361)

## 0.16.24

- Release type: patch
- Previous libraries tag: libraries-v0.16.23
- Manual override: no

## Fixes

- fix(core): restore aggregate formulas after row delete undo (376211e4)

## 0.16.23

- Release type: patch
- Previous libraries tag: libraries-v0.16.22
- Manual override: no

## Fixes

- perf(headless): speed up dense fresh cell allocation (4cd88da8)

## Internal runtime changes

- refactor(web): remove browser sqlite storage (031da20d)
- docs(growth): make show hn copy less generic (61044fd9)
- chore(release): refresh public evidence for 0.16.22 (185d2059)

## 0.16.22

- Release type: patch
- Previous libraries tag: libraries-v0.16.21
- Manual override: no

## Fixes

- perf(headless): preserve hydrated formula family runs (cc942261)

## Internal runtime changes

- docs(growth): add xlsx recalculation proof (7f8f8832)
- test(benchmarks): expand WorkPaper suite to 100 workloads (e70921e7)
- docs(growth): add xlsx recalculation decision page (45c8251e)
- docs(growth): add xlsx-calc alternative page (50dc2885)

## 0.16.21

- Release type: patch
- Previous libraries tag: libraries-v0.16.20
- Manual override: no

## Fixes

- perf(headless): reuse matrix plan numeric shape (4c9fb63f)

## Internal runtime changes

- chore(release): refresh headless footprint evidence (c0586bb2)

## 0.16.20

- Release type: patch
- Previous libraries tag: libraries-v0.16.19
- Manual override: no

## Fixes

- perf(headless): trim scalar closure allocations (356f880b)

## Internal runtime changes

- docs(growth): track jsgrids and refresh evidence (d8d71044)

## 0.16.19

- Release type: patch
- Previous libraries tag: libraries-v0.16.18
- Manual override: no

## Fixes

- perf(headless): streamline scalar formula cascades (085e7f5f)

## Internal runtime changes

- docs(growth): surface public review proof path (7a9d98df)

## 0.16.18

- Release type: patch
- Previous libraries tag: libraries-v0.16.17
- Manual override: no

## Fixes

- perf(headless): split fresh matrix literals before formulas (6a3eb7d7)

## Internal runtime changes

- chore(growth): refresh public evidence for 0.16.17 (bc14e116)

## 0.16.17

- Release type: patch
- Previous libraries tag: libraries-v0.16.16
- Manual override: no

## Fixes

- perf(headless): speed up structural aggregate row deletes (ee52d072)

## 0.16.16

- Release type: patch
- Previous libraries tag: libraries-v0.16.15
- Manual override: no

## Fixes

- fix(benchmarks): tolerate platform float drift (eb8f21d0)

## Internal runtime changes

- chore(growth): refresh headless performance evidence (bf2013be)

## 0.16.15

- Release type: patch
- Previous libraries tag: libraries-v0.16.14
- Manual override: no

## Fixes

- perf(core): defer structural delete formula undo work (d7ef3fac)

## Internal runtime changes

- docs(growth): refresh release evidence (25e2e4ed)
- docs(growth): route overview evaluators to proof paths (a42f2da0)

## 0.16.14

- Release type: patch
- Previous libraries tag: libraries-v0.16.13
- Manual override: no

## Fixes

- perf(core): skip region probes for fresh aggregate rows (9762dc68)

## Internal runtime changes

- docs(growth): refresh public release evidence (feba38f4)
- docs(growth): surface adoption blocker intake (d983476a)
- docs(growth): add release watch path (0b7080d1)

## 0.16.13

- Release type: patch
- Previous libraries tag: libraries-v0.16.12
- Manual override: no

## Fixes

- perf(core): reuse copied criteria formula results (f213cbf1)

## 0.16.12

- Release type: patch
- Previous libraries tag: libraries-v0.16.11
- Manual override: no

## Fixes

- perf(headless): trim append formula region work (1c2f1e0c)

## Internal runtime changes

- docs(headless): add formula workbook proof page (15030c53)

## 0.16.11

- Release type: patch
- Previous libraries tag: libraries-v0.16.10
- Manual override: no

## Fixes

- fix(storage): sanitize local style projections (1662cbe5)

## Internal runtime changes

- docs(headless): align public evidence with runtime release (7b736fac)

## 0.16.10

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
- perf(core): skip exact uniform lookup owner binding (22f8e13b)
- perf(core): skip empty tracked invalidation patches (6d9eaed0)
- fix(core): reject unsafe direct formula rows (e3d914e8)
- fix(core): guard unsafe template row keys (c88ff52b)
- perf(headless): keep initial formula refs compact (efe03036)
- fix(sync): validate event sequence integers (87f2e312)
- fix(protocol): validate cell snapshot metadata (6bda4133)
- fix(domain): validate structural op coordinates (18a00a4e)
- fix(domain): validate object footprint dimensions (8f72557a)
- fix(domain): validate sheet identity metadata (6456ba4f)
- fix(domain): validate metadata sequence fields (76786154)
- perf(headless): keep appended formula changes lazy (a98603c7)
- fix(sync): reject malformed literal events (5ebc1a7e)
- perf(headless): collapse safe formula matrix writes (66b4cc64)
- fix(protocol): validate cell snapshot values (c99e51c9)
- fix(protocol): validate workbook snapshot entries (96d2728a)

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
- chore(release): prepare runtime libraries 0.16.6 (32f4f64f)
- chore(release): prepare runtime libraries 0.16.7 (e61bc460)
- chore(release): prepare runtime libraries 0.16.9 (e30ef786)
- ci(release): cancel stale runtime package runs (ae4499f4)
- ci(release): isolate runtime package workflow runs (2a754125)
- ci(release): retry runtime metadata push races (a9cba127)

## 0.16.9

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
- perf(core): skip exact uniform lookup owner binding (22f8e13b)
- perf(core): skip empty tracked invalidation patches (6d9eaed0)
- fix(core): reject unsafe direct formula rows (e3d914e8)
- fix(core): guard unsafe template row keys (c88ff52b)
- perf(headless): keep initial formula refs compact (efe03036)

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
- chore(release): prepare runtime libraries 0.16.6 (32f4f64f)
- chore(release): prepare runtime libraries 0.16.7 (e61bc460)

## 0.16.8

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
- perf(core): skip exact uniform lookup owner binding (22f8e13b)
- perf(core): skip empty tracked invalidation patches (6d9eaed0)
- fix(core): reject unsafe direct formula rows (e3d914e8)
- fix(core): guard unsafe template row keys (c88ff52b)

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
- chore(release): prepare runtime libraries 0.16.6 (32f4f64f)
- chore(release): prepare runtime libraries 0.16.7 (e61bc460)

## 0.16.7

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
- perf(core): skip exact uniform lookup owner binding (22f8e13b)
- perf(core): skip empty tracked invalidation patches (6d9eaed0)

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
- chore(release): prepare runtime libraries 0.16.6 (32f4f64f)

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