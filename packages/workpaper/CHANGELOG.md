# Changelog

All notable changes to `@bilig/headless` will be documented in this file.

This package is released as part of the aligned bilig library package set.

## 0.164.11

- Release type: patch
- Previous libraries tag: libraries-v0.164.10
- Manual override: no

## Fixes

- fix(ci): restore green debt baseline (7cdfb765)

## 0.164.10

- Release type: patch
- Previous libraries tag: libraries-v0.164.9
- Manual override: no

## Fixes

- fix(core): stabilize direct aggregate fast paths (07448bd4)

## 0.164.9

- Release type: patch
- Previous libraries tag: libraries-v0.164.8
- Manual override: no

## Fixes

- fix(core): preserve table headers during undo (3b22698d)

## Internal runtime changes

- docs(discovery): deslop workpaper public surfaces (355eb34d)

## 0.164.8

- Release type: patch
- Previous libraries tag: libraries-v0.164.5
- Manual override: no

## Fixes

- fix(headless): clarify workpaper proof limits (5d4f7a2b)
- fix(release): sync workpaper service proof versions (b9664688)

## Internal runtime changes

- examples(hono): add WorkPaper route smoke (d449b623)
- docs(workpaper): clarify recompute boundary (414f52d1)
- chore(release): runtime packages v0.164.6 (db4eec32)
- docs(workpaper): refresh service proof version (55545acd)
- chore(release): runtime packages v0.164.7 (e2899bfc)

## 0.164.7

- Release type: patch
- Previous libraries tag: libraries-v0.164.5
- Manual override: no

## Fixes

- fix(headless): clarify workpaper proof limits (5d4f7a2b)

## Internal runtime changes

- examples(hono): add WorkPaper route smoke (d449b623)
- docs(workpaper): clarify recompute boundary (414f52d1)
- chore(release): runtime packages v0.164.6 (db4eec32)
- docs(workpaper): refresh service proof version (55545acd)

## 0.164.6

- Release type: patch
- Previous libraries tag: libraries-v0.164.5
- Manual override: no

## Fixes

- fix(headless): clarify workpaper proof limits (5d4f7a2b)

## Internal runtime changes

- examples(hono): add WorkPaper route smoke (d449b623)
- docs(workpaper): clarify recompute boundary (414f52d1)

## 0.164.5

- Release type: patch
- Previous libraries tag: libraries-v0.164.4
- Manual override: no

## Fixes

- fix(create-workpaper): retarget starter package metadata (484c96cd)

## 0.164.4

- Release type: patch
- Previous libraries tag: libraries-v0.164.3
- Manual override: no

## Fixes

- fix(discovery): retarget unscoped runtime package metadata (76601f5b)

## Internal runtime changes

- docs(discovery): align workpaper public surfaces (aff81268)
- docs(discovery): retarget runtime package metadata (9387f24b)

## 0.164.3

- Release type: patch
- Previous libraries tag: libraries-v0.164.2
- Manual override: no

## Fixes

- fix(create-workpaper): clean starter discovery metadata (2890512f)

## 0.164.2

- Release type: patch
- Previous libraries tag: libraries-v0.164.1
- Manual override: no

## Fixes

- fix(workpaper): clean runtime discovery metadata (38445ee8)

## Internal runtime changes

- docs(growth): remove forced star copy (cf9fc858)
- docs(agent): make WorkPaper skill metadata primary (92c834a1)
- docs(growth): remove stale xlsx public route (5e0946d3)
- docs(growth): humanize feedback surfaces (8f3def07)
- docs(growth): remove stale xlsx bias (8648ab9a)
- docs(growth): clean remaining saved-file wording (5eb96f2f)
- docs(growth): humanize agent handoff surfaces (44fb158b)
- docs(growth): clean package discovery metadata (49fc5041)
- docs(growth): retarget discovery toward tool hosts (59024d3c)
- docs(growth): rebalance tool-host discovery copy (cface212)

## 0.164.1

- Release type: patch
- Previous libraries tag: libraries-v0.164.0
- Manual override: yes

## Fixes

- fix(release): publish missing runtime packages first (3af29ee5)

## 0.164.0

- Release type: minor
- Previous libraries tag: libraries-v0.163.0
- Manual override: yes

## Features

- feat(xlsx): route simple exports through @bilig/xlsx writer (0402b079)
- feat(xlsx): export bordered simple workbooks with @bilig/xlsx (f569640d)
- feat(xlsx): export sparse style artifacts with @bilig/xlsx (46e9ef1e)
- feat(xlsx): export rich text with @bilig/xlsx (223ad9df)
- feat(xlsx): write macro fixtures with @bilig/xlsx (4bf4f3a2)
- feat(xlsx): write workpaper corpus fixtures with @bilig/xlsx (3c35d222)
- feat(xlsx): read same-corpus proof cells natively (79a1ab0b)
- feat(xlsx): scan workpaper corpus natively (41fc17a9)
- feat(workpaper): add A1 facade proof API (815daaa6)
- feat(workpaper): add atomic A1 readback proof (10d5a1fc)
- feat(xlsx): add native streaming recalc engine (3fd29acb)
- feat(xlsx): use wasm row-chain native recalc (3313b566)
- feat(xlsx): add native ratio row-chain recalc (b85507f9)
- feat(xlsx): route inline ratio chains natively (310a8c00)
- feat(xlsx): evaluate direct scalar formulas natively (8c35f905)
- feat(xlsx): evaluate sum ranges natively (55c4fed0)
- feat(xlsx): evaluate row conditionals natively (cc54d056)
- feat(xlsx): add native recalc XLSX fixture corpus gate (8e421a5b)
- feat(xlsx): support native cross-sheet scalar refs (704c4430)
- feat(xlsx): evaluate exact vlookup natively (2034336b)
- feat(xlsx): support indirect native vlookup (61a3a531)
- feat(xlsx): support native sum counta ranges (2e739d01)
- feat(xlsx): support native forecast lookup formulas (915e139f)
- feat(xlsx): use native streaming cache inspection (9ef328fc)
- feat(xlsx): move cache reader to native package (4c62bc44)
- feat(xlsx): hydrate native external link caches (586854f6)
- feat(xlsx): hydrate native external companions (12e59b88)
- feat(xlsx): patch native external link artifacts (a5a740a8)
- feat(xlsx): export simple hyperlinks natively (31601f75)
- feat(xlsx): batch exact native vlookup (a68976cc)
- feat(xlsx): batch approximate native vlookup (a63e0540)
- feat(xlsx): batch typed native range aggregates (8324492b)
- feat(xlsx): report native recalc fallback status (dd95d6d9)

## Fixes

- fix(xlsx): hydrate external caches natively (231bd902)
- fix(headless): route workpaper xlsx export through @bilig/xlsx (c5759851)
- fix(excel-import): make SheetJS fallback optional (2ffc9daf)
- perf(xlsx): retain fallback source readers (ea5630f0)
- perf(xlsx): copy untouched source entries (486e813a)
- perf(xlsx): stream sync source-preserving patches (1bfd9692)
- fix(xlsx): keep workbook runtime on native paths (8ca78173)
- fix(xlsx): route imports through native reader (aef1ac9e)
- fix(xlsx): stream formula cache patches natively (7aa2f9d6)
- fix(xlsx): reduce imported workbook memory (fbbaec39)
- fix(core): defer large cached formula instances (fbc67cca)
- fix(xlsx): stream recalculated workbook output (a4ef9711)
- fix(xlsx): keep recalc zip patching internal (204bd052)
- fix(xlsx): use native async scoped recalc cli (7c90c1fd)
- fix(xlsx): use native async compatibility recalc cli (a175be4e)
- fix(xlsx): require async native file recalc cli (e95a68f4)
- fix(xlsx): support native unary text formulas (62bddca1)
- fix(xlsx-recalc): remove primary WorkPaper fallback (ed7296ab)
- fix(xlsx-recalc): split legacy WorkPaper option types (79ebf804)
- fix(workpaper): own evaluator doors outside xlsx recalc (6ca86124)
- fix(xlsx-recalc): move legacy WorkPaper path to workpaper (109a69f2)
- fix(workpaper): move xlsx risk report to native package (75d98fef)
- fix(xlsx): expand native recalc corpus coverage (4ed796f7)
- fix(xlsx): route recalc bridges through native path (ba52c1be)
- fix(xlsx): expose native workpaper file recalc (a04f16b9)
- fix(xlsx): expose native cache inspection API (f8f98c43)
- fix(xlsx): stream workbook compatibility reports (433cb297)
- fix(xlsx): bound native cache inspection defaults (eb061910)
- perf(xlsx): stream source-preserving file patches (3297613a)
- fix(xlsx): keep mcp risk scan file-backed (24768544)
- fix(headless): use file-backed xlsx mcp import (c898f07e)
- fix(xlsx): cap legacy bytes recalc path (29613fd0)
- fix(headless): keep formula clinic large files native (9dab66fa)
- fix(xlsx): cap source-preserving byte patch APIs (0fb37484)
- fix(xlsx): cap external workbook byte inputs (afde8734)
- fix(xlsx): cap SheetJS fallback source bytes (559583a3)
- fix(xlsx): cap MCP WorkPaper XLSX imports (14d5a578)
- fix(xlsx): disable sync CLI file byte input (2dc0ffd2)
- fix(xlsx): cap file source byte reads (98ae1e64)
- fix(excel-import): cap imported source byte reads (87758981)
- fix(xlsx): avoid report external workbook byte reads (d447e675)
- fix(xlsx): standardize report native diagnostics (1a4c4faf)
- fix(excel-import): block large source export fallback (c2e5b7c5)
- fix(excel-import): require explicit legacy fallback opt-in (5243ca6b)
- fix(mcp): bound xlsx risk preflight (539e2951)
- fix(xlsx): bound formula cache scans by default (9334a075)
- fix(xlsx-recalc): cap byte-buffer recalc input (4724fe75)
- fix(excel-import): cap large xlsx byte imports (80e30591)
- fix(excel-import): route large workbook dispatch through byte source (022fb2bb)
- fix(workpaper): expose xlsx recalc dependency (85d0ae6a)

## Internal runtime changes

- chore(xlsx): normalize @bilig/xlsx naming (975c7b67)
- test(xlsx): use native fixture builders (bb1dcff9)
- docs(agent): guard public rule targets (e05188d1)
- test(xlsx): expand native fixture coverage (40f6bf9e)
- test(xlsx): guard native recalc from sheetjs load (48c4060d)
- refactor(xlsx): own streaming native recalc (0c8b91e3)
- refactor(xlsx): keep native cli off workpaper path (38192a78)
- refactor(xlsx): isolate native recalc path types (c778dcd3)
- refactor(xlsx): split native recalc root from legacy workpaper (59ee0a9a)
- refactor(xlsx): keep native recalc install off headless (c0b3342e)
- refactor(xlsx): lazy load evaluator workpaper paths (588f4f49)
- refactor(xlsx): scan compatibility reports natively (96e94376)
- refactor(xlsx): keep workpaper selectors out of native package (bba963b2)
- test(excel-import): replace sheetjs lazy artifact fixture (0882493a)
- test(excel-import): replace sheetjs zip release fixture (9a629333)
- test(excel-import): replace sheetjs print fixtures (a2da8d0c)
- test(excel-import): replace sheetjs protection fixtures (682418e0)
- test(excel-import): replace sheetjs metadata fixtures (b4cccf06)
- test(excel-import): replace sheetjs conditional format fixtures (b3f8fd25)
- test(excel-import): replace sheetjs legacy comment fixtures (16463d88)
- test(excel-import): replace sheetjs ms-oi29500 fixtures (8db0a3ee)
- test(excel-import): replace sheetjs drawing fixtures (c8197db2)
- test(excel-import): replace sheetjs worksheet path fixture (49174016)
- test(excel-import): replace sheetjs generated fixtures (1852ec75)
- test(headless): replace sheetjs source-preserving fixture (c55a17d7)
- test(headless): replace sheetjs protection fixtures (bc6b4424)
- test(headless): replace sheetjs hyperlink fixture (93d5e371)
- test(headless): replace sheetjs comment fixtures (9589b770)
- test(headless): replace sheetjs external-link fixture (51bbbd13)
- test(headless): replace sheetjs chart fixture (4b92c5c0)
- test(headless): replace sheetjs preserved metadata fixture (cce302c8)
- test(excel-import): isolate sheetjs legacy fixtures (4d9e7ae9)
- test(corpus): remove sheetjs oracle fallback (f0013423)
- refactor(xlsx): share native workbook core for file scans (0ab9b0f2)
- refactor(xlsx): route native recalc through workbook core (824f3998)
- refactor(xlsx): store native recalc rows in typed arena (602aef55)
- refactor(xlsx): share workbook paths for large import preflight (1a397e2e)
- refactor(xlsx): unify source preserving patch writer (855b96b4)
- refactor(xlsx): centralize workbook report scans (c1db7bcf)
- refactor(corpus): use xlsx zip primitives (1a8e65b8)
- refactor(excel-import): use xlsx zip reader for byte sources (b0687f07)
- refactor(excel-import): use xlsx zip reader fallback (6855290d)
- refactor(excel-import): delegate zip runtime to xlsx core (4e8121a8)
- refactor(excel-import): use native sheet path resolver (5a47ec2c)
- refactor(xlsx): split native recalc source modules (6d591b9f)
- test(excel-import): route large fixtures through byte source (1e8ed360)

## 0.163.0

- Release type: patch
- Previous libraries tag: libraries-v0.162.0
- Manual override: yes

## Fixes

- fix(release): allow first-time runtime npm packages (c0196223)
- fix(xlsx): drop SheetJS CDN dependency pin (97f6e659)

## 0.162.0

- Release type: minor
- Previous libraries tag: libraries-v0.161.0
- Manual override: yes

## Features

- feat(xlsx): add @bilig/xlsx source-preserving package (75c1b926)
- feat(xlsx): expose @bilig/xlsx source patch exports (64fb24be)

## Fixes

- fix(headless): drop direct xlsx dependency (21e8e75a)
- fix(headless): build @bilig/xlsx before package checks (dbbf2c61)

## Internal runtime changes

- docs(discovery): add agent framework mcp recipes (dbf468bf)
- docs(agent): add Aider WorkPaper conventions (bafb31fe)
- docs(agent): add Goose WorkPaper MCP recipe (bd0da1d2)
- docs(agent): add xlsx risk preflight proof (5a71e298)
- docs(agent): add Roo Code WorkPaper rules (36d994e5)
- docs(agent): add Kiro WorkPaper steering (4bf0979f)
- docs(agent): add Continue WorkPaper MCP config (8fd0d04d)
- docs(agent): add Zed WorkPaper MCP setup (4ad94655)
- docs(agent): add Trae WorkPaper MCP setup (ee8001c1)
- docs(agent): add Qodo WorkPaper MCP setup (7e4028ee)
- docs(agent): add OpenHands and Goose start rules (5184b55a)

## 0.161.0

- Release type: minor
- Previous libraries tag: libraries-v0.160.2
- Manual override: no

## Features

- feat(agent): add xlsx mcp startup guidance (d92fd792)

## 0.160.2

- Release type: patch
- Previous libraries tag: libraries-v0.160.1
- Manual override: no

## Fixes

- fix(mcp): allow direct xlsx risk imports (cbf622bf)

## 0.160.1

- Release type: patch
- Previous libraries tag: libraries-v0.160.0
- Manual override: no

## Fixes

- fix(mcp): expose xlsx risk tool from package bin (19e7bd67)

## Internal runtime changes

- refactor(smoke): split xlsx risk package smoke helper (22d119ff)

## 0.160.0

- Release type: minor
- Previous libraries tag: libraries-v0.159.0
- Manual override: no

## Features

- feat(mcp): add xlsx risk preflight tool (e6cd641c)

## Fixes

- fix(ci): build xlsx runtime before workpaper (d75e7d64)

## Internal runtime changes

- ci(release): publish mcp registry inline (df861ffd)
- docs(agents): advertise compatibility evaluator (7524ee9f)
- docs(discovery): harden mcp metadata claims (96848173)

## 0.159.0

- Release type: minor
- Previous libraries tag: libraries-v0.157.0
- Manual override: no

## Features

- feat(xlsx): add workbook compatibility report (7cd5c940)

## Fixes

- fix(release): restore runtime version anchor (00a4e209)

## Internal runtime changes

- docs(n8n): improve template sticky notes (fd01b202)
- docs(agents): add openai hosted mcp proof (5a2360b2)
- docs(agent): add google adk workpaper proof (53fec737)
- docs(agent): advertise ai sdk workpaper tools (02a829b9)
- test(workpaper): copy external typecheck config (3eb87e55)
- docs(agent): add workpaper evaluator matrix (99631461)
- docs(agent): expand agent-start rule targets (aecbdcf0)
- docs(agent): expand coding agent discovery (190e48a4)
- docs(growth): add stale formula readback chooser (059d739a)
- docs(agent): add proof transcripts (5f19d493)
- docs(agent): expand proof discovery surfaces (1700b124)
- docs(growth): sharpen workbook fixture intake (5b56eb34)
- docs(growth): publish XLSX fixture corpus report (2208d89a)
- chore(release): runtime packages v0.158.0 (3d7569af)
- style(release): format generated changelog (07b28e76)

## 0.158.0

- Release type: minor
- Previous libraries tag: libraries-v0.157.0
- Manual override: no

## Features

- feat(xlsx): add workbook compatibility report (7cd5c940)

## Fixes

- fix(release): restore runtime version anchor (00a4e209)

## Internal runtime changes

- docs(n8n): improve template sticky notes (fd01b202)
- docs(agents): add openai hosted mcp proof (5a2360b2)
- docs(agent): add google adk workpaper proof (53fec737)
- docs(agent): advertise ai sdk workpaper tools (02a829b9)
- test(workpaper): copy external typecheck config (3eb87e55)
- docs(agent): add workpaper evaluator matrix (99631461)
- docs(agent): expand agent-start rule targets (aecbdcf0)
- docs(agent): expand coding agent discovery (190e48a4)
- docs(growth): add stale formula readback chooser (059d739a)
- docs(agent): add proof transcripts (5f19d493)
- docs(agent): expand proof discovery surfaces (1700b124)
- docs(growth): sharpen workbook fixture intake (5b56eb34)
- docs(growth): publish XLSX fixture corpus report (2208d89a)

## 0.157.0

- Release type: minor
- Previous libraries tag: libraries-v0.155.0
- Manual override: no

## Features

- feat(agent): emit installable WorkPaper rules (8baf3b7b)

## Internal runtime changes

- chore(release): runtime packages v0.156.0 (02d058e1)
- chore(format): satisfy markdown formatter (1d7c27a4)

## 0.156.0

- Release type: minor
- Previous libraries tag: libraries-v0.155.0
- Manual override: no

## Features

- feat(agent): emit installable WorkPaper rules (8baf3b7b)

## 0.155.0

- Release type: minor
- Previous libraries tag: libraries-v0.154.0
- Manual override: no

## Features

- feat(agent): add agent start decision CLI (fb47e3a9)

## Internal runtime changes

- docs(growth): add Agno WorkPaper MCP guide (851e67c0)
- docs(growth): add Pydantic AI WorkPaper MCP guide (540ca9b0)

## 0.154.0

- Release type: minor
- Previous libraries tag: libraries-v0.153.0
- Manual override: yes

## Internal runtime changes

- docs(workpaper): show npm proof output (cf1ae506)

## 0.153.0

- Release type: minor
- Previous libraries tag: libraries-v0.152.0
- Manual override: no

## Features

- feat(create-workpaper): add existing-repo agent handoff (40b80f38)

## 0.152.0

- Release type: minor
- Previous libraries tag: libraries-v0.151.0
- Manual override: no

## Features

- feat(formula): support QUERY labels (2b78f772)

## Internal runtime changes

- docs(growth): add Google Sheets query WorkPaper path (fef12afd)

## 0.151.0

- Release type: minor
- Previous libraries tag: libraries-v0.150.0
- Manual override: no

## Features

- feat(formula): support QUERY grouped aggregates (08c4c7e6)

## Internal runtime changes

- docs(agent): add compact WorkPaper start contract (73b1b989)

## 0.150.0

- Release type: minor
- Previous libraries tag: libraries-v0.149.0
- Manual override: no

## Features

- feat(formula): recognize Sheets import functions (3e61e0b7)

## 0.149.0

- Release type: minor
- Previous libraries tag: libraries-v0.148.0
- Manual override: no

## Features

- feat(formula): add ARRAYFORMULA (fd76cac8)

## 0.148.0

- Release type: minor
- Previous libraries tag: libraries-v0.147.0
- Manual override: no

## Features

- feat(formula): add COUNTUNIQUEIFS (db24a7a5)

## 0.147.0

- Release type: minor
- Previous libraries tag: libraries-v0.146.0
- Manual override: no

## Features

- feat(formula): add Google Sheets SORTN (660370a7)

## 0.146.0

- Release type: minor
- Previous libraries tag: libraries-v0.145.0
- Manual override: no

## Features

- feat(formula): add Google Sheets QUERY subset (6384f6f6)

## 0.145.0

- Release type: minor
- Previous libraries tag: libraries-v0.144.0
- Manual override: no

## Features

- feat(workpaper): add provider-backed agent evaluator (78ba66f9)

## 0.144.0

- Release type: minor
- Previous libraries tag: libraries-v0.143.0
- Manual override: no

## Features

- feat(workpaper): clarify agent formula adapter boundaries (30248af1)

## 0.143.0

- Release type: minor
- Previous libraries tag: libraries-v0.142.0
- Manual override: no

## Features

- feat(workpaper): add revenue plan evaluator scenario (3c84ae4f)

## 0.142.0

- Release type: minor
- Previous libraries tag: libraries-v0.141.0
- Manual override: no

## Features

- feat(workpaper): import xlsx into file-backed mcp (a41da056)

## 0.141.0

- Release type: minor
- Previous libraries tag: libraries-v0.140.1
- Manual override: no

## Features

- feat(formula): add provider-backed IMPORTRANGE (ec7be9f2)

## 0.140.1

- Release type: patch
- Previous libraries tag: libraries-v0.140.0
- Manual override: no

## Fixes

- fix(create-workpaper): align agent verifier with evaluator (ae5cb2e9)

## 0.140.0

- Release type: minor
- Previous libraries tag: libraries-v0.139.0
- Manual override: no

## Features

- feat(formula): add Google Sheets compatibility helpers (f6943122)

## 0.139.0

- Release type: minor
- Previous libraries tag: libraries-v0.138.0
- Manual override: no

## Features

- feat(formula): recognize python in excel formulas (e7c82c88)

## 0.138.0

- Release type: minor
- Previous libraries tag: libraries-v0.137.0
- Manual override: no

## Features

- feat(create-workpaper): keep overlay state under .bilig (67d3bfb4)

## 0.137.0

- Release type: minor
- Previous libraries tag: libraries-v0.136.1
- Manual override: no

## Features

- feat(create-workpaper): add Claude skill overlay (c2f487b2)

## 0.136.1

- Release type: patch
- Previous libraries tag: libraries-v0.136.0
- Manual override: no

## Fixes

- fix(formula): register copilot provider formula (56253379)

## 0.136.0

- Release type: minor
- Previous libraries tag: libraries-v0.135.0
- Manual override: no

## Features

- feat(create-workpaper): add existing repo agent overlay (f6122598)

## 0.135.0

- Release type: minor
- Previous libraries tag: libraries-v0.134.0
- Manual override: no

## Features

- feat(formula): support sheets split (65ffc449)

## 0.134.0

- Release type: minor
- Previous libraries tag: libraries-v0.133.0
- Manual override: no

## Features

- feat(formula): add regexmatch compatibility alias (89d7a30d)

## Internal runtime changes

- chore(formula): refresh formula contract snapshot (36eb1751)

## 0.133.0

- Release type: minor
- Previous libraries tag: libraries-v0.132.0
- Manual override: no

## Features

- feat(formula): track Office IS predicates (c959baff)

## 0.132.0

- Release type: minor
- Previous libraries tag: libraries-v0.131.3
- Manual override: no

## Features

- feat(formula): add protocol ids for runtime formulas (619bb403)

## 0.131.3

- Release type: patch
- Previous libraries tag: libraries-v0.131.2
- Manual override: no

## Fixes

- fix(packages): sharpen npm spreadsheet formula discovery (dbca8c70)

## Internal runtime changes

- docs(agent): add mastra workpaper smoke (e93a53cb)
- docs(fastmcp): add file-backed workpaper proof (9ba9b121)
- docs(growth): add huggingface workpaper space template (2fac9aa8)

## 0.131.2

- Release type: patch
- Previous libraries tag: libraries-v0.131.1
- Manual override: no

## Fixes

- fix(n8n): align community node verification copy (06ebcdd6)

## Internal runtime changes

- docs: tighten evaluator copy (cb4abd8b)
- docs(growth): compress npm evaluator path (c7ea04f2)
- docs(n8n): surface verified WorkPaper node (f1f69e99)

## 0.131.1

- Release type: patch
- Previous libraries tag: libraries-v0.131.0
- Manual override: no

## Fixes

- fix(release): sync cache doctor README pin (7d2e49ff)

## 0.131.0

- Release type: minor
- Previous libraries tag: libraries-v0.130.7
- Manual override: no

## Features

- feat(evaluator): add unified proof CLI (4bac9336)

## Internal runtime changes

- docs(examples): cover express workpaper route smoke (8d0d530b)

## 0.130.7

- Release type: patch
- Previous libraries tag: libraries-v0.130.6
- Manual override: no

## Fixes

- fix(headless): relocate whitespace-prefixed fill formulas (1693cb39)
- fix(ci): route oracle tests through vitest harness (5f581cbc)

## Internal runtime changes

- test(headless): cover sheet read and formula prefix edges (79e9ec23)
- docs(growth): make WorkPaper first mile headless (fa2ddbd0)
- test(headless): cover restored readback invariants (dc70e9d8)

## 0.130.6

- Release type: patch
- Previous libraries tag: libraries-v0.130.5
- Manual override: no

## Fixes

- fix(release): align cache doctor action pin (4db6886c)

## 0.130.5

- Release type: patch
- Previous libraries tag: libraries-v0.130.4
- Manual override: no

## Fixes

- fix(quality): restore static analysis gate (784c9c82)

## 0.130.4

- Release type: patch
- Previous libraries tag: libraries-v0.130.3
- Manual override: no

## Fixes

- fix(formula): align text coercion edge cases (6e587448)
- fix(formula): align choose and radix edge cases (6420e072)
- fix(formula): align omitted IF condition parity (17b62d1b)

## Internal runtime changes

- docs(site): sharpen workpaper adoption path (4a92203b)

## 0.130.3

- Release type: patch
- Previous libraries tag: libraries-v0.130.2
- Manual override: no

## Fixes

- fix(formula): align combina and sln semantics (1eceee65)
- fix(action): batch xlsx cache doctor inspection (3ce59531)

## 0.130.2

- Release type: patch
- Previous libraries tag: libraries-v0.130.1
- Manual override: no

## Fixes

- fix(packages): stabilize runtime cli bins (e92c7fc5)

## 0.130.1

- Release type: patch
- Previous libraries tag: libraries-v0.130.0
- Manual override: no

## Fixes

- fix(agent): point skill discovery at raw manifest (d8b66cb3)

## Internal runtime changes

- docs(site): sharpen cache doctor first mile (e604c75e)

## 0.130.0

- Release type: minor
- Previous libraries tag: libraries-v0.129.2
- Manual override: no

## Features

- feat(xlsx): add cache inspection API (dfc2e41b)

## 0.129.2

- Release type: patch
- Previous libraries tag: libraries-v0.129.1
- Manual override: no

## Fixes

- fix(action): align cache doctor runtime pin (9d61fbb4)
- fix(action): version cache doctor reports (baf4a887)

## 0.129.1

- Release type: patch
- Previous libraries tag: libraries-v0.129.0
- Manual override: no

## Fixes

- fix(action): pin cache doctor runtime (6488c914)

## 0.129.0

- Release type: minor
- Previous libraries tag: libraries-v0.128.0
- Manual override: no

## Features

- feat(xlsx): report cache doctor status buckets (8d05a9b7)

## 0.128.0

- Release type: minor
- Previous libraries tag: libraries-v0.127.0
- Manual override: no

## Features

- feat(xlsx): make cache doctor demo first-run check (4d9e68f2)

## Internal runtime changes

- docs(site): lead with xlsx cache doctor (db9f5780)

## 0.127.0

- Release type: minor
- Previous libraries tag: libraries-v0.126.0
- Manual override: no

## Features

- feat(action): publish cache doctor markdown report (5fa1bb9a)

## Internal runtime changes

- docs(agent): clean WorkPaper discovery copy (c00151bf)

## 0.126.0

- Release type: minor
- Previous libraries tag: libraries-v0.124.1
- Manual override: no

## Features

- feat(xlsx): scale cache doctor action to workbook globs (b1cf0324)

## Internal runtime changes

- chore(release): runtime packages v0.125.0 (12abb604)
- docs(actions): make cache doctor workflow report-first (0042bca1)

## 0.125.0

- Release type: minor
- Previous libraries tag: libraries-v0.124.1
- Manual override: no

## Features

- feat(xlsx): scale cache doctor action to workbook globs (b1cf0324)

## 0.124.1

- Release type: patch
- Previous libraries tag: libraries-v0.124.0
- Manual override: no

## Fixes

- fix(xlsx): inspect all cache doctor formulas (3a4431e5)

## Internal runtime changes

- docs(xlsx): add cache doctor CI example (f7e31090)

## 0.124.0

- Release type: minor
- Previous libraries tag: libraries-v0.123.0
- Manual override: no

## Features

- feat(xlsx): add cache doctor action (53ae8c20)

## 0.123.0

- Release type: minor
- Previous libraries tag: libraries-v0.122.0
- Manual override: no

## Features

- feat(xlsx-recalc): clarify stale readback proof json (a24fe51c)

## 0.122.0

- Release type: minor
- Previous libraries tag: libraries-v0.121.0
- Manual override: no

## Features

- feat(create-workpaper): seed agent rule starters (dcd49942)

## 0.121.0

- Release type: minor
- Previous libraries tag: libraries-v0.119.6
- Manual override: no

## Features

- feat(xlsx): add workbook formula inspection (9c4c322b)

## Fixes

- fix(release): sync Gemini extension version (bdd0cc9d)

## Internal runtime changes

- docs(agent): add copilot workpaper proof surface (c12c7e52)
- docs(growth): fix public recipe links and proof discovery (65328bf7)
- chore(release): runtime packages v0.120.0 (f4021431)

## 0.120.0

- Release type: minor
- Previous libraries tag: libraries-v0.119.6
- Manual override: no

## Features

- feat(xlsx): add workbook formula inspection (9c4c322b)

## Internal runtime changes

- docs(agent): add copilot workpaper proof surface (c12c7e52)
- docs(growth): fix public recipe links and proof discovery (65328bf7)

## 0.119.6

- Release type: patch
- Previous libraries tag: libraries-v0.119.5
- Manual override: no

## Fixes

- fix(package): keep headless keywords compressed (ae3dc5d0)

## Internal runtime changes

- docs(agent): prefer app-host skill discovery (871185d4)
- docs(growth): add Kestra WorkPaper blueprint (cf46c7d1)

## 0.119.5

- Release type: patch
- Previous libraries tag: libraries-v0.119.4
- Manual override: no

## Fixes

- fix(agent-discovery): serve skill from app host (edc11fb5)

## 0.119.4

- Release type: patch
- Previous libraries tag: libraries-v0.119.3
- Manual override: no

## Fixes

- fix(agent-discovery): publish skills-compatible index (4854ae93)

## Internal runtime changes

- docs(agent): add Claude Code WorkPaper skill (e1165a28)
- docs(agent): add Cursor and Windsurf WorkPaper rules (d034b650)
- docs(agent): add Cline and Continue WorkPaper rules (8defd5ce)
- docs(agent): add adoption kit (264d80f8)

## 0.119.3

- Release type: patch
- Previous libraries tag: libraries-v0.119.2
- Manual override: no

## Fixes

- fix(agent): align MCP challenge proof contract (fdb0ceac)

## Internal runtime changes

- docs(growth): align evaluator proof copy (27572894)

## 0.119.2

- Release type: patch
- Previous libraries tag: libraries-v0.119.1
- Manual override: no

## Fixes

- fix(docs): publish homepage recipe targets (09e61d14)

## 0.119.1

- Release type: patch
- Previous libraries tag: libraries-v0.119.0
- Manual override: no

## Fixes

- fix(ui): reject non-grid mutation screenshot proof (a5a1706c)

## Internal runtime changes

- docs(growth): compress evaluator path (fbe98a6f)
- docs(growth): surface merged Dify plugin (086c3d5b)
- docs(growth): add Pipedream WorkPaper discovery (e3a35105)

## 0.119.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)
- fix(formula): align scalar distribution domain errors (c1831db3)
- fix(excel-import): stream large xlsx workbooks by default (a40ab22a)
- fix(formula): align text and dollar edge semantics (792e5933)
- fix(core): count booleans in direct aggregate deltas (880660e0)
- fix(formula): align empty text aggregate coercion (47b12ef5)
- fix(formula): align direct logical text coercion (1e14448c)
- fix(core): preserve generated table header fill noops (5b0134ae)
- fix(formula): align empty text search bounds (65c46959)
- fix(formula): align replace append bounds (8204c341)
- fix(formula): reject negative fractional text counts (c1c65b83)
- perf(headless): reduce fresh formula binding allocations (fd69c52c)
- perf(headless): trim scalar binding hot paths (3409e125)
- fix(headless): preserve metadata rename undo source (8bfc4e29)
- fix(formula): align text coercion semantics (13459f91)
- perf(headless): cache direct count criteria aggregates (9b8bd9f1)
- fix(headless): preserve source xlsx scalar edits (48ee7a56)
- fix(formula): align scalar text coercion fast paths (9a876de8)
- fix(formula): coerce financial text arguments (ca744550)
- fix(core): preserve authored blank insert undo history (b3e44aa3)
- fix(formula): return errors for missing scalar args (1d094e4f)
- fix(formula): use natural radix width when places omitted (0fa9728f)
- fix(formula): honor text formatting bounds (f4167b66)
- fix(formula): normalize days360 february month end (3dae1f00)
- perf(headless): reduce batch write overhead (425792c9)
- perf(core): fast-path sheet rename history (711cc0d5)
- perf(core): defer dense axis indexes (5ac10fe5)
- perf(headless): classify scalar formula sources (7ddd4777)
- fix(formula): correct days360 february eom semantics (d132150f)
- fix(formula): support iso weeknum return type (25cb3104)
- fix(formula): ignore date text in timevalue (e811610f)
- fix(formula): coerce days text dates (c983a2cf)
- fix(formula): coerce datedif text dates (01c9a5a0)
- fix(formula): treat scalar rows columns as one cell (31e96220)
- fix(formula): return num for negative fractional powers (5828cefd)
- fix(xlsx): resolve cached external defined names (44c30b4d)
- fix(core): preserve external formula caches (b8090bbc)
- perf(headless): harden ironcalc benchmark hot paths (7bedb12c)
- fix(formula): support negative odd root exponentiation (e5086bad)
- fix(formula): tighten cache preservation and value semantics (c7c182ce)
- perf(headless): accelerate numeric batch hot paths (ffc8d142)
- fix(headless): avoid scratch regression in row-pair batch path (82af144a)
- fix(core): restore formula sources on insert undo (963aabfb)
- fix(core): preserve lazy column insert undo (cebdc395)
- fix(formula): return num errors for week domains (4681fbe5)
- fix(formula): return documented date domain errors (574218b1)
- fix(formula): return num for invalid yearfrac basis (837cd387)
- fix(formula): enforce rept text length limit (573dc49a)
- fix(formula): enforce text domain length limits (7a9aae14)
- fix(formula): preserve 1900 weekday compatibility (3c42b210)
- fix(formula): honor address logical ref style (c6cc31db)
- fix(formula): format address sheet prefixes (50b0d41d)
- fix(formula): parse value date text (efc31e4e)
- fix(formula): normalize overflow time text (90263bd9)
- fix(formula): honor missing if false branch (c54c349c)
- perf(headless): tighten lazy tracked index proxy (6cb4b80f)
- perf(core): speed direct scalar run binding (d3d98762)
- perf(core): lazily allocate operation appliers (58bd7a2e)
- perf(core): lazily allocate rectangular batch helpers (daa8ef3d)
- perf(core): lazily allocate runtime scratch buffers (182c290d)
- perf(headless): trim lazy batch allocation overhead (a226c998)
- perf(core): reduce formula setup p95 overhead (5005fb2e)
- perf(core): reduce direct batch column version allocations (2d0a233d)
- fix(formula): align choose and unicode text semantics (33776c60)
- fix(formula): align substitute empty old text (a496f063)
- fix(formula): make switch matching type strict (fca4ebe3)
- fix(formula): coerce date time numeric text (a0d783af)
- fix(formula): evaluate isformula from metadata (fbbf8ec0)
- fix(formula): align zero multiple rounding (dce509e4)
- fix(formula): allow lookup short result vectors (3098ba32)
- fix(formula): align lookup and error type accuracy (7abb040c)
- fix(formula): recompute stale initial scalar preevals (fea9c4cd)
- fix(formula): align error precedence and cached text (8bfa2bf1)
- fix(formula): round decimal ties after binary drift (01fea239)
- fix(formula): snap decimal multiple quotients (cf053aff)
- fix(formula): snap directed decimal rounding (08637266)
- fix(formula): coerce date text arithmetic (f6fe885d)
- fix(formula): align scalar aggregate text coercion (7e1ac601)
- fix(formula): align scalar fast-path semantics (aed37b8e)
- fix(formula): align operator precedence and bitwise arity (f9e270b4)
- fix(formula): reject malformed information arities (da7f36e0)
- fix(formula): reject malformed dollar arity (296deea0)
- fix(formula): reject surplus scalar arguments (03a720ab)
- fix(core): preserve table headers when undoing formulas (2f9b8309)
- fix(formula): preserve workday argument boundaries (31563be0)
- fix(formula): reject surplus special call arguments (03d19da2)
- fix(formula): tighten statistical arity checks (2268ba79)
- fix(formula): tighten remaining raw call arities (58c25fa7)
- fix(formula): align math domain errors (9a5a7421)
- fix(formula): tighten workday and series domains (7deed701)
- fix(formula): tighten small large data domains (4737f8c8)
- fix(formula): reject month shift underflow dates (e9644965)
- fix(formula): align empty numeric aggregate errors (ed25eb01)
- fix(formula): parse time text in time parts (a1aaba27)
- fix(formula): coerce date text in date serials (c9a63757)
- fix(formula): detect isref references before dereference (0babd288)
- fix(formula): propagate iseven isodd input errors (92435555)
- fix(formula): keep value internal spaces invalid (9954b8cf)
- fix(formula): align intl weekend and logical errors (c7e3d951)
- fix(formula): align empty text statistics coercion (bbccc9c6)
- fix(formula): align cashflow financial edge cases (ad0e6c6f)
- fix(formula): align depreciation domain errors (3103739a)
- perf(headless): beat ironcalc rust benchmark (d3e9d518)
- fix(headless): preserve cross-sheet aggregate dependencies (99d8913a)
- perf(core): index initial aggregate dependency checks (1f912d26)
- fix(release): sync benchmark evidence before publish (7a3f6a59)
- perf(headless): publish all-provider benchmark report (a1c26626)
- fix(release): install benchmark card renderer for publish (e9f1715d)
- perf(headless): publish all-provider benchmark report (94fa5896)
- fix(release): regenerate evidence after version bump (053a7134)
- fix(release): keep benchmark checks stable across version bumps (e892b3cf)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)
- chore(release): runtime packages v0.111.0 (bd89ec0d)
- chore(release): format headless changelog (f12b0a0e)
- chore(release): runtime packages v0.112.0 (db9c9b76)
- chore(headless): format release changelog (3e7326e4)
- chore(release): runtime packages v0.113.0 (4e0f39b1)
- chore(format): normalize headless changelog (6e53216f)
- chore(headless): refresh hyperformula surface (bb847655)
- test(headless): allow workpaper sheet moving (760bd719)
- test(core): expect formula source undo ops (4c759ff7)
- test(formula): cover overflow timevalue edges (a84988a7)
- test(wasm): align VALUE kernel expectations (2d686b86)
- refactor(core): split live service config types (bf77cedc)
- docs(growth): add Airbyte WorkPaper validation (d0a8d474)
- docs(growth): prove Airbyte global state validation (1db441d3)
- docs(growth): add Meltano WorkPaper utility proof (3fb0b11d)
- docs(growth): add LangChain MCP ToolNode proof (ae5aa827)
- docs(agent): harden WorkPaper skill discovery proof (857cfb55)
- test(wasm): align malformed information arity expectations (cd84bfea)
- chore(release): runtime packages v0.114.0 (79ef74f0)
- chore(headless): format changelog (86ac9c85)
- test(headless): stream excel oracle harness (1687bcbc)
- chore(release): runtime packages v0.115.0 (ddc372aa)
- chore(release): runtime packages v0.116.0 (74056776)
- chore(release): runtime packages v0.117.0 (4f807f06)
- chore(release): runtime packages v0.118.0 (63f7d319)

## 0.118.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)
- fix(formula): align scalar distribution domain errors (c1831db3)
- fix(excel-import): stream large xlsx workbooks by default (a40ab22a)
- fix(formula): align text and dollar edge semantics (792e5933)
- fix(core): count booleans in direct aggregate deltas (880660e0)
- fix(formula): align empty text aggregate coercion (47b12ef5)
- fix(formula): align direct logical text coercion (1e14448c)
- fix(core): preserve generated table header fill noops (5b0134ae)
- fix(formula): align empty text search bounds (65c46959)
- fix(formula): align replace append bounds (8204c341)
- fix(formula): reject negative fractional text counts (c1c65b83)
- perf(headless): reduce fresh formula binding allocations (fd69c52c)
- perf(headless): trim scalar binding hot paths (3409e125)
- fix(headless): preserve metadata rename undo source (8bfc4e29)
- fix(formula): align text coercion semantics (13459f91)
- perf(headless): cache direct count criteria aggregates (9b8bd9f1)
- fix(headless): preserve source xlsx scalar edits (48ee7a56)
- fix(formula): align scalar text coercion fast paths (9a876de8)
- fix(formula): coerce financial text arguments (ca744550)
- fix(core): preserve authored blank insert undo history (b3e44aa3)
- fix(formula): return errors for missing scalar args (1d094e4f)
- fix(formula): use natural radix width when places omitted (0fa9728f)
- fix(formula): honor text formatting bounds (f4167b66)
- fix(formula): normalize days360 february month end (3dae1f00)
- perf(headless): reduce batch write overhead (425792c9)
- perf(core): fast-path sheet rename history (711cc0d5)
- perf(core): defer dense axis indexes (5ac10fe5)
- perf(headless): classify scalar formula sources (7ddd4777)
- fix(formula): correct days360 february eom semantics (d132150f)
- fix(formula): support iso weeknum return type (25cb3104)
- fix(formula): ignore date text in timevalue (e811610f)
- fix(formula): coerce days text dates (c983a2cf)
- fix(formula): coerce datedif text dates (01c9a5a0)
- fix(formula): treat scalar rows columns as one cell (31e96220)
- fix(formula): return num for negative fractional powers (5828cefd)
- fix(xlsx): resolve cached external defined names (44c30b4d)
- fix(core): preserve external formula caches (b8090bbc)
- perf(headless): harden ironcalc benchmark hot paths (7bedb12c)
- fix(formula): support negative odd root exponentiation (e5086bad)
- fix(formula): tighten cache preservation and value semantics (c7c182ce)
- perf(headless): accelerate numeric batch hot paths (ffc8d142)
- fix(headless): avoid scratch regression in row-pair batch path (82af144a)
- fix(core): restore formula sources on insert undo (963aabfb)
- fix(core): preserve lazy column insert undo (cebdc395)
- fix(formula): return num errors for week domains (4681fbe5)
- fix(formula): return documented date domain errors (574218b1)
- fix(formula): return num for invalid yearfrac basis (837cd387)
- fix(formula): enforce rept text length limit (573dc49a)
- fix(formula): enforce text domain length limits (7a9aae14)
- fix(formula): preserve 1900 weekday compatibility (3c42b210)
- fix(formula): honor address logical ref style (c6cc31db)
- fix(formula): format address sheet prefixes (50b0d41d)
- fix(formula): parse value date text (efc31e4e)
- fix(formula): normalize overflow time text (90263bd9)
- fix(formula): honor missing if false branch (c54c349c)
- perf(headless): tighten lazy tracked index proxy (6cb4b80f)
- perf(core): speed direct scalar run binding (d3d98762)
- perf(core): lazily allocate operation appliers (58bd7a2e)
- perf(core): lazily allocate rectangular batch helpers (daa8ef3d)
- perf(core): lazily allocate runtime scratch buffers (182c290d)
- perf(headless): trim lazy batch allocation overhead (a226c998)
- perf(core): reduce formula setup p95 overhead (5005fb2e)
- perf(core): reduce direct batch column version allocations (2d0a233d)
- fix(formula): align choose and unicode text semantics (33776c60)
- fix(formula): align substitute empty old text (a496f063)
- fix(formula): make switch matching type strict (fca4ebe3)
- fix(formula): coerce date time numeric text (a0d783af)
- fix(formula): evaluate isformula from metadata (fbbf8ec0)
- fix(formula): align zero multiple rounding (dce509e4)
- fix(formula): allow lookup short result vectors (3098ba32)
- fix(formula): align lookup and error type accuracy (7abb040c)
- fix(formula): recompute stale initial scalar preevals (fea9c4cd)
- fix(formula): align error precedence and cached text (8bfa2bf1)
- fix(formula): round decimal ties after binary drift (01fea239)
- fix(formula): snap decimal multiple quotients (cf053aff)
- fix(formula): snap directed decimal rounding (08637266)
- fix(formula): coerce date text arithmetic (f6fe885d)
- fix(formula): align scalar aggregate text coercion (7e1ac601)
- fix(formula): align scalar fast-path semantics (aed37b8e)
- fix(formula): align operator precedence and bitwise arity (f9e270b4)
- fix(formula): reject malformed information arities (da7f36e0)
- fix(formula): reject malformed dollar arity (296deea0)
- fix(formula): reject surplus scalar arguments (03a720ab)
- fix(core): preserve table headers when undoing formulas (2f9b8309)
- fix(formula): preserve workday argument boundaries (31563be0)
- fix(formula): reject surplus special call arguments (03d19da2)
- fix(formula): tighten statistical arity checks (2268ba79)
- fix(formula): tighten remaining raw call arities (58c25fa7)
- fix(formula): align math domain errors (9a5a7421)
- fix(formula): tighten workday and series domains (7deed701)
- fix(formula): tighten small large data domains (4737f8c8)
- fix(formula): reject month shift underflow dates (e9644965)
- fix(formula): align empty numeric aggregate errors (ed25eb01)
- fix(formula): parse time text in time parts (a1aaba27)
- fix(formula): coerce date text in date serials (c9a63757)
- fix(formula): detect isref references before dereference (0babd288)
- fix(formula): propagate iseven isodd input errors (92435555)
- fix(formula): keep value internal spaces invalid (9954b8cf)
- fix(formula): align intl weekend and logical errors (c7e3d951)
- fix(formula): align empty text statistics coercion (bbccc9c6)
- fix(formula): align cashflow financial edge cases (ad0e6c6f)
- fix(formula): align depreciation domain errors (3103739a)
- perf(headless): beat ironcalc rust benchmark (d3e9d518)
- fix(headless): preserve cross-sheet aggregate dependencies (99d8913a)
- perf(core): index initial aggregate dependency checks (1f912d26)
- fix(release): sync benchmark evidence before publish (7a3f6a59)
- perf(headless): publish all-provider benchmark report (a1c26626)
- fix(release): install benchmark card renderer for publish (e9f1715d)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)
- chore(release): runtime packages v0.111.0 (bd89ec0d)
- chore(release): format headless changelog (f12b0a0e)
- chore(release): runtime packages v0.112.0 (db9c9b76)
- chore(headless): format release changelog (3e7326e4)
- chore(release): runtime packages v0.113.0 (4e0f39b1)
- chore(format): normalize headless changelog (6e53216f)
- chore(headless): refresh hyperformula surface (bb847655)
- test(headless): allow workpaper sheet moving (760bd719)
- test(core): expect formula source undo ops (4c759ff7)
- test(formula): cover overflow timevalue edges (a84988a7)
- test(wasm): align VALUE kernel expectations (2d686b86)
- refactor(core): split live service config types (bf77cedc)
- docs(growth): add Airbyte WorkPaper validation (d0a8d474)
- docs(growth): prove Airbyte global state validation (1db441d3)
- docs(growth): add Meltano WorkPaper utility proof (3fb0b11d)
- docs(growth): add LangChain MCP ToolNode proof (ae5aa827)
- docs(agent): harden WorkPaper skill discovery proof (857cfb55)
- test(wasm): align malformed information arity expectations (cd84bfea)
- chore(release): runtime packages v0.114.0 (79ef74f0)
- chore(headless): format changelog (86ac9c85)
- test(headless): stream excel oracle harness (1687bcbc)
- chore(release): runtime packages v0.115.0 (ddc372aa)
- chore(release): runtime packages v0.116.0 (74056776)
- chore(release): runtime packages v0.117.0 (4f807f06)

## 0.117.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)
- fix(formula): align scalar distribution domain errors (c1831db3)
- fix(excel-import): stream large xlsx workbooks by default (a40ab22a)
- fix(formula): align text and dollar edge semantics (792e5933)
- fix(core): count booleans in direct aggregate deltas (880660e0)
- fix(formula): align empty text aggregate coercion (47b12ef5)
- fix(formula): align direct logical text coercion (1e14448c)
- fix(core): preserve generated table header fill noops (5b0134ae)
- fix(formula): align empty text search bounds (65c46959)
- fix(formula): align replace append bounds (8204c341)
- fix(formula): reject negative fractional text counts (c1c65b83)
- perf(headless): reduce fresh formula binding allocations (fd69c52c)
- perf(headless): trim scalar binding hot paths (3409e125)
- fix(headless): preserve metadata rename undo source (8bfc4e29)
- fix(formula): align text coercion semantics (13459f91)
- perf(headless): cache direct count criteria aggregates (9b8bd9f1)
- fix(headless): preserve source xlsx scalar edits (48ee7a56)
- fix(formula): align scalar text coercion fast paths (9a876de8)
- fix(formula): coerce financial text arguments (ca744550)
- fix(core): preserve authored blank insert undo history (b3e44aa3)
- fix(formula): return errors for missing scalar args (1d094e4f)
- fix(formula): use natural radix width when places omitted (0fa9728f)
- fix(formula): honor text formatting bounds (f4167b66)
- fix(formula): normalize days360 february month end (3dae1f00)
- perf(headless): reduce batch write overhead (425792c9)
- perf(core): fast-path sheet rename history (711cc0d5)
- perf(core): defer dense axis indexes (5ac10fe5)
- perf(headless): classify scalar formula sources (7ddd4777)
- fix(formula): correct days360 february eom semantics (d132150f)
- fix(formula): support iso weeknum return type (25cb3104)
- fix(formula): ignore date text in timevalue (e811610f)
- fix(formula): coerce days text dates (c983a2cf)
- fix(formula): coerce datedif text dates (01c9a5a0)
- fix(formula): treat scalar rows columns as one cell (31e96220)
- fix(formula): return num for negative fractional powers (5828cefd)
- fix(xlsx): resolve cached external defined names (44c30b4d)
- fix(core): preserve external formula caches (b8090bbc)
- perf(headless): harden ironcalc benchmark hot paths (7bedb12c)
- fix(formula): support negative odd root exponentiation (e5086bad)
- fix(formula): tighten cache preservation and value semantics (c7c182ce)
- perf(headless): accelerate numeric batch hot paths (ffc8d142)
- fix(headless): avoid scratch regression in row-pair batch path (82af144a)
- fix(core): restore formula sources on insert undo (963aabfb)
- fix(core): preserve lazy column insert undo (cebdc395)
- fix(formula): return num errors for week domains (4681fbe5)
- fix(formula): return documented date domain errors (574218b1)
- fix(formula): return num for invalid yearfrac basis (837cd387)
- fix(formula): enforce rept text length limit (573dc49a)
- fix(formula): enforce text domain length limits (7a9aae14)
- fix(formula): preserve 1900 weekday compatibility (3c42b210)
- fix(formula): honor address logical ref style (c6cc31db)
- fix(formula): format address sheet prefixes (50b0d41d)
- fix(formula): parse value date text (efc31e4e)
- fix(formula): normalize overflow time text (90263bd9)
- fix(formula): honor missing if false branch (c54c349c)
- perf(headless): tighten lazy tracked index proxy (6cb4b80f)
- perf(core): speed direct scalar run binding (d3d98762)
- perf(core): lazily allocate operation appliers (58bd7a2e)
- perf(core): lazily allocate rectangular batch helpers (daa8ef3d)
- perf(core): lazily allocate runtime scratch buffers (182c290d)
- perf(headless): trim lazy batch allocation overhead (a226c998)
- perf(core): reduce formula setup p95 overhead (5005fb2e)
- perf(core): reduce direct batch column version allocations (2d0a233d)
- fix(formula): align choose and unicode text semantics (33776c60)
- fix(formula): align substitute empty old text (a496f063)
- fix(formula): make switch matching type strict (fca4ebe3)
- fix(formula): coerce date time numeric text (a0d783af)
- fix(formula): evaluate isformula from metadata (fbbf8ec0)
- fix(formula): align zero multiple rounding (dce509e4)
- fix(formula): allow lookup short result vectors (3098ba32)
- fix(formula): align lookup and error type accuracy (7abb040c)
- fix(formula): recompute stale initial scalar preevals (fea9c4cd)
- fix(formula): align error precedence and cached text (8bfa2bf1)
- fix(formula): round decimal ties after binary drift (01fea239)
- fix(formula): snap decimal multiple quotients (cf053aff)
- fix(formula): snap directed decimal rounding (08637266)
- fix(formula): coerce date text arithmetic (f6fe885d)
- fix(formula): align scalar aggregate text coercion (7e1ac601)
- fix(formula): align scalar fast-path semantics (aed37b8e)
- fix(formula): align operator precedence and bitwise arity (f9e270b4)
- fix(formula): reject malformed information arities (da7f36e0)
- fix(formula): reject malformed dollar arity (296deea0)
- fix(formula): reject surplus scalar arguments (03a720ab)
- fix(core): preserve table headers when undoing formulas (2f9b8309)
- fix(formula): preserve workday argument boundaries (31563be0)
- fix(formula): reject surplus special call arguments (03d19da2)
- fix(formula): tighten statistical arity checks (2268ba79)
- fix(formula): tighten remaining raw call arities (58c25fa7)
- fix(formula): align math domain errors (9a5a7421)
- fix(formula): tighten workday and series domains (7deed701)
- fix(formula): tighten small large data domains (4737f8c8)
- fix(formula): reject month shift underflow dates (e9644965)
- fix(formula): align empty numeric aggregate errors (ed25eb01)
- fix(formula): parse time text in time parts (a1aaba27)
- fix(formula): coerce date text in date serials (c9a63757)
- fix(formula): detect isref references before dereference (0babd288)
- fix(formula): propagate iseven isodd input errors (92435555)
- fix(formula): keep value internal spaces invalid (9954b8cf)
- fix(formula): align intl weekend and logical errors (c7e3d951)
- fix(formula): align empty text statistics coercion (bbccc9c6)
- fix(formula): align cashflow financial edge cases (ad0e6c6f)
- fix(formula): align depreciation domain errors (3103739a)
- perf(headless): beat ironcalc rust benchmark (d3e9d518)
- fix(headless): preserve cross-sheet aggregate dependencies (99d8913a)
- perf(core): index initial aggregate dependency checks (1f912d26)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)
- chore(release): runtime packages v0.111.0 (bd89ec0d)
- chore(release): format headless changelog (f12b0a0e)
- chore(release): runtime packages v0.112.0 (db9c9b76)
- chore(headless): format release changelog (3e7326e4)
- chore(release): runtime packages v0.113.0 (4e0f39b1)
- chore(format): normalize headless changelog (6e53216f)
- chore(headless): refresh hyperformula surface (bb847655)
- test(headless): allow workpaper sheet moving (760bd719)
- test(core): expect formula source undo ops (4c759ff7)
- test(formula): cover overflow timevalue edges (a84988a7)
- test(wasm): align VALUE kernel expectations (2d686b86)
- refactor(core): split live service config types (bf77cedc)
- docs(growth): add Airbyte WorkPaper validation (d0a8d474)
- docs(growth): prove Airbyte global state validation (1db441d3)
- docs(growth): add Meltano WorkPaper utility proof (3fb0b11d)
- docs(growth): add LangChain MCP ToolNode proof (ae5aa827)
- docs(agent): harden WorkPaper skill discovery proof (857cfb55)
- test(wasm): align malformed information arity expectations (cd84bfea)
- chore(release): runtime packages v0.114.0 (79ef74f0)
- chore(headless): format changelog (86ac9c85)
- test(headless): stream excel oracle harness (1687bcbc)
- chore(release): runtime packages v0.115.0 (ddc372aa)
- chore(release): runtime packages v0.116.0 (74056776)

## 0.116.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)
- fix(formula): align scalar distribution domain errors (c1831db3)
- fix(excel-import): stream large xlsx workbooks by default (a40ab22a)
- fix(formula): align text and dollar edge semantics (792e5933)
- fix(core): count booleans in direct aggregate deltas (880660e0)
- fix(formula): align empty text aggregate coercion (47b12ef5)
- fix(formula): align direct logical text coercion (1e14448c)
- fix(core): preserve generated table header fill noops (5b0134ae)
- fix(formula): align empty text search bounds (65c46959)
- fix(formula): align replace append bounds (8204c341)
- fix(formula): reject negative fractional text counts (c1c65b83)
- perf(headless): reduce fresh formula binding allocations (fd69c52c)
- perf(headless): trim scalar binding hot paths (3409e125)
- fix(headless): preserve metadata rename undo source (8bfc4e29)
- fix(formula): align text coercion semantics (13459f91)
- perf(headless): cache direct count criteria aggregates (9b8bd9f1)
- fix(headless): preserve source xlsx scalar edits (48ee7a56)
- fix(formula): align scalar text coercion fast paths (9a876de8)
- fix(formula): coerce financial text arguments (ca744550)
- fix(core): preserve authored blank insert undo history (b3e44aa3)
- fix(formula): return errors for missing scalar args (1d094e4f)
- fix(formula): use natural radix width when places omitted (0fa9728f)
- fix(formula): honor text formatting bounds (f4167b66)
- fix(formula): normalize days360 february month end (3dae1f00)
- perf(headless): reduce batch write overhead (425792c9)
- perf(core): fast-path sheet rename history (711cc0d5)
- perf(core): defer dense axis indexes (5ac10fe5)
- perf(headless): classify scalar formula sources (7ddd4777)
- fix(formula): correct days360 february eom semantics (d132150f)
- fix(formula): support iso weeknum return type (25cb3104)
- fix(formula): ignore date text in timevalue (e811610f)
- fix(formula): coerce days text dates (c983a2cf)
- fix(formula): coerce datedif text dates (01c9a5a0)
- fix(formula): treat scalar rows columns as one cell (31e96220)
- fix(formula): return num for negative fractional powers (5828cefd)
- fix(xlsx): resolve cached external defined names (44c30b4d)
- fix(core): preserve external formula caches (b8090bbc)
- perf(headless): harden ironcalc benchmark hot paths (7bedb12c)
- fix(formula): support negative odd root exponentiation (e5086bad)
- fix(formula): tighten cache preservation and value semantics (c7c182ce)
- perf(headless): accelerate numeric batch hot paths (ffc8d142)
- fix(headless): avoid scratch regression in row-pair batch path (82af144a)
- fix(core): restore formula sources on insert undo (963aabfb)
- fix(core): preserve lazy column insert undo (cebdc395)
- fix(formula): return num errors for week domains (4681fbe5)
- fix(formula): return documented date domain errors (574218b1)
- fix(formula): return num for invalid yearfrac basis (837cd387)
- fix(formula): enforce rept text length limit (573dc49a)
- fix(formula): enforce text domain length limits (7a9aae14)
- fix(formula): preserve 1900 weekday compatibility (3c42b210)
- fix(formula): honor address logical ref style (c6cc31db)
- fix(formula): format address sheet prefixes (50b0d41d)
- fix(formula): parse value date text (efc31e4e)
- fix(formula): normalize overflow time text (90263bd9)
- fix(formula): honor missing if false branch (c54c349c)
- perf(headless): tighten lazy tracked index proxy (6cb4b80f)
- perf(core): speed direct scalar run binding (d3d98762)
- perf(core): lazily allocate operation appliers (58bd7a2e)
- perf(core): lazily allocate rectangular batch helpers (daa8ef3d)
- perf(core): lazily allocate runtime scratch buffers (182c290d)
- perf(headless): trim lazy batch allocation overhead (a226c998)
- perf(core): reduce formula setup p95 overhead (5005fb2e)
- perf(core): reduce direct batch column version allocations (2d0a233d)
- fix(formula): align choose and unicode text semantics (33776c60)
- fix(formula): align substitute empty old text (a496f063)
- fix(formula): make switch matching type strict (fca4ebe3)
- fix(formula): coerce date time numeric text (a0d783af)
- fix(formula): evaluate isformula from metadata (fbbf8ec0)
- fix(formula): align zero multiple rounding (dce509e4)
- fix(formula): allow lookup short result vectors (3098ba32)
- fix(formula): align lookup and error type accuracy (7abb040c)
- fix(formula): recompute stale initial scalar preevals (fea9c4cd)
- fix(formula): align error precedence and cached text (8bfa2bf1)
- fix(formula): round decimal ties after binary drift (01fea239)
- fix(formula): snap decimal multiple quotients (cf053aff)
- fix(formula): snap directed decimal rounding (08637266)
- fix(formula): coerce date text arithmetic (f6fe885d)
- fix(formula): align scalar aggregate text coercion (7e1ac601)
- fix(formula): align scalar fast-path semantics (aed37b8e)
- fix(formula): align operator precedence and bitwise arity (f9e270b4)
- fix(formula): reject malformed information arities (da7f36e0)
- fix(formula): reject malformed dollar arity (296deea0)
- fix(formula): reject surplus scalar arguments (03a720ab)
- fix(core): preserve table headers when undoing formulas (2f9b8309)
- fix(formula): preserve workday argument boundaries (31563be0)
- fix(formula): reject surplus special call arguments (03d19da2)
- fix(formula): tighten statistical arity checks (2268ba79)
- fix(formula): tighten remaining raw call arities (58c25fa7)
- fix(formula): align math domain errors (9a5a7421)
- fix(formula): tighten workday and series domains (7deed701)
- fix(formula): tighten small large data domains (4737f8c8)
- fix(formula): reject month shift underflow dates (e9644965)
- fix(formula): align empty numeric aggregate errors (ed25eb01)
- fix(formula): parse time text in time parts (a1aaba27)
- fix(formula): coerce date text in date serials (c9a63757)
- fix(formula): detect isref references before dereference (0babd288)
- fix(formula): propagate iseven isodd input errors (92435555)
- fix(formula): keep value internal spaces invalid (9954b8cf)
- fix(formula): align intl weekend and logical errors (c7e3d951)
- fix(formula): align empty text statistics coercion (bbccc9c6)
- fix(formula): align cashflow financial edge cases (ad0e6c6f)
- fix(formula): align depreciation domain errors (3103739a)
- perf(headless): beat ironcalc rust benchmark (d3e9d518)
- fix(headless): preserve cross-sheet aggregate dependencies (99d8913a)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)
- chore(release): runtime packages v0.111.0 (bd89ec0d)
- chore(release): format headless changelog (f12b0a0e)
- chore(release): runtime packages v0.112.0 (db9c9b76)
- chore(headless): format release changelog (3e7326e4)
- chore(release): runtime packages v0.113.0 (4e0f39b1)
- chore(format): normalize headless changelog (6e53216f)
- chore(headless): refresh hyperformula surface (bb847655)
- test(headless): allow workpaper sheet moving (760bd719)
- test(core): expect formula source undo ops (4c759ff7)
- test(formula): cover overflow timevalue edges (a84988a7)
- test(wasm): align VALUE kernel expectations (2d686b86)
- refactor(core): split live service config types (bf77cedc)
- docs(growth): add Airbyte WorkPaper validation (d0a8d474)
- docs(growth): prove Airbyte global state validation (1db441d3)
- docs(growth): add Meltano WorkPaper utility proof (3fb0b11d)
- docs(growth): add LangChain MCP ToolNode proof (ae5aa827)
- docs(agent): harden WorkPaper skill discovery proof (857cfb55)
- test(wasm): align malformed information arity expectations (cd84bfea)
- chore(release): runtime packages v0.114.0 (79ef74f0)
- chore(headless): format changelog (86ac9c85)
- test(headless): stream excel oracle harness (1687bcbc)
- chore(release): runtime packages v0.115.0 (ddc372aa)

## 0.115.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)
- fix(formula): align scalar distribution domain errors (c1831db3)
- fix(excel-import): stream large xlsx workbooks by default (a40ab22a)
- fix(formula): align text and dollar edge semantics (792e5933)
- fix(core): count booleans in direct aggregate deltas (880660e0)
- fix(formula): align empty text aggregate coercion (47b12ef5)
- fix(formula): align direct logical text coercion (1e14448c)
- fix(core): preserve generated table header fill noops (5b0134ae)
- fix(formula): align empty text search bounds (65c46959)
- fix(formula): align replace append bounds (8204c341)
- fix(formula): reject negative fractional text counts (c1c65b83)
- perf(headless): reduce fresh formula binding allocations (fd69c52c)
- perf(headless): trim scalar binding hot paths (3409e125)
- fix(headless): preserve metadata rename undo source (8bfc4e29)
- fix(formula): align text coercion semantics (13459f91)
- perf(headless): cache direct count criteria aggregates (9b8bd9f1)
- fix(headless): preserve source xlsx scalar edits (48ee7a56)
- fix(formula): align scalar text coercion fast paths (9a876de8)
- fix(formula): coerce financial text arguments (ca744550)
- fix(core): preserve authored blank insert undo history (b3e44aa3)
- fix(formula): return errors for missing scalar args (1d094e4f)
- fix(formula): use natural radix width when places omitted (0fa9728f)
- fix(formula): honor text formatting bounds (f4167b66)
- fix(formula): normalize days360 february month end (3dae1f00)
- perf(headless): reduce batch write overhead (425792c9)
- perf(core): fast-path sheet rename history (711cc0d5)
- perf(core): defer dense axis indexes (5ac10fe5)
- perf(headless): classify scalar formula sources (7ddd4777)
- fix(formula): correct days360 february eom semantics (d132150f)
- fix(formula): support iso weeknum return type (25cb3104)
- fix(formula): ignore date text in timevalue (e811610f)
- fix(formula): coerce days text dates (c983a2cf)
- fix(formula): coerce datedif text dates (01c9a5a0)
- fix(formula): treat scalar rows columns as one cell (31e96220)
- fix(formula): return num for negative fractional powers (5828cefd)
- fix(xlsx): resolve cached external defined names (44c30b4d)
- fix(core): preserve external formula caches (b8090bbc)
- perf(headless): harden ironcalc benchmark hot paths (7bedb12c)
- fix(formula): support negative odd root exponentiation (e5086bad)
- fix(formula): tighten cache preservation and value semantics (c7c182ce)
- perf(headless): accelerate numeric batch hot paths (ffc8d142)
- fix(headless): avoid scratch regression in row-pair batch path (82af144a)
- fix(core): restore formula sources on insert undo (963aabfb)
- fix(core): preserve lazy column insert undo (cebdc395)
- fix(formula): return num errors for week domains (4681fbe5)
- fix(formula): return documented date domain errors (574218b1)
- fix(formula): return num for invalid yearfrac basis (837cd387)
- fix(formula): enforce rept text length limit (573dc49a)
- fix(formula): enforce text domain length limits (7a9aae14)
- fix(formula): preserve 1900 weekday compatibility (3c42b210)
- fix(formula): honor address logical ref style (c6cc31db)
- fix(formula): format address sheet prefixes (50b0d41d)
- fix(formula): parse value date text (efc31e4e)
- fix(formula): normalize overflow time text (90263bd9)
- fix(formula): honor missing if false branch (c54c349c)
- perf(headless): tighten lazy tracked index proxy (6cb4b80f)
- perf(core): speed direct scalar run binding (d3d98762)
- perf(core): lazily allocate operation appliers (58bd7a2e)
- perf(core): lazily allocate rectangular batch helpers (daa8ef3d)
- perf(core): lazily allocate runtime scratch buffers (182c290d)
- perf(headless): trim lazy batch allocation overhead (a226c998)
- perf(core): reduce formula setup p95 overhead (5005fb2e)
- perf(core): reduce direct batch column version allocations (2d0a233d)
- fix(formula): align choose and unicode text semantics (33776c60)
- fix(formula): align substitute empty old text (a496f063)
- fix(formula): make switch matching type strict (fca4ebe3)
- fix(formula): coerce date time numeric text (a0d783af)
- fix(formula): evaluate isformula from metadata (fbbf8ec0)
- fix(formula): align zero multiple rounding (dce509e4)
- fix(formula): allow lookup short result vectors (3098ba32)
- fix(formula): align lookup and error type accuracy (7abb040c)
- fix(formula): recompute stale initial scalar preevals (fea9c4cd)
- fix(formula): align error precedence and cached text (8bfa2bf1)
- fix(formula): round decimal ties after binary drift (01fea239)
- fix(formula): snap decimal multiple quotients (cf053aff)
- fix(formula): snap directed decimal rounding (08637266)
- fix(formula): coerce date text arithmetic (f6fe885d)
- fix(formula): align scalar aggregate text coercion (7e1ac601)
- fix(formula): align scalar fast-path semantics (aed37b8e)
- fix(formula): align operator precedence and bitwise arity (f9e270b4)
- fix(formula): reject malformed information arities (da7f36e0)
- fix(formula): reject malformed dollar arity (296deea0)
- fix(formula): reject surplus scalar arguments (03a720ab)
- fix(core): preserve table headers when undoing formulas (2f9b8309)
- fix(formula): preserve workday argument boundaries (31563be0)
- fix(formula): reject surplus special call arguments (03d19da2)
- fix(formula): tighten statistical arity checks (2268ba79)
- fix(formula): tighten remaining raw call arities (58c25fa7)
- fix(formula): align math domain errors (9a5a7421)
- fix(formula): tighten workday and series domains (7deed701)
- fix(formula): tighten small large data domains (4737f8c8)
- fix(formula): reject month shift underflow dates (e9644965)
- fix(formula): align empty numeric aggregate errors (ed25eb01)
- fix(formula): parse time text in time parts (a1aaba27)
- fix(formula): coerce date text in date serials (c9a63757)
- fix(formula): detect isref references before dereference (0babd288)
- fix(formula): propagate iseven isodd input errors (92435555)
- fix(formula): keep value internal spaces invalid (9954b8cf)
- fix(formula): align intl weekend and logical errors (c7e3d951)
- fix(formula): align empty text statistics coercion (bbccc9c6)
- fix(formula): align cashflow financial edge cases (ad0e6c6f)
- fix(formula): align depreciation domain errors (3103739a)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)
- chore(release): runtime packages v0.111.0 (bd89ec0d)
- chore(release): format headless changelog (f12b0a0e)
- chore(release): runtime packages v0.112.0 (db9c9b76)
- chore(headless): format release changelog (3e7326e4)
- chore(release): runtime packages v0.113.0 (4e0f39b1)
- chore(format): normalize headless changelog (6e53216f)
- chore(headless): refresh hyperformula surface (bb847655)
- test(headless): allow workpaper sheet moving (760bd719)
- test(core): expect formula source undo ops (4c759ff7)
- test(formula): cover overflow timevalue edges (a84988a7)
- test(wasm): align VALUE kernel expectations (2d686b86)
- refactor(core): split live service config types (bf77cedc)
- docs(growth): add Airbyte WorkPaper validation (d0a8d474)
- docs(growth): prove Airbyte global state validation (1db441d3)
- docs(growth): add Meltano WorkPaper utility proof (3fb0b11d)
- docs(growth): add LangChain MCP ToolNode proof (ae5aa827)
- docs(agent): harden WorkPaper skill discovery proof (857cfb55)
- test(wasm): align malformed information arity expectations (cd84bfea)
- chore(release): runtime packages v0.114.0 (79ef74f0)
- chore(headless): format changelog (86ac9c85)
- test(headless): stream excel oracle harness (1687bcbc)

## 0.114.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)
- fix(formula): align scalar distribution domain errors (c1831db3)
- fix(excel-import): stream large xlsx workbooks by default (a40ab22a)
- fix(formula): align text and dollar edge semantics (792e5933)
- fix(core): count booleans in direct aggregate deltas (880660e0)
- fix(formula): align empty text aggregate coercion (47b12ef5)
- fix(formula): align direct logical text coercion (1e14448c)
- fix(core): preserve generated table header fill noops (5b0134ae)
- fix(formula): align empty text search bounds (65c46959)
- fix(formula): align replace append bounds (8204c341)
- fix(formula): reject negative fractional text counts (c1c65b83)
- perf(headless): reduce fresh formula binding allocations (fd69c52c)
- perf(headless): trim scalar binding hot paths (3409e125)
- fix(headless): preserve metadata rename undo source (8bfc4e29)
- fix(formula): align text coercion semantics (13459f91)
- perf(headless): cache direct count criteria aggregates (9b8bd9f1)
- fix(headless): preserve source xlsx scalar edits (48ee7a56)
- fix(formula): align scalar text coercion fast paths (9a876de8)
- fix(formula): coerce financial text arguments (ca744550)
- fix(core): preserve authored blank insert undo history (b3e44aa3)
- fix(formula): return errors for missing scalar args (1d094e4f)
- fix(formula): use natural radix width when places omitted (0fa9728f)
- fix(formula): honor text formatting bounds (f4167b66)
- fix(formula): normalize days360 february month end (3dae1f00)
- perf(headless): reduce batch write overhead (425792c9)
- perf(core): fast-path sheet rename history (711cc0d5)
- perf(core): defer dense axis indexes (5ac10fe5)
- perf(headless): classify scalar formula sources (7ddd4777)
- fix(formula): correct days360 february eom semantics (d132150f)
- fix(formula): support iso weeknum return type (25cb3104)
- fix(formula): ignore date text in timevalue (e811610f)
- fix(formula): coerce days text dates (c983a2cf)
- fix(formula): coerce datedif text dates (01c9a5a0)
- fix(formula): treat scalar rows columns as one cell (31e96220)
- fix(formula): return num for negative fractional powers (5828cefd)
- fix(xlsx): resolve cached external defined names (44c30b4d)
- fix(core): preserve external formula caches (b8090bbc)
- perf(headless): harden ironcalc benchmark hot paths (7bedb12c)
- fix(formula): support negative odd root exponentiation (e5086bad)
- fix(formula): tighten cache preservation and value semantics (c7c182ce)
- perf(headless): accelerate numeric batch hot paths (ffc8d142)
- fix(headless): avoid scratch regression in row-pair batch path (82af144a)
- fix(core): restore formula sources on insert undo (963aabfb)
- fix(core): preserve lazy column insert undo (cebdc395)
- fix(formula): return num errors for week domains (4681fbe5)
- fix(formula): return documented date domain errors (574218b1)
- fix(formula): return num for invalid yearfrac basis (837cd387)
- fix(formula): enforce rept text length limit (573dc49a)
- fix(formula): enforce text domain length limits (7a9aae14)
- fix(formula): preserve 1900 weekday compatibility (3c42b210)
- fix(formula): honor address logical ref style (c6cc31db)
- fix(formula): format address sheet prefixes (50b0d41d)
- fix(formula): parse value date text (efc31e4e)
- fix(formula): normalize overflow time text (90263bd9)
- fix(formula): honor missing if false branch (c54c349c)
- perf(headless): tighten lazy tracked index proxy (6cb4b80f)
- perf(core): speed direct scalar run binding (d3d98762)
- perf(core): lazily allocate operation appliers (58bd7a2e)
- perf(core): lazily allocate rectangular batch helpers (daa8ef3d)
- perf(core): lazily allocate runtime scratch buffers (182c290d)
- perf(headless): trim lazy batch allocation overhead (a226c998)
- perf(core): reduce formula setup p95 overhead (5005fb2e)
- perf(core): reduce direct batch column version allocations (2d0a233d)
- fix(formula): align choose and unicode text semantics (33776c60)
- fix(formula): align substitute empty old text (a496f063)
- fix(formula): make switch matching type strict (fca4ebe3)
- fix(formula): coerce date time numeric text (a0d783af)
- fix(formula): evaluate isformula from metadata (fbbf8ec0)
- fix(formula): align zero multiple rounding (dce509e4)
- fix(formula): allow lookup short result vectors (3098ba32)
- fix(formula): align lookup and error type accuracy (7abb040c)
- fix(formula): recompute stale initial scalar preevals (fea9c4cd)
- fix(formula): align error precedence and cached text (8bfa2bf1)
- fix(formula): round decimal ties after binary drift (01fea239)
- fix(formula): snap decimal multiple quotients (cf053aff)
- fix(formula): snap directed decimal rounding (08637266)
- fix(formula): coerce date text arithmetic (f6fe885d)
- fix(formula): align scalar aggregate text coercion (7e1ac601)
- fix(formula): align scalar fast-path semantics (aed37b8e)
- fix(formula): align operator precedence and bitwise arity (f9e270b4)
- fix(formula): reject malformed information arities (da7f36e0)
- fix(formula): reject malformed dollar arity (296deea0)
- fix(formula): reject surplus scalar arguments (03a720ab)
- fix(core): preserve table headers when undoing formulas (2f9b8309)
- fix(formula): preserve workday argument boundaries (31563be0)
- fix(formula): reject surplus special call arguments (03d19da2)
- fix(formula): tighten statistical arity checks (2268ba79)
- fix(formula): tighten remaining raw call arities (58c25fa7)
- fix(formula): align math domain errors (9a5a7421)
- fix(formula): tighten workday and series domains (7deed701)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)
- chore(release): runtime packages v0.111.0 (bd89ec0d)
- chore(release): format headless changelog (f12b0a0e)
- chore(release): runtime packages v0.112.0 (db9c9b76)
- chore(headless): format release changelog (3e7326e4)
- chore(release): runtime packages v0.113.0 (4e0f39b1)
- chore(format): normalize headless changelog (6e53216f)
- chore(headless): refresh hyperformula surface (bb847655)
- test(headless): allow workpaper sheet moving (760bd719)
- test(core): expect formula source undo ops (4c759ff7)
- test(formula): cover overflow timevalue edges (a84988a7)
- test(wasm): align VALUE kernel expectations (2d686b86)
- refactor(core): split live service config types (bf77cedc)
- docs(growth): add Airbyte WorkPaper validation (d0a8d474)
- docs(growth): prove Airbyte global state validation (1db441d3)
- docs(growth): add Meltano WorkPaper utility proof (3fb0b11d)
- docs(growth): add LangChain MCP ToolNode proof (ae5aa827)
- docs(agent): harden WorkPaper skill discovery proof (857cfb55)
- test(wasm): align malformed information arity expectations (cd84bfea)

## 0.113.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)
- fix(formula): align scalar distribution domain errors (c1831db3)
- fix(excel-import): stream large xlsx workbooks by default (a40ab22a)
- fix(formula): align text and dollar edge semantics (792e5933)
- fix(core): count booleans in direct aggregate deltas (880660e0)
- fix(formula): align empty text aggregate coercion (47b12ef5)
- fix(formula): align direct logical text coercion (1e14448c)
- fix(core): preserve generated table header fill noops (5b0134ae)
- fix(formula): align empty text search bounds (65c46959)
- fix(formula): align replace append bounds (8204c341)
- fix(formula): reject negative fractional text counts (c1c65b83)
- perf(headless): reduce fresh formula binding allocations (fd69c52c)
- perf(headless): trim scalar binding hot paths (3409e125)
- fix(headless): preserve metadata rename undo source (8bfc4e29)
- fix(formula): align text coercion semantics (13459f91)
- perf(headless): cache direct count criteria aggregates (9b8bd9f1)
- fix(headless): preserve source xlsx scalar edits (48ee7a56)
- fix(formula): align scalar text coercion fast paths (9a876de8)
- fix(formula): coerce financial text arguments (ca744550)
- fix(core): preserve authored blank insert undo history (b3e44aa3)
- fix(formula): return errors for missing scalar args (1d094e4f)
- fix(formula): use natural radix width when places omitted (0fa9728f)
- fix(formula): honor text formatting bounds (f4167b66)
- fix(formula): normalize days360 february month end (3dae1f00)
- perf(headless): reduce batch write overhead (425792c9)
- perf(core): fast-path sheet rename history (711cc0d5)
- perf(core): defer dense axis indexes (5ac10fe5)
- perf(headless): classify scalar formula sources (7ddd4777)
- fix(formula): correct days360 february eom semantics (d132150f)
- fix(formula): support iso weeknum return type (25cb3104)
- fix(formula): ignore date text in timevalue (e811610f)
- fix(formula): coerce days text dates (c983a2cf)
- fix(formula): coerce datedif text dates (01c9a5a0)
- fix(formula): treat scalar rows columns as one cell (31e96220)
- fix(formula): return num for negative fractional powers (5828cefd)
- fix(xlsx): resolve cached external defined names (44c30b4d)
- fix(core): preserve external formula caches (b8090bbc)
- perf(headless): harden ironcalc benchmark hot paths (7bedb12c)
- fix(formula): support negative odd root exponentiation (e5086bad)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)
- chore(release): runtime packages v0.111.0 (bd89ec0d)
- chore(release): format headless changelog (f12b0a0e)
- chore(release): runtime packages v0.112.0 (db9c9b76)
- chore(headless): format release changelog (3e7326e4)

## 0.112.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)
- fix(formula): align scalar distribution domain errors (c1831db3)
- fix(excel-import): stream large xlsx workbooks by default (a40ab22a)
- fix(formula): align text and dollar edge semantics (792e5933)
- fix(core): count booleans in direct aggregate deltas (880660e0)
- fix(formula): align empty text aggregate coercion (47b12ef5)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)
- chore(release): runtime packages v0.111.0 (bd89ec0d)
- chore(release): format headless changelog (f12b0a0e)

## 0.111.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)
- feat(workbook): expose resolved ref proof option (fc16d103)
- feat(mcp): add Semantic Kernel proof (f4815070)
- feat(workpaper): expose ai sdk tool helpers (57c05861)
- feat(growth): standardize proof next steps (f0243533)
- feat(mcp): add transactional WorkPaper readback tool (a4d1fcb1)
- feat(openwebui): add WorkPaper OpenAPI tool route (1044d88e)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)
- perf(headless): inline fast range value reads (c44615f5)
- perf(headless): reduce formula build overhead (dc0093f2)
- Revert "perf(headless): reduce formula build overhead" (797aadf5)
- fix(workbook): classify transported plans predictably (942abe8b)
- fix(formula): match excel FLOOR sign error semantics (07081b06)
- perf(headless): reduce formula build overhead (ff274360)
- perf(headless): defer operation runtime setup (964c82b3)
- fix(formula): return num errors for combinatoric domains (0dc5946a)
- fix(formula): return num errors for square root domains (a8c32aa5)
- fix(workbook): harden formula expression arrays (065009b5)
- fix(formula): return num errors for inverse domains (ac25b767)
- fix(workbook): harden readback proof arrays (228bc516)
- fix(formula): return num errors for bessel domains (fd268742)
- fix(workbook): require plain proof arrays (6636fabc)
- fix(workbook): require plain receipt arrays (0c674ea2)
- fix(core): skip no-op table delete history (0667e84c)
- fix(formula): return num errors for distribution domains (5c6dfc8e)
- fix(workbook): require plain plan arrays (fa23a236)
- fix(formula): return num errors for rate domains (e1c8ae8c)
- fix(workbook): require plain runtime requirement arrays (b5258bb8)
- fix(formula): align ordered statistics domain errors (fbbc2eaa)
- fix(formula): align integer math domain errors (27ae3ac0)
- fix(formula): align BASE domain errors (9307d905)
- fix(formula): align cumulative finance domain errors (d869295a)
- perf(core): bypass operation setup for scalar edits (c541e2e3)
- fix(formula): align distribution and security domain errors (3f7d06ee)
- fix(formula): align distribution and depreciation domain errors (982048ab)
- fix(formula): align radix domain errors (7967d5ac)
- fix(formula): align scalar math text and overflow errors (4c3d9367)
- perf(headless): fast-path numeric range reads (50b827f3)
- perf(core): avoid scalar replacement allocations (eb0b0013)
- revert(core): drop scalar replacement allocation shortcut (eaefc2c5)
- fix(formula): align bitwise domains (749db808)
- perf(core): use typed numeric history for bulk mutations (00788860)
- fix(formula): align text coercion domains (e0cbac98)
- fix(formula): preserve scalar math error semantics (23467c49)
- fix(ui): verify same-corpus proof archive files (10d31de0)
- fix(formula): preserve utility and statistics semantics (a7562fe5)
- perf(core): tighten dense grid and scalar closure paths (e20b40be)
- fix(formula): preserve helper error semantics (f715dc24)
- fix(core): preserve qualified defined-name ref errors (4cdefea4)
- fix(core): track sort-key structural history (0605cb6e)
- perf(core): avoid unused scalar closure buffer allocation (fe82099b)
- fix(formula): harden error precedence semantics (8c39f037)
- fix(formula): align zero-multiple rounding semantics (170b3e0e)
- fix(formula): preserve complex argument errors (e3380fe0)
- fix(formula): coerce average direct text (73c50198)
- fix(formula): preserve financial argument errors (301053f5)
- fix(formula): align radix errors and scalar index (45c7b25f)
- fix(formula): align aggregate and value semantics (7dea933c)
- fix(formula): align logical reference semantics (2d1e2ad0)
- fix(formula): align radix conversion semantics (78006279)
- fix(formula): align aggregate and text edge semantics (4940afb8)
- fix(formula): align aggregate statistical error semantics (9618746b)
- perf(headless): speed scalar formula hot paths (695a250d)
- fix(formula): align rank and gamma inverse semantics (573c71f2)
- fix(formula): align confidence interval semantics (75d33a99)
- fix(core): ignore identity-only axis metadata (f5840f10)
- fix(formula): stabilize combinatoric overflow semantics (62a44fae)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)
- chore(release): runtime packages v0.110.0 (c11c5d4d)
- docs(growth): add LangGraph ToolNode proof (f6aa22c4)
- chore(docs): format package docs (eb40be41)
- docs(agents): add open multi-agent workpaper mcp guide (5fd868a6)
- chore(docs): format bilig skill example (aa4db347)
- test(ci): stabilize full coverage assertions (53b1306e)
- test(formula): align direct max text coercion (3d57c93d)
- chore(agent): refresh discovery docs (437a700c)
- chore(docs): format bilig skill example (5aaa72c6)
- chore(agent): refresh discovery docs (d040efb2)
- chore(docs): format bilig skill example (3d9a2dd3)
- test(core): cover stale sort-key delete history (a047dfb4)
- chore(docs): format bilig skill example (41fea1e2)
- docs(growth): sharpen proof conversion path (eb03902b)
- test(formula): align js evaluator aggregate semantics (24a9dd63)
- refactor(formula): centralize tracked metadata refresh (fdca4a34)

## 0.110.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)
- perf(headless): skip redundant scalar mutation checks (276c3b2d)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)
- chore(release): runtime packages v0.109.0 (b2b814c4)

## 0.109.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)
- feat(xlsx): add external workbook proof conversion (5f4525c7)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)
- perf(core): batch dense logical axis ids (407d1d93)
- fix(formula): match excel ATAN2 coordinates (30cf27ae)
- fix(release): use workflow token for GitHub API waits (e46138a3)
- fix(docs): avoid unpublished MCPB links (fee18c69)
- fix(docs): refresh headless footprint evidence (c0098efb)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)
- chore(release): runtime packages v0.108.0 (e005aeab)
- docs(workbook): expose agent intent API (65b27cae)
- docs(agent): prefer latest package coordinate (a23cb3be)
- docs(growth): add smolagents WorkPaper proof (5f7170ba)
- docs(growth): add Inngest WorkPaper step proof (16aa5aec)

## 0.108.0

- Release type: minor
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Features

- feat(n8n): evaluate WorkPaper JSON documents (1520edf9)

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)
- fix(workbook): validate noop format effects (470d0216)
- perf(core): short-circuit sorted direct delta probes (bd5c757f)
- perf(headless): reserve initial sheet cells once (7b225d02)
- perf(core): use sparse formula init membership (cb1815cd)
- perf(core): propagate affine scalar deltas (d57e4633)
- fix(formula): match excel MOD sign semantics (7f9f8e2d)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)
- chore(release): runtime packages v0.107.16 (f1eb2738)
- docs(growth): add FastMCP WorkPaper client (148a7a21)
- docs(agent): prove OpenAI Agents MCP tools (c0ac6d47)

## 0.107.16

- Release type: patch
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)
- fix(workbook): validate noop effect descriptions (1e806a8b)
- fix(workbook): align noop proof schema (060d80d5)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)
- chore(release): runtime packages v0.107.15 (32641948)
- docs(growth): add Temporal WorkPaper activity (29de7b8a)

## 0.107.15

- Release type: patch
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)
- fix(workbook): close oracle proof blockers (7b094e5b)
- perf(headless): split constant scalar delta hot path (45e869c1)
- fix(workbook): harden run option and command schemas (82559f78)
- fix(workbook): align command bundle schema guards (2e42a437)
- fix(workbook): prove idempotent command noops (546a136d)
- perf(headless): bound scalar closure preallocation (1e1bbf1a)
- fix(workbook): require command-bound noop proof (bfab4a7b)
- fix(workbook): preserve noop run descriptions (80c09654)
- fix(workbook): bind noop proof to full ops (062719e6)
- fix(workbook): bind noop descriptions to receipts (b3f402bd)
- fix(grid): harden native text proof (e1bb2d4f)
- perf(core): reuse empty dependency arrays (4f9f8d1b)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)
- chore(release): runtime packages v0.107.14 (0561ed73)
- docs(growth): add Kestra WorkPaper flow (57e0694a)
- docs(growth): add Prefect WorkPaper flow (638893e0)
- docs(growth): add Airflow WorkPaper DAG (969b7dc3)
- docs(growth): add Dagster WorkPaper asset (b9caf0f3)
- test(workbook): remove business-specific labels (e0fcb106)

## 0.107.14

- Release type: patch
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)
- perf(headless): reduce metadata rename versioning overhead (3f40baa7)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)
- chore(release): runtime packages v0.107.13 (caf12943)

## 0.107.13

- Release type: patch
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)
- fix(workbook): isolate feature plugin surface (55e43e6b)
- fix(workbook): align schema integer bounds (1146ed8e)
- fix(workbook): tighten transport schemas (f71695be)
- fix(workbook): harden transported op hydration (579f7e13)
- fix(ui): rank same-corpus speed gaps (0855dc45)
- fix(release): skip stale metadata mutation (39deee59)
- fix(workbook): reject custom prototype transport data (97885c44)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)
- chore(release): runtime packages v0.107.12 (d17ccd46)
- docs(workbook): clarify feature command imports (c74f2c78)

## 0.107.12

- Release type: patch
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)
- fix(workbook): avoid inherited ref getters (fc119be7)
- fix(release): refresh IronCalc evidence during runtime release (799e4967)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)
- chore(release): runtime packages v0.107.11 (ef648a78)

## 0.107.11

- Release type: patch
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)
- fix(workbook): align comparison schemas (e158c4a7)
- fix(workbook): reject empty format intent (15d15cc5)
- perf(headless): fast-path suspended literal queueing (6f88495d)
- fix(workbook): require full command proof (0d406311)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)
- docs(growth): add Trigger.dev WorkPaper task (98ed4839)
- chore(release): runtime packages v0.107.10 (16999484)
- docs(headless): format release package docs (29e558d1)

## 0.107.10

- Release type: patch
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)
- fix(workbook): avoid inherited plan accessors (d280def0)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)
- docs(growth): add Windmill WorkPaper script (479a394f)
- chore(release): runtime packages v0.107.9 (383e2bd2)

## 0.107.9

- Release type: patch
- Previous libraries tag: libraries-v0.107.8
- Manual override: no

## Fixes

- fix(workbook): close strict format proof gaps (2da279d3)

## Internal runtime changes

- docs(growth): add Sim MCP setup (e353fb02)

## 0.107.8

- Release type: patch
- Previous libraries tag: libraries-v0.107.3
- Manual override: no

## Fixes

- fix(workbook): bind format receipts to ranges (bd36d30e)
- fix(workbook): bind format receipts to payloads (2b98c0b4)
- fix(workbook): prove op command receipts (87c067c8)
- fix(workbook): harden public proof contracts (147f8f7f)
- fix(workbook): prove all rendered mutation ranges (33316b6a)

## Internal runtime changes

- test(core): align formula replacement counters (dc369fcf)
- docs(growth): add LobeHub MCP setup (24672255)
- chore(release): runtime packages v0.107.4 (0948540d)
- docs(growth): add AnythingLLM MCP setup (45ffc05e)
- chore(release): runtime packages v0.107.5 (d5c5e151)
- test(workbook): include guard fuzz coverage (112f4fa2)
- chore(release): runtime packages v0.107.6 (3f51ca29)
- chore(release): runtime packages v0.107.7 (afde4bee)
- test(workbook): bind op schema to source union (66e77329)

## 0.107.7

- Release type: patch
- Previous libraries tag: libraries-v0.107.3
- Manual override: no

## Fixes

- fix(workbook): bind format receipts to ranges (bd36d30e)
- fix(workbook): bind format receipts to payloads (2b98c0b4)
- fix(workbook): prove op command receipts (87c067c8)
- fix(workbook): harden public proof contracts (147f8f7f)

## Internal runtime changes

- test(core): align formula replacement counters (dc369fcf)
- docs(growth): add LobeHub MCP setup (24672255)
- chore(release): runtime packages v0.107.4 (0948540d)
- docs(growth): add AnythingLLM MCP setup (45ffc05e)
- chore(release): runtime packages v0.107.5 (d5c5e151)
- test(workbook): include guard fuzz coverage (112f4fa2)
- chore(release): runtime packages v0.107.6 (3f51ca29)

## 0.107.6

- Release type: patch
- Previous libraries tag: libraries-v0.107.3
- Manual override: no

## Fixes

- fix(workbook): bind format receipts to ranges (bd36d30e)
- fix(workbook): bind format receipts to payloads (2b98c0b4)
- fix(workbook): prove op command receipts (87c067c8)

## Internal runtime changes

- test(core): align formula replacement counters (dc369fcf)
- docs(growth): add LobeHub MCP setup (24672255)
- chore(release): runtime packages v0.107.4 (0948540d)
- docs(growth): add AnythingLLM MCP setup (45ffc05e)
- chore(release): runtime packages v0.107.5 (d5c5e151)
- test(workbook): include guard fuzz coverage (112f4fa2)

## 0.107.5

- Release type: patch
- Previous libraries tag: libraries-v0.107.3
- Manual override: no

## Fixes

- fix(workbook): bind format receipts to ranges (bd36d30e)
- fix(workbook): bind format receipts to payloads (2b98c0b4)
- fix(workbook): prove op command receipts (87c067c8)

## Internal runtime changes

- test(core): align formula replacement counters (dc369fcf)
- docs(growth): add LobeHub MCP setup (24672255)
- chore(release): runtime packages v0.107.4 (0948540d)
- docs(growth): add AnythingLLM MCP setup (45ffc05e)

## 0.107.4

- Release type: patch
- Previous libraries tag: libraries-v0.107.3
- Manual override: no

## Fixes

- fix(workbook): bind format receipts to ranges (bd36d30e)

## Internal runtime changes

- test(core): align formula replacement counters (dc369fcf)
- docs(growth): add LobeHub MCP setup (24672255)

## 0.107.3

- Release type: patch
- Previous libraries tag: libraries-v0.107.2
- Manual override: no

## Fixes

- perf(headless): defer dense logical indexes (3a79e475)

## 0.107.2

- Release type: patch
- Previous libraries tag: libraries-v0.107.1
- Manual override: no

## Fixes

- perf(headless): skip empty metadata rename rewrite (555263f8)

## 0.107.1

- Release type: patch
- Previous libraries tag: libraries-v0.107.0
- Manual override: no

## Fixes

- perf(headless): add IronCalc Rust benchmark lane (55b35ad3)

## 0.107.0

- Release type: minor
- Previous libraries tag: libraries-v0.105.0
- Manual override: no

## Features

- feat(workbook): bind receipts to resolved ranges (97f727b8)

## Fixes

- fix(workbook): verify multi-cell formula receipts (726e583e)
- fix(claims): gate Google Sheets 10x releases (d1b98256)

## Internal runtime changes

- docs(growth): add Open WebUI MCP setup (7b389b0d)
- chore(release): runtime packages v0.106.0 (1da50c18)

## 0.106.0

- Release type: minor
- Previous libraries tag: libraries-v0.105.0
- Manual override: no

## Features

- feat(workbook): bind receipts to resolved ranges (97f727b8)

## Internal runtime changes

- docs(growth): add Open WebUI MCP setup (7b389b0d)

## 0.105.0

- Release type: minor
- Previous libraries tag: libraries-v0.103.0
- Manual override: no

## Features

- feat(workbook): constrain action input metadata (03d2f83c)
- feat(workbook): publish model manifest contract (a1e6c251)

## Internal runtime changes

- docs(growth): add Directus WorkPaper flow example (4d8840e4)
- chore(release): runtime packages v0.104.0 (f429c823)

## 0.104.0

- Release type: minor
- Previous libraries tag: libraries-v0.103.0
- Manual override: no

## Features

- feat(workbook): constrain action input metadata (03d2f83c)

## Internal runtime changes

- docs(growth): add Directus WorkPaper flow example (4d8840e4)

## 0.103.0

- Release type: minor
- Previous libraries tag: libraries-v0.96.0
- Manual override: no

## Features

- feat(workbook): prepare canonical action handoff (5eddcfff)
- feat(workbook): expose prepare subpath (b12bcc94)
- feat(workbook): validate action names before planning (eaa1f78c)
- feat(workbook): specify run result proof schema (0f6632bd)
- feat(workbook): validate plan run proof handoff (7de2b939)
- feat(workbook): add adapter conformance checks (fc953a49)
- feat(workbook): add features subpath (180d0653)
- feat(workbook): publish runtime requirements schema (d59978f6)

## Fixes

- fix(excel-import): preserve null external cache errors (d2315862)
- fix(excel-oracle): serialize macos excel harness (40b4dc3a)
- fix(workbook): match Excel active tab on sheet moves (2d19224c)
- fix(core): restore structural correctness gates (9bea79ca)
- fix(ui): require same-corpus scenario fields (862c3d01)

## Internal runtime changes

- chore(release): runtime packages v0.97.0 (df8a3278)
- refactor(workbook): split run proof helpers (69d3184b)
- chore(release): runtime packages v0.98.0 (3be2efd5)
- chore(release): runtime packages v0.99.0 (91524404)
- test(xlsx): cover native external-link package refresh (4a8d82b5)
- docs(workbook): clarify agent first readme (328f04ba)
- docs(growth): add proof-first evaluator doors (85039328)
- refactor(core): split oversized preserved metadata helpers (325fe5b4)
- chore(release): runtime packages v0.100.0 (296d52f5)
- docs(workbook): clarify package identity (2a39226b)
- chore(release): runtime packages v0.101.0 (90b4c9bb)
- chore(release): runtime packages v0.102.0 (05c00b0a)

## 0.102.0

- Release type: minor
- Previous libraries tag: libraries-v0.96.0
- Manual override: no

## Features

- feat(workbook): prepare canonical action handoff (5eddcfff)
- feat(workbook): expose prepare subpath (b12bcc94)
- feat(workbook): validate action names before planning (eaa1f78c)
- feat(workbook): specify run result proof schema (0f6632bd)
- feat(workbook): validate plan run proof handoff (7de2b939)
- feat(workbook): add adapter conformance checks (fc953a49)
- feat(workbook): add features subpath (180d0653)

## Fixes

- fix(excel-import): preserve null external cache errors (d2315862)
- fix(excel-oracle): serialize macos excel harness (40b4dc3a)
- fix(workbook): match Excel active tab on sheet moves (2d19224c)
- fix(core): restore structural correctness gates (9bea79ca)

## Internal runtime changes

- chore(release): runtime packages v0.97.0 (df8a3278)
- refactor(workbook): split run proof helpers (69d3184b)
- chore(release): runtime packages v0.98.0 (3be2efd5)
- chore(release): runtime packages v0.99.0 (91524404)
- test(xlsx): cover native external-link package refresh (4a8d82b5)
- docs(workbook): clarify agent first readme (328f04ba)
- docs(growth): add proof-first evaluator doors (85039328)
- refactor(core): split oversized preserved metadata helpers (325fe5b4)
- chore(release): runtime packages v0.100.0 (296d52f5)
- docs(workbook): clarify package identity (2a39226b)
- chore(release): runtime packages v0.101.0 (90b4c9bb)

## 0.101.0

- Release type: minor
- Previous libraries tag: libraries-v0.96.0
- Manual override: no

## Features

- feat(workbook): prepare canonical action handoff (5eddcfff)
- feat(workbook): expose prepare subpath (b12bcc94)
- feat(workbook): validate action names before planning (eaa1f78c)
- feat(workbook): specify run result proof schema (0f6632bd)
- feat(workbook): validate plan run proof handoff (7de2b939)
- feat(workbook): add adapter conformance checks (fc953a49)
- feat(workbook): add features subpath (180d0653)

## Fixes

- fix(excel-import): preserve null external cache errors (d2315862)
- fix(excel-oracle): serialize macos excel harness (40b4dc3a)
- fix(workbook): match Excel active tab on sheet moves (2d19224c)

## Internal runtime changes

- chore(release): runtime packages v0.97.0 (df8a3278)
- refactor(workbook): split run proof helpers (69d3184b)
- chore(release): runtime packages v0.98.0 (3be2efd5)
- chore(release): runtime packages v0.99.0 (91524404)
- test(xlsx): cover native external-link package refresh (4a8d82b5)
- docs(workbook): clarify agent first readme (328f04ba)
- docs(growth): add proof-first evaluator doors (85039328)
- refactor(core): split oversized preserved metadata helpers (325fe5b4)
- chore(release): runtime packages v0.100.0 (296d52f5)

## 0.100.0

- Release type: minor
- Previous libraries tag: libraries-v0.96.0
- Manual override: no

## Features

- feat(workbook): prepare canonical action handoff (5eddcfff)
- feat(workbook): expose prepare subpath (b12bcc94)
- feat(workbook): validate action names before planning (eaa1f78c)
- feat(workbook): specify run result proof schema (0f6632bd)
- feat(workbook): validate plan run proof handoff (7de2b939)

## Fixes

- fix(excel-import): preserve null external cache errors (d2315862)
- fix(excel-oracle): serialize macos excel harness (40b4dc3a)
- fix(workbook): match Excel active tab on sheet moves (2d19224c)

## Internal runtime changes

- chore(release): runtime packages v0.97.0 (df8a3278)
- refactor(workbook): split run proof helpers (69d3184b)
- chore(release): runtime packages v0.98.0 (3be2efd5)
- chore(release): runtime packages v0.99.0 (91524404)
- test(xlsx): cover native external-link package refresh (4a8d82b5)
- docs(workbook): clarify agent first readme (328f04ba)
- docs(growth): add proof-first evaluator doors (85039328)
- refactor(core): split oversized preserved metadata helpers (325fe5b4)

## 0.99.0

- Release type: minor
- Previous libraries tag: libraries-v0.96.0
- Manual override: no

## Features

- feat(workbook): prepare canonical action handoff (5eddcfff)
- feat(workbook): expose prepare subpath (b12bcc94)
- feat(workbook): validate action names before planning (eaa1f78c)
- feat(workbook): specify run result proof schema (0f6632bd)

## Fixes

- fix(excel-import): preserve null external cache errors (d2315862)

## Internal runtime changes

- chore(release): runtime packages v0.97.0 (df8a3278)
- refactor(workbook): split run proof helpers (69d3184b)
- chore(release): runtime packages v0.98.0 (3be2efd5)

## 0.98.0

- Release type: minor
- Previous libraries tag: libraries-v0.96.0
- Manual override: no

## Features

- feat(workbook): prepare canonical action handoff (5eddcfff)
- feat(workbook): expose prepare subpath (b12bcc94)
- feat(workbook): validate action names before planning (eaa1f78c)

## Fixes

- fix(excel-import): preserve null external cache errors (d2315862)

## Internal runtime changes

- chore(release): runtime packages v0.97.0 (df8a3278)

## 0.97.0

- Release type: minor
- Previous libraries tag: libraries-v0.96.0
- Manual override: no

## Features

- feat(workbook): prepare canonical action handoff (5eddcfff)

## 0.96.0

- Release type: minor
- Previous libraries tag: libraries-v0.94.0
- Manual override: no

## Features

- feat(workbook): verify strict resolved refs (26a703de)
- feat(workbook): publish agent-focused subpaths (154624da)
- feat(workbook): publish contract schemas (3417c005)

## Fixes

- fix(xlsm): keep macro code names aligned on sheet changes (4947ebcb)
- fix(xlsx): classify data model connections (d3a4173f)
- fix(release): skip stale post-metadata publishes (f7455cd7)
- fix(xlsx): hydrate sparse external ranges (459fa22c)

## Internal runtime changes

- chore(release): runtime packages v0.94.1 (67db3e9f)
- chore(release): runtime packages v0.95.0 (62fabcc7)

## 0.95.0

- Release type: minor
- Previous libraries tag: libraries-v0.94.0
- Manual override: no

## Features

- feat(workbook): verify strict resolved refs (26a703de)
- feat(workbook): publish agent-focused subpaths (154624da)

## Fixes

- fix(xlsm): keep macro code names aligned on sheet changes (4947ebcb)
- fix(xlsx): classify data model connections (d3a4173f)
- fix(release): skip stale post-metadata publishes (f7455cd7)

## Internal runtime changes

- chore(release): runtime packages v0.94.1 (67db3e9f)

## 0.94.1

- Release type: patch
- Previous libraries tag: libraries-v0.94.0
- Manual override: no

## Fixes

- fix(xlsm): keep macro code names aligned on sheet changes (4947ebcb)

## 0.94.0

- Release type: minor
- Previous libraries tag: libraries-v0.93.0
- Manual override: no

## Features

- feat(workbook): prove formula readbacks with labels (6255be2a)

## Fixes

- fix(excel-import): fail closed on ambiguous companions (1270002b)

## Internal runtime changes

- refactor(excel-import): split external reference formula helpers (c12ea184)

## 0.93.0

- Release type: minor
- Previous libraries tag: libraries-v0.92.0
- Manual override: no

## Features

- feat(workbook): verify formula labels by tokens (199509e0)

## Fixes

- fix(excel-import): scope external caches on dense parse (92d4abeb)
- fix(excel-import): preserve refreshed external cache artifacts (96ea5f21)

## 0.92.0

- Release type: minor
- Previous libraries tag: libraries-v0.91.0
- Manual override: no

## Features

- feat(workbook): bind receipts to planned ops (316a9960)

## Fixes

- fix(xlsx-recalc): hydrate external link caches from companions (e45ac2ec)

## Internal runtime changes

- refactor(excel-import): split external cache helpers (e858ba86)

## 0.91.0

- Release type: minor
- Previous libraries tag: libraries-v0.90.8
- Manual override: no

## Features

- feat(workbook): require strict run proof (5ed6f099)

## 0.90.8

- Release type: patch
- Previous libraries tag: libraries-v0.90.7
- Manual override: no

## Fixes

- fix(excel): preopen linked companions before oracle targets (ff925b26)
- fix(workbook): reject custom ref data records (0c91233a)
- fix(xlsx): surface unsupported formula cache warnings (83fe836c)

## 0.90.7

- Release type: patch
- Previous libraries tag: libraries-v0.90.6
- Manual override: no

## Fixes

- fix(xlsx): reject risky defined-name cache passthrough (a52af388)

## 0.90.6

- Release type: patch
- Previous libraries tag: libraries-v0.90.5
- Manual override: no

## Fixes

- fix(xlsx): reject risky cached formula passthrough (c7ed48fb)

## 0.90.5

- Release type: patch
- Previous libraries tag: libraries-v0.90.4
- Manual override: no

## Fixes

- fix(workbook): reject custom-prototype model roots (954ad32f)

## 0.90.4

- Release type: patch
- Previous libraries tag: libraries-v0.90.0
- Manual override: no

## Fixes

- fix(xlsx): preserve table query ownership (36dc4a9a)
- fix(workbook): require undo proof data fields (9bc6df4d)
- fix(workbook): require command result data fields (8cdd388a)
- fix(xlsx): prune pivot cache sidecars on sheet delete (7f5d6b63)
- fix(workbook): validate receipt changed ranges (1b00d1a6)
- fix(xlsx): prune query tables on table delete (2e248432)
- fix(workbook): reject array-backed runtime proof (128a6b22)
- fix(xlsx): preserve native calculation save metadata (d0c02fc0)
- fix(workbook): reject array-backed runtime requirements (e493cc0c)
- fix(ui): surface authoritative render proof gaps (47eef0ea)
- fix(workbook): reject array-backed plan data (3856e93a)
- fix(xlsx): clear stale calc metadata after recalc (e4112ebd)
- fix(workbook): reject array-backed readback proof (25e9c7c4)
- fix(xlsx): require complete calc chain for cached open (a0d94cb1)
- fix(ci): include Dockerfile in runtime release paths (fdd14737)
- fix(workbook): reject array-backed model data (f7045943)
- fix(ci): refresh agent docs after footprint sync (9e34bdc8)

## Internal runtime changes

- chore(release): runtime packages v0.90.1 (1ff330dc)
- chore(release): runtime packages v0.90.2 (f4d27d75)
- chore(release): runtime packages v0.90.3 (c0aec7af)

## 0.90.3

- Release type: patch
- Previous libraries tag: libraries-v0.90.0
- Manual override: no

## Fixes

- fix(xlsx): preserve table query ownership (36dc4a9a)
- fix(workbook): require undo proof data fields (9bc6df4d)
- fix(workbook): require command result data fields (8cdd388a)
- fix(xlsx): prune pivot cache sidecars on sheet delete (7f5d6b63)
- fix(workbook): validate receipt changed ranges (1b00d1a6)
- fix(xlsx): prune query tables on table delete (2e248432)
- fix(workbook): reject array-backed runtime proof (128a6b22)
- fix(xlsx): preserve native calculation save metadata (d0c02fc0)
- fix(workbook): reject array-backed runtime requirements (e493cc0c)
- fix(ui): surface authoritative render proof gaps (47eef0ea)
- fix(workbook): reject array-backed plan data (3856e93a)
- fix(xlsx): clear stale calc metadata after recalc (e4112ebd)
- fix(workbook): reject array-backed readback proof (25e9c7c4)
- fix(xlsx): require complete calc chain for cached open (a0d94cb1)
- fix(ci): include Dockerfile in runtime release paths (fdd14737)

## Internal runtime changes

- chore(release): runtime packages v0.90.1 (1ff330dc)
- chore(release): runtime packages v0.90.2 (f4d27d75)

## 0.90.2

- Release type: patch
- Previous libraries tag: libraries-v0.90.0
- Manual override: no

## Fixes

- fix(xlsx): preserve table query ownership (36dc4a9a)
- fix(workbook): require undo proof data fields (9bc6df4d)
- fix(workbook): require command result data fields (8cdd388a)
- fix(xlsx): prune pivot cache sidecars on sheet delete (7f5d6b63)

## Internal runtime changes

- chore(release): runtime packages v0.90.1 (1ff330dc)

## 0.90.1

- Release type: patch
- Previous libraries tag: libraries-v0.90.0
- Manual override: no

## Fixes

- fix(xlsx): preserve table query ownership (36dc4a9a)
- fix(workbook): require undo proof data fields (9bc6df4d)

## 0.90.0

- Release type: minor
- Previous libraries tag: libraries-v0.86.1
- Manual override: no

## Features

- feat(workbook): check transported ref data (2a685fd7)
- feat(workbook): check live refs (998efa61)
- feat(workbook): canonicalize transported plan data (3dd51d0e)

## Fixes

- fix(xlsx): preserve legacy comment vml after sheet rename (60706d82)
- fix(workbook): ignore command envelope scratch fields (047a916d)
- fix(workbook): allow command scratch metadata (53de4c71)
- fix(workbook): canonicalize receipt ranges (35a7dc53)
- fix(engine): block protected direct cell mutations (7983cce8)

## Internal runtime changes

- chore(release): runtime packages v0.87.0 (b7833b50)
- chore(release): runtime packages v0.88.0 (ecbea8e9)
- chore(release): runtime packages v0.89.0 (6c59925a)

## 0.89.0

- Release type: minor
- Previous libraries tag: libraries-v0.86.1
- Manual override: no

## Features

- feat(workbook): check transported ref data (2a685fd7)
- feat(workbook): check live refs (998efa61)
- feat(workbook): canonicalize transported plan data (3dd51d0e)

## Fixes

- fix(xlsx): preserve legacy comment vml after sheet rename (60706d82)
- fix(workbook): ignore command envelope scratch fields (047a916d)
- fix(workbook): allow command scratch metadata (53de4c71)
- fix(workbook): canonicalize receipt ranges (35a7dc53)

## Internal runtime changes

- chore(release): runtime packages v0.87.0 (b7833b50)
- chore(release): runtime packages v0.88.0 (ecbea8e9)

## 0.88.0

- Release type: minor
- Previous libraries tag: libraries-v0.86.1
- Manual override: no

## Features

- feat(workbook): check transported ref data (2a685fd7)
- feat(workbook): check live refs (998efa61)
- feat(workbook): canonicalize transported plan data (3dd51d0e)

## Fixes

- fix(xlsx): preserve legacy comment vml after sheet rename (60706d82)

## Internal runtime changes

- chore(release): runtime packages v0.87.0 (b7833b50)

## 0.87.0

- Release type: minor
- Previous libraries tag: libraries-v0.86.1
- Manual override: no

## Features

- feat(workbook): check transported ref data (2a685fd7)

## 0.86.1

- Release type: patch
- Previous libraries tag: libraries-v0.86.0
- Manual override: no

## Fixes

- fix(core): ignore no-op workbook history entries (60d14724)

## 0.86.0

- Release type: minor
- Previous libraries tag: libraries-v0.84.0
- Manual override: no

## Features

- feat(workbook): execute transported plan data (3ca2740a)
- feat(workbook): persist plan run proof (b82cdac0)
- feat(workbook): check run result descriptions (e33d7cb6)

## Fixes

- fix(xlsx): preserve worksheet query table topology (4d92e1c2)
- fix(protection): allow unlocked protected-sheet inputs (fa8f9b5e)
- fix(xlsx): preserve power query package artifacts (81310bb0)
- fix(workbook): reject unknown result statuses (aa11e19a)

## Internal runtime changes

- refactor(workbook): keep run handoff inspectable (42a6d8a1)
- test(headless): unlock protected range oracle edit cell (c74f2cba)
- chore(release): runtime packages v0.85.0 (3c0cd4b3)

## 0.85.0

- Release type: minor
- Previous libraries tag: libraries-v0.84.0
- Manual override: no

## Features

- feat(workbook): execute transported plan data (3ca2740a)
- feat(workbook): persist plan run proof (b82cdac0)

## Fixes

- fix(xlsx): preserve worksheet query table topology (4d92e1c2)
- fix(protection): allow unlocked protected-sheet inputs (fa8f9b5e)
- fix(xlsx): preserve power query package artifacts (81310bb0)
- fix(workbook): reject unknown result statuses (aa11e19a)

## Internal runtime changes

- refactor(workbook): keep run handoff inspectable (42a6d8a1)
- test(headless): unlock protected range oracle edit cell (c74f2cba)

## 0.84.0

- Release type: minor
- Previous libraries tag: libraries-v0.78.0
- Manual override: no

## Features

- feat(workbook): freeze runtime requirements (9d2ec121)
- feat(workbook): freeze agent descriptions (a401b3d9)
- feat(workbook): freeze proof verdicts (e5298df2)
- feat(workbook): freeze handoff validators (593be6e8)
- feat(workbook): freeze run results (8bcb2f8d)
- feat(workbook): freeze planning results (0fac6f26)
- feat(workbook): freeze helper namespaces (a780256b)
- feat(workbook): canonicalize formula readback proof (95b60f47)
- feat(workbook): validate transported row selectors (324fde09)
- feat(workbook): require strict command proof (bdd25e4b)
- feat(workbook): require agent apply result proof (44119f8e)

## Fixes

- perf(headless): skip stable rebuild value materialization (6f01a160)
- fix(xlsx): prune orphaned external link caches (3abf1875)
- fix(workbook): enforce protected structure locks (5611d516)
- fix(xlsx): preserve protected range security attrs (6cc9a443)
- fix(ui): require visible operation response proof (3b2c9320)
- fix(xlsx): preserve xlsm macro package artifacts (6c6f2a6e)

## Internal runtime changes

- chore(release): runtime packages v0.79.0 (3f2b5ea0)
- chore(release): runtime packages v0.80.0 (d53af3bb)
- chore(release): runtime packages v0.81.0 (70878ea5)
- chore(release): runtime packages v0.82.0 (c53ac932)
- chore(release): runtime packages v0.83.0 (5a3be99c)

## 0.83.0

- Release type: minor
- Previous libraries tag: libraries-v0.78.0
- Manual override: no

## Features

- feat(workbook): freeze runtime requirements (9d2ec121)
- feat(workbook): freeze agent descriptions (a401b3d9)
- feat(workbook): freeze proof verdicts (e5298df2)
- feat(workbook): freeze handoff validators (593be6e8)
- feat(workbook): freeze run results (8bcb2f8d)
- feat(workbook): freeze planning results (0fac6f26)
- feat(workbook): freeze helper namespaces (a780256b)
- feat(workbook): canonicalize formula readback proof (95b60f47)
- feat(workbook): validate transported row selectors (324fde09)
- feat(workbook): require strict command proof (bdd25e4b)

## Fixes

- perf(headless): skip stable rebuild value materialization (6f01a160)
- fix(xlsx): prune orphaned external link caches (3abf1875)
- fix(workbook): enforce protected structure locks (5611d516)
- fix(xlsx): preserve protected range security attrs (6cc9a443)
- fix(ui): require visible operation response proof (3b2c9320)
- fix(xlsx): preserve xlsm macro package artifacts (6c6f2a6e)

## Internal runtime changes

- chore(release): runtime packages v0.79.0 (3f2b5ea0)
- chore(release): runtime packages v0.80.0 (d53af3bb)
- chore(release): runtime packages v0.81.0 (70878ea5)
- chore(release): runtime packages v0.82.0 (c53ac932)

## 0.82.0

- Release type: minor
- Previous libraries tag: libraries-v0.78.0
- Manual override: no

## Features

- feat(workbook): freeze runtime requirements (9d2ec121)
- feat(workbook): freeze agent descriptions (a401b3d9)
- feat(workbook): freeze proof verdicts (e5298df2)
- feat(workbook): freeze handoff validators (593be6e8)
- feat(workbook): freeze run results (8bcb2f8d)
- feat(workbook): freeze planning results (0fac6f26)
- feat(workbook): freeze helper namespaces (a780256b)
- feat(workbook): canonicalize formula readback proof (95b60f47)
- feat(workbook): validate transported row selectors (324fde09)

## Fixes

- perf(headless): skip stable rebuild value materialization (6f01a160)
- fix(xlsx): prune orphaned external link caches (3abf1875)
- fix(workbook): enforce protected structure locks (5611d516)
- fix(xlsx): preserve protected range security attrs (6cc9a443)
- fix(ui): require visible operation response proof (3b2c9320)

## Internal runtime changes

- chore(release): runtime packages v0.79.0 (3f2b5ea0)
- chore(release): runtime packages v0.80.0 (d53af3bb)
- chore(release): runtime packages v0.81.0 (70878ea5)

## 0.81.0

- Release type: minor
- Previous libraries tag: libraries-v0.78.0
- Manual override: no

## Features

- feat(workbook): freeze runtime requirements (9d2ec121)
- feat(workbook): freeze agent descriptions (a401b3d9)
- feat(workbook): freeze proof verdicts (e5298df2)
- feat(workbook): freeze handoff validators (593be6e8)
- feat(workbook): freeze run results (8bcb2f8d)
- feat(workbook): freeze planning results (0fac6f26)
- feat(workbook): freeze helper namespaces (a780256b)

## Fixes

- perf(headless): skip stable rebuild value materialization (6f01a160)
- fix(xlsx): prune orphaned external link caches (3abf1875)
- fix(workbook): enforce protected structure locks (5611d516)
- fix(xlsx): preserve protected range security attrs (6cc9a443)
- fix(ui): require visible operation response proof (3b2c9320)

## Internal runtime changes

- chore(release): runtime packages v0.79.0 (3f2b5ea0)
- chore(release): runtime packages v0.80.0 (d53af3bb)

## 0.80.0

- Release type: minor
- Previous libraries tag: libraries-v0.78.0
- Manual override: no

## Features

- feat(workbook): freeze runtime requirements (9d2ec121)
- feat(workbook): freeze agent descriptions (a401b3d9)
- feat(workbook): freeze proof verdicts (e5298df2)
- feat(workbook): freeze handoff validators (593be6e8)
- feat(workbook): freeze run results (8bcb2f8d)

## Fixes

- perf(headless): skip stable rebuild value materialization (6f01a160)
- fix(xlsx): prune orphaned external link caches (3abf1875)
- fix(workbook): enforce protected structure locks (5611d516)

## Internal runtime changes

- chore(release): runtime packages v0.79.0 (3f2b5ea0)

## 0.79.0

- Release type: minor
- Previous libraries tag: libraries-v0.78.0
- Manual override: no

## Features

- feat(workbook): freeze runtime requirements (9d2ec121)
- feat(workbook): freeze agent descriptions (a401b3d9)

## Fixes

- perf(headless): skip stable rebuild value materialization (6f01a160)
- fix(xlsx): prune orphaned external link caches (3abf1875)

## 0.78.0

- Release type: minor
- Previous libraries tag: libraries-v0.75.0
- Manual override: no

## Features

- feat(workbook): validate check helper inputs (5a5cc867)
- feat(workbook): validate formula helper inputs (a6152282)
- feat(workbook): freeze ref transport data (a3aa8b6b)
- feat(workbook): fail closed on invalid plan verification (f23ea256)
- feat(workbook): validate model verification inputs (d209467a)

## Fixes

- fix(xlsx): prune deleted pivot package artifacts (ee5ed149)
- fix(xlsx): preserve formula validation sources (d30fb42e)
- fix(workbook): preserve defined names after sheet deletes (5a44bee3)

## Internal runtime changes

- chore(release): runtime packages v0.76.0 (48bbac33)
- chore(release): runtime packages v0.77.0 (c10cbb2e)

## 0.77.0

- Release type: minor
- Previous libraries tag: libraries-v0.75.0
- Manual override: no

## Features

- feat(workbook): validate check helper inputs (5a5cc867)
- feat(workbook): validate formula helper inputs (a6152282)
- feat(workbook): freeze ref transport data (a3aa8b6b)
- feat(workbook): fail closed on invalid plan verification (f23ea256)

## Fixes

- fix(xlsx): prune deleted pivot package artifacts (ee5ed149)
- fix(xlsx): preserve formula validation sources (d30fb42e)

## Internal runtime changes

- chore(release): runtime packages v0.76.0 (48bbac33)

## 0.76.0

- Release type: minor
- Previous libraries tag: libraries-v0.75.0
- Manual override: no

## Features

- feat(workbook): validate check helper inputs (5a5cc867)
- feat(workbook): validate formula helper inputs (a6152282)

## Fixes

- fix(xlsx): prune deleted pivot package artifacts (ee5ed149)

## 0.75.0

- Release type: minor
- Previous libraries tag: libraries-v0.73.0
- Manual override: no

## Features

- feat(workbook): harden command receipt proof (d56f4488)
- feat(workbook): add strict run proof mode (a30baf0b)
- feat(workbook): enforce command result scopes (f9c1bdc1)
- feat(workbook): reject duplicate command ids (b0fc1947)
- feat(workbook): reject ambiguous accepted results (bb9f6778)

## Fixes

- fix(xlsx): prune threaded comments on sheet delete (4ba0445e)
- fix(xlsx): keep inserted row cells ordered (d299caf5)
- fix(xlsx): invalidate raw chart refs on sheet delete (dce4427e)
- fix(xlsx): prune deleted chart sheet artifacts (1489ba0b)

## Internal runtime changes

- docs: sharpen first-touch WorkPaper path (a93d9e9b)
- chore(release): runtime packages v0.74.0 (650174af)

## 0.74.0

- Release type: minor
- Previous libraries tag: libraries-v0.73.0
- Manual override: no

## Features

- feat(workbook): harden command receipt proof (d56f4488)
- feat(workbook): add strict run proof mode (a30baf0b)
- feat(workbook): enforce command result scopes (f9c1bdc1)
- feat(workbook): reject duplicate command ids (b0fc1947)
- feat(workbook): reject ambiguous accepted results (bb9f6778)

## Fixes

- fix(xlsx): prune threaded comments on sheet delete (4ba0445e)
- fix(xlsx): keep inserted row cells ordered (d299caf5)
- fix(xlsx): invalidate raw chart refs on sheet delete (dce4427e)

## Internal runtime changes

- docs: sharpen first-touch WorkPaper path (a93d9e9b)

## 0.73.0

- Release type: minor
- Previous libraries tag: libraries-v0.71.0
- Manual override: no

## Features

- feat(workbook): bind agent runtime proof (6d3d81cc)
- feat(workbook): bind command apply receipts (3571060c)
- feat(workbook): require scoped command ranges (5154145c)
- feat(workbook): reject duplicate readbacks (027afd87)

## Fixes

- fix(workbook): rewrite preserved view tabs on sheet delete (572a0edb)
- fix(xlsx): preserve calc-chain sheet ids (4e137088)
- fix(ui): serve production same-corpus capture (0e79caba)
- fix(workbook): move sheet tabs like Excel (38ea9170)
- fix(xlsx): respect sheet relationship paths after tab moves (78b70177)
- fix(xlsx): skip zip path lookup for legacy XLS imports (c4f3f80d)
- fix(xlsx): import sheet metadata by workbook relationships (16168d2d)
- fix(xlsx): import worksheet policy metadata by relationships (6b9ab90a)
- fix(xlsx): import table and legacy comment parts by relationships (c59c6e61)
- fix(xlsx): prune deleted sheet slicer artifacts (59bd2ad3)
- fix(xlsx): preserve shared slicer parts on sheet delete (21cbb192)

## Internal runtime changes

- docs(growth): remove proof-output star asks (#23) (d3a0761f)
- test(headless): update HyperFormula surface snapshot (e536e81d)
- chore(release): runtime packages v0.72.0 (afc93c29)

## 0.72.0

- Release type: minor
- Previous libraries tag: libraries-v0.71.0
- Manual override: no

## Features

- feat(workbook): bind agent runtime proof (6d3d81cc)

## Fixes

- fix(workbook): rewrite preserved view tabs on sheet delete (572a0edb)
- fix(xlsx): preserve calc-chain sheet ids (4e137088)
- fix(ui): serve production same-corpus capture (0e79caba)
- fix(workbook): move sheet tabs like Excel (38ea9170)
- fix(xlsx): respect sheet relationship paths after tab moves (78b70177)
- fix(xlsx): skip zip path lookup for legacy XLS imports (c4f3f80d)
- fix(xlsx): import sheet metadata by workbook relationships (16168d2d)

## Internal runtime changes

- docs(growth): remove proof-output star asks (#23) (d3a0761f)
- test(headless): update HyperFormula surface snapshot (e536e81d)

## 0.71.0

- Release type: minor
- Previous libraries tag: libraries-v0.70.0
- Manual override: no

## Features

- feat(workbook): persist agent command proof (a1fcf8b5)

## Fixes

- fix(workbook): rename raw chart package refs (38f58840)
- fix(ui): require production same-corpus captures (761dac47)

## 0.70.0

- Release type: minor
- Previous libraries tag: libraries-v0.68.0
- Manual override: no

## Features

- feat(workbook): validate agent command handoff (3249a739)
- feat(workbook): add command result proof boundary (9cf062eb)

## Fixes

- fix(workbook): rewrite preserved style artifacts (cb76bf8e)
- perf(headless): speed up scalar sheet inspection (a18a7f37)
- fix(workbook): rewrite preserved pivot artifacts (f657279b)
- fix(workbook): rewrite raw drawing anchors (2a5127dc)
- perf(core): skip recalc payloads for tracked listeners (48e3930c)

## Internal runtime changes

- chore(release): runtime packages v0.68.1 (d2d0fe4f)
- chore(release): runtime packages v0.69.0 (cd42d69f)

## 0.69.0

- Release type: minor
- Previous libraries tag: libraries-v0.68.0
- Manual override: no

## Features

- feat(workbook): validate agent command handoff (3249a739)

## Fixes

- fix(workbook): rewrite preserved style artifacts (cb76bf8e)
- perf(headless): speed up scalar sheet inspection (a18a7f37)
- fix(workbook): rewrite preserved pivot artifacts (f657279b)

## Internal runtime changes

- chore(release): runtime packages v0.68.1 (d2d0fe4f)

## 0.68.1

- Release type: patch
- Previous libraries tag: libraries-v0.68.0
- Manual override: no

## Fixes

- fix(workbook): rewrite preserved style artifacts (cb76bf8e)

## 0.68.0

- Release type: minor
- Previous libraries tag: libraries-v0.67.15
- Manual override: no

## Features

- feat(workbook): add command bundle validator (044c0177)

## Fixes

- perf(core): reuse prefix aggregate templates (43253308)
- fix(workbook): preserve control artifacts (cbec564a)
- fix(workbook): preserve imported package metadata (3862603f)

## 0.67.15

- Release type: patch
- Previous libraries tag: libraries-v0.67.14
- Manual override: no

## Fixes

- fix(workbook): structure model verification failures (fb51d654)
- fix(workbook): preserve cell metadata refs (0ea43dd8)

## 0.67.14

- Release type: patch
- Previous libraries tag: libraries-v0.67.13
- Manual override: no

## Fixes

- fix(workbook): harden model inspection data (7000058d)

## 0.67.13

- Release type: patch
- Previous libraries tag: libraries-v0.67.12
- Manual override: no

## Fixes

- fix(workbook): harden model manifest data (003c6164)

## 0.67.12

- Release type: patch
- Previous libraries tag: libraries-v0.67.11
- Manual override: no

## Fixes

- fix(workbook): harden readback proof data (a2bedb46)
- fix(workbook): harden returned check data (04a83841)

## 0.67.11

- Release type: patch
- Previous libraries tag: libraries-v0.67.10
- Manual override: no

## Fixes

- fix(workbook): harden description data reads (af287709)

## 0.67.10

- Release type: patch
- Previous libraries tag: libraries-v0.67.9
- Manual override: no

## Fixes

- fix(workbook): harden ref transport traversal (aff1d5a6)
- fix(workbook): harden runtime requirement arrays (42fc4528)
- fix(workbook): harden runtime evidence arrays (df0c913b)

## 0.67.9

- Release type: patch
- Previous libraries tag: libraries-v0.67.7
- Manual override: no

## Fixes

- fix(workbook): harden runtime proof data (0ea4d8c6)
- fix(workbook): harden low-level op guards (19c8f39c)

## Internal runtime changes

- chore(release): runtime packages v0.67.8 (9d5e1795)

## 0.67.8

- Release type: patch
- Previous libraries tag: libraries-v0.67.6
- Manual override: no

## Fixes

- fix(workbook): harden feature receipt proof data (1d4db2db)
- fix(workbook): harden runtime proof data (0ea4d8c6)

## Internal runtime changes

- chore(release): runtime packages v0.67.7 (2e573247)

## 0.67.7

- Release type: patch
- Previous libraries tag: libraries-v0.67.5
- Manual override: no

## Fixes

- fix(workbook): ignore ref accessors during planning (bd84fa1e)
- fix(workbook): reject accessor-backed action input (36350981)
- fix(workbook): harden feature receipt proof data (1d4db2db)

## Internal runtime changes

- chore(release): runtime packages v0.67.6 (03344975)

## 0.67.6

- Release type: patch
- Previous libraries tag: libraries-v0.67.5
- Manual override: no

## Fixes

- fix(workbook): ignore ref accessors during planning (bd84fa1e)
- fix(workbook): reject accessor-backed action input (36350981)

## 0.67.5

- Release type: patch
- Previous libraries tag: libraries-v0.67.4
- Manual override: no

## Fixes

- fix(workbook): ignore inherited model actions (312ade88)

## 0.67.4

- Release type: patch
- Previous libraries tag: libraries-v0.67.3
- Manual override: no

## Fixes

- fix(workbook): canonicalize feature receipt proof (a6d031dd)

## 0.67.3

- Release type: patch
- Previous libraries tag: libraries-v0.67.2
- Manual override: no

## Fixes

- perf(core): reduce fresh scalar initialization overhead (be931d09)

## 0.67.2

- Release type: patch
- Previous libraries tag: libraries-v0.67.1
- Manual override: no

## Fixes

- fix(workbook): require own runtime proof fields (275d431f)
- fix(core): prevent stale structural undo styles (32a80ba5)

## 0.67.1

- Release type: patch
- Previous libraries tag: libraries-v0.67.0
- Manual override: no

## Fixes

- fix(workbook): require own transport payload fields (2b300cad)
- fix(workbook): require own feature payload fields (d5e94db0)

## 0.67.0

- Release type: minor
- Previous libraries tag: libraries-v0.66.0
- Manual override: no

## Features

- feat(workbook): preserve feature handoff issue paths (a25a9cee)

## Fixes

- perf(core): batch initial direct scalar bindings (50784028)

## 0.66.0

- Release type: minor
- Previous libraries tag: libraries-v0.64.0
- Manual override: no

## Features

- feat(workbook): validate feature plugins (f625484d)
- feat(workbook): preserve run error issue paths (3c5e911a)

## Fixes

- fix(workbook): preserve rich text artifacts (a07c402a)
- perf(core): defer deleted literal undo capture (ae1586c1)

## Internal runtime changes

- docs(workpaper): expose agent workflow entrypoints (8f0848a1)
- chore(release): runtime packages v0.65.0 (580bf991)

## 0.65.0

- Release type: minor
- Previous libraries tag: libraries-v0.64.0
- Manual override: no

## Features

- feat(workbook): validate feature plugins (f625484d)

## Fixes

- fix(workbook): preserve rich text artifacts (a07c402a)

## 0.64.0

- Release type: minor
- Previous libraries tag: libraries-v0.62.0
- Manual override: no

## Features

- feat(workbook): report plan data validation issues (359cc188)
- feat(workbook): validate runtime requirements (d6a6d05c)
- feat(workpaper): add local n8n formula server (b18fab04)
- feat(workbook): validate command requests (43f76b31)
- feat(workbook): validate command receipts (42ae649a)

## Fixes

- fix(workbook): rewrite cross-sheet conditional formats (d341ce28)
- fix(workbook): rewrite x14 conditional format ranges (17248b55)
- fix(workpaper): use indexed env access (6e4d8e53)
- perf(core): reduce direct scalar dependency allocations (635b96bd)
- fix(workbook): rewrite cross-sheet sparkline refs (edbc7566)
- fix(workbook): rewrite ignored error refs (bdd53ec3)

## Internal runtime changes

- chore(release): runtime packages v0.63.0 (4d242ce0)

## 0.63.0

- Release type: minor
- Previous libraries tag: libraries-v0.62.0
- Manual override: no

## Features

- feat(workbook): report plan data validation issues (359cc188)
- feat(workbook): validate runtime requirements (d6a6d05c)
- feat(workpaper): add local n8n formula server (b18fab04)

## Fixes

- fix(workbook): rewrite cross-sheet conditional formats (d341ce28)
- fix(workbook): rewrite x14 conditional format ranges (17248b55)
- fix(workpaper): use indexed env access (6e4d8e53)
- perf(core): reduce direct scalar dependency allocations (635b96bd)

## 0.62.0

- Release type: minor
- Previous libraries tag: libraries-v0.59.0
- Manual override: no

## Features

- feat(engine): refresh source-backed pivot caches (229ff9e5)
- feat(engine): preserve imported drawing artifacts (9d9a24e4)
- feat(workbook): add transport-safe refs (05cdf0f4)
- feat(charts): preserve Excel drawing anchors (34ca1ac3)
- feat(comments): preserve threaded comment artifacts (640df570)
- feat(workbook): expose formula labels in plans (f770ee89)
- feat(workbook): add feature kernel tables slice (1e7d1ac2)
- feat(workbook): preserve external link artifacts (b6a2361e)
- feat(workbook): run transported plan data (137ba1fd)
- feat(workbook): check action inputs before planning (ff13c823)
- feat(workbook): preserve sparkline artifacts (5d7f2064)
- feat(workbook): check runtime adapter capabilities (d00cd647)

## Fixes

- fix(corpus): keep threaded comment fixture out of parity sweep (46a1a7b7)
- fix(formula): skip blank keys in approximate lookups (9c95b946)
- fix(workbook): rewrite prefixed conditional format artifacts (257fde05)
- fix(formula): build project references (2f476f04)
- fix(core): build project references (d69a2e22)
- fix(workbook): honor optional action input metadata (c55399ea)

## Internal runtime changes

- test(bench): parallelize competitive benchmark generation (942d1782)
- test(excel): align pivot export cache assertion (d59c936c)
- chore(release): runtime packages v0.60.0 (8835ff32)
- refactor(core): split workbook metadata drawing service (377e94a4)
- refactor(core): split chart anchor metadata rewrite (266825fb)
- test(excel): expect chart anchors in xlsx roundtrip (ecf549a9)
- chore(release): runtime packages v0.61.0 (748c3cd8)
- refactor(core): split metadata artifact service methods (2c6a10d7)
- refactor(core): split metadata cell record service (a0afa5c0)
- test(bench): add Univer headless comparison (2ed8ec46)

## 0.61.0

- Release type: minor
- Previous libraries tag: libraries-v0.59.0
- Manual override: no

## Features

- feat(engine): refresh source-backed pivot caches (229ff9e5)
- feat(engine): preserve imported drawing artifacts (9d9a24e4)
- feat(workbook): add transport-safe refs (05cdf0f4)
- feat(charts): preserve Excel drawing anchors (34ca1ac3)

## Internal runtime changes

- test(bench): parallelize competitive benchmark generation (942d1782)
- test(excel): align pivot export cache assertion (d59c936c)
- chore(release): runtime packages v0.60.0 (8835ff32)
- refactor(core): split workbook metadata drawing service (377e94a4)
- refactor(core): split chart anchor metadata rewrite (266825fb)
- test(excel): expect chart anchors in xlsx roundtrip (ecf549a9)

## 0.60.0

- Release type: minor
- Previous libraries tag: libraries-v0.59.0
- Manual override: no

## Features

- feat(engine): refresh source-backed pivot caches (229ff9e5)

## Internal runtime changes

- test(bench): parallelize competitive benchmark generation (942d1782)
- test(excel): align pivot export cache assertion (d59c936c)

## 0.59.0

- Release type: minor
- Previous libraries tag: libraries-v0.58.0
- Manual override: no

## Features

- feat(workbook): preserve failed run proof (b70bd9c9)

## Fixes

- perf(core): fast translate row-offset scalar templates (21855684)

## 0.58.0

- Release type: minor
- Previous libraries tag: libraries-v0.57.0
- Manual override: no

## Features

- feat(engine): preserve hyperlink metadata structurally (e2f7f148)

## 0.57.0

- Release type: minor
- Previous libraries tag: libraries-v0.55.0
- Manual override: no

## Features

- feat(formula): preserve native structured references (9e417452)
- feat(workbook): avoid mutating model configs (5c42661b)
- feat(workbook): require own action metadata (be1bfc42)
- feat(workbook): freeze planned action handoffs (c3ac870f)
- feat(engine): execute excel autofilters (c4994591)
- feat(engine): support precision as displayed (308415f0)
- feat(workbook): verify immutable plan refs (55859f20)
- feat(engine): preserve excel conditional format artifacts (492ebfd0)
- feat(workbook): validate selector contracts (2c15f1d2)
- feat(engine): rewrite conditional format artifacts (fae790a0)
- feat(workbook): attach readback proof (87638938)

## Fixes

- perf(headless): skip empty structural insert planning (ed08fc47)
- perf(headless): append fresh formula family runs (e2569f4c)

## Internal runtime changes

- refactor(core): split structured reference resolver (f8e83d75)
- ci(runtime): fix WorkPaper external smoke contract (3da277b8)
- chore(release): runtime packages v0.56.0 (33e7c2b5)

## 0.56.0

- Release type: minor
- Previous libraries tag: libraries-v0.55.0
- Manual override: no

## Features

- feat(formula): preserve native structured references (9e417452)
- feat(workbook): avoid mutating model configs (5c42661b)
- feat(workbook): require own action metadata (be1bfc42)
- feat(workbook): freeze planned action handoffs (c3ac870f)
- feat(engine): execute excel autofilters (c4994591)
- feat(engine): support precision as displayed (308415f0)
- feat(workbook): verify immutable plan refs (55859f20)
- feat(engine): preserve excel conditional format artifacts (492ebfd0)
- feat(workbook): validate selector contracts (2c15f1d2)

## Fixes

- perf(headless): skip empty structural insert planning (ed08fc47)
- perf(headless): append fresh formula family runs (e2569f4c)

## Internal runtime changes

- refactor(core): split structured reference resolver (f8e83d75)
- ci(runtime): fix WorkPaper external smoke contract (3da277b8)

## 0.55.0

- Release type: minor
- Previous libraries tag: libraries-v0.54.0
- Manual override: no

## Features

- feat(workbook): expose selector contracts (dbbcc196)

## 0.54.0

- Release type: minor
- Previous libraries tag: libraries-v0.53.0
- Manual override: no

## Features

- feat(workbook): freeze model definitions (0428554d)

## Internal runtime changes

- test(headless): lock Desktop Excel sort oracle (991d1115)

## 0.53.0

- Release type: minor
- Previous libraries tag: libraries-v0.52.0
- Manual override: no

## Features

- feat(workbook): harden runtime proof boundary (1e6f3549)
- feat(core): add Desktop Excel-backed table sort (679b24f1)

## Fixes

- perf(headless): compact deferred sheet renames (f5f3556d)

## 0.52.0

- Release type: minor
- Previous libraries tag: libraries-v0.51.5
- Manual override: no

## Features

- feat(workbook): expose action input guards (390bd5e6)
- feat(workbook): include runtime requirements in model verification (ea3fdc66)

## Fixes

- fix(excel-import): preserve explicit numeric xlsx cells (99cb4cec)
- perf(headless): speed up direct aggregate row inserts (baaeac11)
- perf(headless): preserve simple column delete values (fb32f675)
- perf(headless): compose deferred column inserts (79773412)
- perf(headless): keep blank axis inserts sparse (038c7c72)
- fix(grid): reject stale visible scene state (9223d224)

## Internal runtime changes

- docs(agent): add framework tool chooser (c286e060)
- test(ui): tighten same-corpus grid proof (75d8aad7)

## 0.51.5

- Release type: patch
- Previous libraries tag: libraries-v0.51.4
- Manual override: no

## Fixes

- perf(headless): route structural inserts directly (22b2556d)

## Internal runtime changes

- ci(release): prevent runtime release self-trigger (60738a42)

## 0.51.4

- Release type: patch
- Previous libraries tag: libraries-v0.51.3
- Manual override: no

## Fixes

- fix(release): retry duplicate npm publish visibility (532f6a51)

## 0.51.3

- Release type: patch
- Previous libraries tag: libraries-v0.51.2
- Manual override: no

## Fixes

- fix(release): tolerate duplicate npm publish races (773dc878)

## 0.51.2

- Release type: patch
- Previous libraries tag: libraries-v0.51.1
- Manual override: no

## Fixes

- fix(corpus): preserve structural smoke and undo fidelity (2770b185)

## 0.51.1

- Release type: patch
- Previous libraries tag: libraries-v0.51.0
- Manual override: no

## Fixes

- perf(headless): skip table metadata on no-table edits (bf1bf013)

## 0.51.0

- Release type: minor
- Previous libraries tag: libraries-v0.50.1
- Manual override: no

## Features

- feat(excel-import): parse worksheet scalars in wasm (4c0feb4a)

## Fixes

- fix(core): avoid structural no-op history (1f28239b)
- fix(ci): use OIDC for runtime npm publishing (401ea159)

## 0.50.1

- Release type: patch
- Previous libraries tag: libraries-v0.50.0
- Manual override: no

## Fixes

- perf(headless): keep sliding aggregate edits inline (312d1055)

## 0.50.0

- Release type: minor
- Previous libraries tag: libraries-v0.48.0
- Manual override: no

## Features

- feat(workbook): add command bundle handoff (cbdb0bca)
- feat(workbook): add run receipts (3af2a5d2)
- feat(workbook): emit core runtime receipts (498ccdde)
- feat(workbook): add formula inspection (fed123c8)
- feat(workbook): add preview-only commands (4d3c16cd)
- feat(workbook): harden command handoff (10e32ecb)
- feat(workbook): add runtime capability preflight (85cdd5df)
- feat: reduce xlsx import memory footprint (659715c0)
- feat(excel-import): add wasm worksheet scan storage (62579199)

## Fixes

- fix(headless): preserve pruned cell history (5cfb183d)
- fix(core): settle cycle dependents after csv formula import (7c88c8dd)
- fix(headless): harden fuzz-found formula mutations (a7d6ac5d)
- fix(core): align fuzz-found formula parity (25ddd7c6)
- fix(formula): keep lookup criteria semantics (bc5821ed)
- fix(grid): keep local tile damage precise (b4ca7249)
- fix(core): isolate delete template rewrites for copied formulas (59bd97e8)
- fix(wasm): keep sort text ordering in parity (e44e0eea)
- fix(core): close fuzz-found recalc parity gaps (17024dd7)
- fix(import): restore merge-time correctness paths (3a9d2174)
- fix(ci): restore coverage gate after main merge (2191771a)
- fix(ci): restore coverage assertions and table fast path (9796fda0)
- fix(excel-import): preserve hidden totals row state (121a2eeb)
- fix(ci): keep merged source under size limit (540dc32e)
- fix(core): recalc overlapping aggregate text writes (ed4e2317)
- fix(core): satisfy aggregate fast path typing (c7b8bdb2)
- fix(core): validate collected aggregate fast path (f2c05b4d)
- fix(excel-import): share lazy materialization threshold (32f15f95)
- fix(core): satisfy direct criteria helper lint (1017f866)
- fix(core): stabilize direct aggregate structural replay (c6428461)
- fix(formula): match Excel partial move rewrites (74cc56ca)
- fix(workbook): harden fuzz regressions and gpu fill proof (638c502c)
- fix(grid): harden typegpu tile presentation (a0f0b4c8)
- fix(release): publish npm before release tags (d6a54d2f)
- fix(release): require dispatch for npm publish (9262aed2)
- fix(core): preserve structural tombstones in planned remaps (d39eab41)
- fix(core): enforce text length validation thresholds (9eb63932)
- fix(release): use oidc for runtime package publish (33d6102c)
- perf(headless): trust fresh tail append matrices (c4f73631)
- fix(release): clear npm auth config for oidc (ca8fb9e8)
- fix(core): preserve fuzz-discovered invariants (29aab55c)
- perf(headless): precompute fresh aggregate matrix results (9795f4e7)
- fix(release): publish from current main during runtime release (2e8f08a2)
- fix(release): preserve setup-node npm config for oidc (e1e19b14)
- perf(headless): reduce direct mutation metadata probes (ccbc0ea5)
- perf(headless): skip empty validation and dimension metadata work (09148e11)
- fix(ci): stage workpaper server release metadata (4324a609)

## Internal runtime changes

- test(fuzz): expand unified fuzz coverage (15cfe495)
- test(fuzz): cover runtime reducer edge cases (09a7ca79)
- test(formula): fuzz lookup and datetime families (eadbd6f4)
- test(import): fuzz xlsx escape fidelity (3760b68e)
- test(fuzz): expand correctness coverage (990b5652)
- test(import): align csv fuzz snapshot cells (e743c53f)
- test(ci): repair coverage contract failures (c88724de)
- test(core): cover inline scalar value helpers (43ab24b8)
- test(core): cover unsupported formula cache decisions (23c603e1)
- refactor(excel-import): split macro code name parsing (28262aa5)
- test(core): expand semantic coverage gate (64d3503b)
- test(core): cover aggregate fast path validation (b9573092)
- test(core): cover direct formula helpers (3dfdec29)
- test(core): cover aggregate post-recalc dependents (d1c87c29)
- refactor(core): extract oversized import helpers (31f8ba6b)
- refactor(xlsx): split worksheet stream metadata scanner (63e3b3e7)
- refactor(xlsx): split large simple arena helpers (b190b6df)
- chore(docs): refresh agent discovery after merge (3b59163f)
- chore(docs): refresh agent discovery after latest merge (fa557dc1)
- chore(docs): refresh agent discovery after merge (0ef96267)
- test(formula): cover serializer edge cases (89aff8f8)
- test(formula): cover workbook special calls (cb53c5cc)
- refactor(excel-import): split large simple cell release helper (c0059b0e)
- refactor(excel-import): split large simple worksheet materialization (350f4b62)
- refactor(core): split direct criteria aggregate helper (3b1608b9)
- test(core): cover structural formula rewrite guards (5d2c6e3e)
- test(headless): gate Desktop Excel oracle suite (3f42a1fa)
- test(headless): harden Desktop Excel oracle cleanup (3503a45c)
- chore(ci): split oversized merged sources (61946495)
- test(core): assert direct protection rejection (66c77f5b)
- test(core): cover literal fast paths (d10d0108)
- chore(release): runtime packages v0.49.0 (841dfe4e)

## 0.49.0

- Release type: minor
- Previous libraries tag: libraries-v0.48.0
- Manual override: no

## Features

- feat(workbook): add command bundle handoff (cbdb0bca)
- feat(workbook): add run receipts (3af2a5d2)
- feat(workbook): emit core runtime receipts (498ccdde)
- feat(workbook): add formula inspection (fed123c8)
- feat(workbook): add preview-only commands (4d3c16cd)
- feat(workbook): harden command handoff (10e32ecb)
- feat(workbook): add runtime capability preflight (85cdd5df)
- feat: reduce xlsx import memory footprint (659715c0)
- feat(excel-import): add wasm worksheet scan storage (62579199)

## Fixes

- fix(headless): preserve pruned cell history (5cfb183d)
- fix(core): settle cycle dependents after csv formula import (7c88c8dd)
- fix(headless): harden fuzz-found formula mutations (a7d6ac5d)
- fix(core): align fuzz-found formula parity (25ddd7c6)
- fix(formula): keep lookup criteria semantics (bc5821ed)
- fix(grid): keep local tile damage precise (b4ca7249)
- fix(core): isolate delete template rewrites for copied formulas (59bd97e8)
- fix(wasm): keep sort text ordering in parity (e44e0eea)
- fix(core): close fuzz-found recalc parity gaps (17024dd7)
- fix(import): restore merge-time correctness paths (3a9d2174)
- fix(ci): restore coverage gate after main merge (2191771a)
- fix(ci): restore coverage assertions and table fast path (9796fda0)
- fix(excel-import): preserve hidden totals row state (121a2eeb)
- fix(ci): keep merged source under size limit (540dc32e)
- fix(core): recalc overlapping aggregate text writes (ed4e2317)
- fix(core): satisfy aggregate fast path typing (c7b8bdb2)
- fix(core): validate collected aggregate fast path (f2c05b4d)
- fix(excel-import): share lazy materialization threshold (32f15f95)
- fix(core): satisfy direct criteria helper lint (1017f866)
- fix(core): stabilize direct aggregate structural replay (c6428461)
- fix(formula): match Excel partial move rewrites (74cc56ca)
- fix(workbook): harden fuzz regressions and gpu fill proof (638c502c)
- fix(grid): harden typegpu tile presentation (a0f0b4c8)
- fix(release): publish npm before release tags (d6a54d2f)
- fix(release): require dispatch for npm publish (9262aed2)
- fix(core): preserve structural tombstones in planned remaps (d39eab41)
- fix(core): enforce text length validation thresholds (9eb63932)
- fix(release): use oidc for runtime package publish (33d6102c)
- perf(headless): trust fresh tail append matrices (c4f73631)
- fix(release): clear npm auth config for oidc (ca8fb9e8)
- fix(core): preserve fuzz-discovered invariants (29aab55c)
- perf(headless): precompute fresh aggregate matrix results (9795f4e7)
- fix(release): publish from current main during runtime release (2e8f08a2)
- fix(release): preserve setup-node npm config for oidc (e1e19b14)
- perf(headless): reduce direct mutation metadata probes (ccbc0ea5)
- perf(headless): skip empty validation and dimension metadata work (09148e11)

## Internal runtime changes

- test(fuzz): expand unified fuzz coverage (15cfe495)
- test(fuzz): cover runtime reducer edge cases (09a7ca79)
- test(formula): fuzz lookup and datetime families (eadbd6f4)
- test(import): fuzz xlsx escape fidelity (3760b68e)
- test(fuzz): expand correctness coverage (990b5652)
- test(import): align csv fuzz snapshot cells (e743c53f)
- test(ci): repair coverage contract failures (c88724de)
- test(core): cover inline scalar value helpers (43ab24b8)
- test(core): cover unsupported formula cache decisions (23c603e1)
- refactor(excel-import): split macro code name parsing (28262aa5)
- test(core): expand semantic coverage gate (64d3503b)
- test(core): cover aggregate fast path validation (b9573092)
- test(core): cover direct formula helpers (3dfdec29)
- test(core): cover aggregate post-recalc dependents (d1c87c29)
- refactor(core): extract oversized import helpers (31f8ba6b)
- refactor(xlsx): split worksheet stream metadata scanner (63e3b3e7)
- refactor(xlsx): split large simple arena helpers (b190b6df)
- chore(docs): refresh agent discovery after merge (3b59163f)
- chore(docs): refresh agent discovery after latest merge (fa557dc1)
- chore(docs): refresh agent discovery after merge (0ef96267)
- test(formula): cover serializer edge cases (89aff8f8)
- test(formula): cover workbook special calls (cb53c5cc)
- refactor(excel-import): split large simple cell release helper (c0059b0e)
- refactor(excel-import): split large simple worksheet materialization (350f4b62)
- refactor(core): split direct criteria aggregate helper (3b1608b9)
- test(core): cover structural formula rewrite guards (5d2c6e3e)
- test(headless): gate Desktop Excel oracle suite (3f42a1fa)
- test(headless): harden Desktop Excel oracle cleanup (3503a45c)
- chore(ci): split oversized merged sources (61946495)
- test(core): assert direct protection rejection (66c77f5b)
- test(core): cover literal fast paths (d10d0108)

## 0.41.0

- Release type: minor
- Previous libraries tag: libraries-v0.40.42
- Manual override: no

## Features

- feat(workbook): publish agent-first workbook API (e2fc16ed)
- feat(workbook): add structured agent planning (8f14305b)
- feat(workbook): expose formula inputs in plans (b04ab9ab)
- feat(workbook): expose resolved refs in plans (51fcfc41)
- feat(workbook): describe agent action plans (3f3cafe6)
- feat(workbook): describe planning failures (fb4eda3b)
- feat(workbook): verify action plans (4313167a)
- feat(workbook): verify entire models (41be866d)
- feat(workbook): expose simple find helpers (dfb05d8a)
- feat(workbook): expose simple check helpers (8f1190a7)
- feat(workbook): describe model manifests (7eaebafe)
- feat(workbook): support custom checks (6505a3c5)
- feat(workbook): track custom check refs (4fa18c6c)
- feat(workbook): compile number format actions (ab2e00ef)
- feat(workbook): allow guarded low-level ops (acd4492b)
- feat(workbook): declare raw formula inputs (0eb46af3)
- feat(workbook): add parameterized action input (27c2fdce)
- feat(workbook): add readback check expectations (ce1613fa)
- feat(workbook): add generic run readback receipts (8fb77566)
- feat(workbook): add generic check verifier hook (0d0e225e)
- feat(core): execute workbook action plans (b18019df)
- feat(excel-import): import Excel data tables (e658ca15)
- feat(excel-import): import one-variable data tables (4bd97433)
- feat(workbook): execute row selectors (733ba978)
- feat(workbook): describe runtime requirements (7d0d525a)
- feat(excel-import): materialize cached external ranges (2bf89130)
- feat(excel-import): back external caches with hidden ranges (b3f8d7f8)
- feat(workbook): describe run results (02b4aacb)
- feat(workbook): harden agent workbook intent (486e6851)
- feat(workbook): scope agent model phases (14351c2f)
- feat(workbook): describe action inputs (42951a13)
- feat(workbook): stabilize run error codes (e96bf90c)

## Fixes

- perf(excel-import): finalize simple sheets earlier (c8c573cd)
- fix(excel-export): preserve dynamic array spill caches (7e7ef8d6)
- perf(excel-import): reuse streamed zip read buffers (e553d848)
- perf(headless): bind fresh direct scalar formulas directly (45f1d9e7)
- perf(excel-import): lazy materialize medium sheets (49e98655)
- fix(headless): match Excel structural move semantics (53a6552d)
- fix(excel-import): preserve reusable zip byte reads (92b675aa)
- fix(workbook): distinguish row selector refs (b0ec8d69)
- fix(workbook): finish package rename after rebase (1819bd40)
- fix(headless): match Excel table structural refs (51a49667)
- perf(excel-import): finalize unstyled sheets before styles (5b50b1b7)
- fix(ci): build renamed workbook package (36e0bf6d)
- fix(headless): rewrite deleted table refs like Excel (e05bf105)
- perf(core): restore mixed runtime family runs (df933ce9)
- fix(corpus): slim memory gate worker (c5e168f3)
- fix(headless): match Excel table header renames (e904bc7b)
- perf(excel-import): lower large xlsx verifier rss (02c71a22)
- perf(excel-import): unblock simple sheets with artifact peers (0d13bd2e)
- fix(headless): rewrite deleted table names like Excel (21ebf354)
- fix(headless): rewrite renamed table names like Excel (5f4794de)
- perf(excel-import): coalesce streamed axis metadata (895b106d)
- fix(headless): canonicalize table headers like Excel (82b9375d)
- perf(excel-import): drain no-op styled sheet arenas (e535bb6c)
- perf(core): route formula literal batches earlier (14576712)
- perf(excel-import): intern streamed metadata in place (7b69df5b)
- fix(headless): preserve spaced table references like Excel (882c118e)
- perf(core): index criteria aggregate range edges (6cbb97ae)
- fix(excel): escape structured table headers (3251a9b8)
- fix(workbook): clear type-aware lint blockers (cd4e52da)
- fix(headless): scalarize standalone INDEX references like Excel (a03ba0de)
- perf(excel-import): reuse stored zip entry buffers (14956861)
- perf(excel-import): combine streamed style artifact parsing (fe10b931)
- fix(headless): scalarize standalone OFFSET references like Excel (a0260c6a)
- perf(excel-import): shrink sparse dimension preallocation (15aeafc9)
- fix(headless): support CHOOSE virtual table spills (a0b6db1d)
- perf(excel-import): spool large public source bytes (9d941257)
- fix(excel-import): encode spill references for desktop excel (cf2b6148)
- fix(excel-import): preserve single implicit intersections (35cb7208)
- perf(excel-import): lazy load sheetjs fallback (e743b6c9)
- fix(grid): refresh visible tile interest and style invalidations (a9befb54)
- fix(excel-oracle): read desktop excel spill errors (55e18e06)
- fix(core): expose default package export (8d15060a)
- perf(excel-import): drop import buffers before scanner load (6637f3db)
- fix(package): expose default runtime exports (fb252a93)
- perf(excel-import): lazy load byte-source scanner modules (bbde0a9d)
- fix(formula): match Excel implicit intersection errors (d61d3180)
- perf(excel-import): cap dimension arena preallocation (21c0f9be)
- perf(excel-import): intern resolved metadata strings (f253215a)
- perf(excel-import): detach lazy sheet cells from arena (f9c05822)
- perf(headless): tune fresh aggregate native threshold (5416ff5d)
- fix(formula): honor aggregate options (d1e0da30)
- fix(workbook): gate workflow applies on mutation proof (313bedbf)
- perf(excel-import): trim formula scan storage (78a2b168)
- fix(workbook): fail unverified checks (11e4816c)
- perf(headless): fast path compact numeric changes (84cf58fa)
- perf(excel-import): transfer detached lazy pools (c0c3bda4)
- fix(spill): preserve one-cell dynamic arrays (aece2475)
- perf(core): shortcut single aggregate owner lookup (45000053)
- fix(spill): match Excel blocked child edits (a4776ff8)
- perf(excel-import): reduce formatted cell materialization memory (42c13802)
- fix(spill): rematerialize structural spill edits (3a800128)
- perf(core): keep small aggregate appends numeric (b84ec201)
- fix(spill): rebind moved spill refs (d6ffbb2f)
- perf(excel-import): finalize spooled sheets earlier (e1b782f3)
- fix(names): preserve deleted refs as ref errors (1440d7a0)
- perf(excel-import): reduce inline string import RSS (cb7cbe12)
- fix(data-table): retarget native metadata (0b16a0a0)
- perf(core): bulk-bind fresh direct scalar runs (15cd6c27)
- perf(excel-import): retain shared strings for lazy sheets (bddb0593)
- fix(array-formulas): preserve native metadata (b9c40823)

## Internal runtime changes

- chore(release): runtime packages v0.40.43 (4c021baa)
- docs(agent): align WorkPaper MCP discovery (f9934bbd)
- refactor(core): split live effect runner (3fe3b47c)
- docs(agent): publish workpaper mcp metadata (cb44b298)
- docs(growth): surface n8n formula readback path (62351045)
- chore(bench): require 200-sample competitive evidence (3326c11e)
- refactor(formula): split workbook reference calls (06a97af8)

## 0.40.43

- Release type: patch
- Previous libraries tag: libraries-v0.40.42
- Manual override: no

## Fixes

- perf(excel-import): finalize simple sheets earlier (c8c573cd)

## 0.40.42

- Release type: patch
- Previous libraries tag: libraries-v0.40.41
- Manual override: no

## Fixes

- perf(core): bulk bind fresh aggregate rows (6444c861)
- perf(excel-import): compact shared string import storage (f5a34983)
- fix(formula): align empty string criteria with Excel (f48b401c)

## Internal runtime changes

- ci(release): block npm publish on red repo ci (080cfb1c)
- chore(workpaper): remove root excel oracle script (3e66b459)
- test(excel): add package-owned macos oracle (ab2080bd)
- ci(release): gate release mutation on green ci (cb9f8c3b)
- test(excel): report desktop oracle comparisons (f4aaa093)
- ci(release): verify release commit before publish (c85cc159)
- test(benchmarks): raise expanded comparison samples (5f345bd6)
- ci(release): skip stale publish runs cleanly (a7417f8c)

## 0.40.41

- Release type: patch
- Previous libraries tag: libraries-v0.40.40
- Manual override: no

## Fixes

- perf(excel-import): retain only rich shared string refs (8c14a27a)

## Internal runtime changes

- ci(n8n): prefer oidc for npm publish (a16fffe0)

## 0.40.40

- Release type: patch
- Previous libraries tag: libraries-v0.40.35
- Manual override: no

## Fixes

- perf(core): index sliding aggregate dependent collection (e726b7b8)
- perf(core): tighten direct formula delta batches (a55ac416)
- perf(excel-import): stream control workbook artifacts (e8ee7810)
- perf(core): preserve mixed direct delta batches (fd03a216)
- perf(excel-import): trim fast xlsx import memory (7d0aa436)
- perf(core): aggregate indexed mixed criteria (dd3290ae)
- perf(excel-import): add headless external stress path (6a75ce87)
- perf(excel-import): dedupe streamed shared strings (4180d45e)
- perf(excel-import): report metadata in headless inspect (ed6f6b03)
- perf(core): skip redundant fresh scalar chunk planning (fd617e24)
- perf(excel-import): build previews after cell release (e96200c1)

## Internal runtime changes

- refactor(excel-import): split style artifact candidates (069bdf62)
- ci(release): require full ci before npm publish (f191097b)
- chore(release): runtime packages v0.40.36 (059f37b7)
- chore(release): runtime packages v0.40.37 (809a0b3b)
- chore(release): runtime packages v0.40.38 (f2d376da)
- chore(release): runtime packages v0.40.39 (7622638e)
- refactor(excel-import): split large simple import types (bb28d403)
- ci(n8n): publish node from trusted workflow (3d47a4a5)
- ci(n8n): gate runtime publish on manual dispatch (ba6b956e)

## 0.40.39

- Release type: patch
- Previous libraries tag: libraries-v0.40.35
- Manual override: no

## Fixes

- perf(core): index sliding aggregate dependent collection (e726b7b8)
- perf(core): tighten direct formula delta batches (a55ac416)
- perf(excel-import): stream control workbook artifacts (e8ee7810)
- perf(core): preserve mixed direct delta batches (fd03a216)
- perf(excel-import): trim fast xlsx import memory (7d0aa436)
- perf(core): aggregate indexed mixed criteria (dd3290ae)
- perf(excel-import): add headless external stress path (6a75ce87)

## Internal runtime changes

- refactor(excel-import): split style artifact candidates (069bdf62)
- ci(release): require full ci before npm publish (f191097b)
- chore(release): runtime packages v0.40.36 (059f37b7)
- chore(release): runtime packages v0.40.37 (809a0b3b)
- chore(release): runtime packages v0.40.38 (f2d376da)

## 0.40.38

- Release type: patch
- Previous libraries tag: libraries-v0.40.35
- Manual override: no

## Fixes

- perf(core): index sliding aggregate dependent collection (e726b7b8)
- perf(core): tighten direct formula delta batches (a55ac416)
- perf(excel-import): stream control workbook artifacts (e8ee7810)
- perf(core): preserve mixed direct delta batches (fd03a216)
- perf(excel-import): trim fast xlsx import memory (7d0aa436)

## Internal runtime changes

- refactor(excel-import): split style artifact candidates (069bdf62)
- ci(release): require full ci before npm publish (f191097b)
- chore(release): runtime packages v0.40.36 (059f37b7)
- chore(release): runtime packages v0.40.37 (809a0b3b)

## 0.40.37

- Release type: patch
- Previous libraries tag: libraries-v0.40.35
- Manual override: no

## Fixes

- perf(core): index sliding aggregate dependent collection (e726b7b8)
- perf(core): tighten direct formula delta batches (a55ac416)
- perf(excel-import): stream control workbook artifacts (e8ee7810)

## Internal runtime changes

- refactor(excel-import): split style artifact candidates (069bdf62)
- ci(release): require full ci before npm publish (f191097b)
- chore(release): runtime packages v0.40.36 (059f37b7)

## 0.40.36

- Release type: patch
- Previous libraries tag: libraries-v0.40.35
- Manual override: no

## Fixes

- perf(core): index sliding aggregate dependent collection (e726b7b8)

## Internal runtime changes

- refactor(excel-import): split style artifact candidates (069bdf62)
- ci(release): require full ci before npm publish (f191097b)

## 0.40.35

- Release type: patch
- Previous libraries tag: libraries-v0.40.34
- Manual override: no

## Fixes

- perf(excel-import): avoid byte-source fallback materialization (c1eb1103)

## 0.40.34

- Release type: patch
- Previous libraries tag: libraries-v0.40.33
- Manual override: no

## Fixes

- perf(core): bind fresh direct formula runs in one pass (9d7c76e5)

## 0.40.33

- Release type: patch
- Previous libraries tag: libraries-v0.40.32
- Manual override: no

## Fixes

- perf(excel-import): release sheet import storage sooner (2dcb5ab3)

## 0.40.32

- Release type: patch
- Previous libraries tag: libraries-v0.40.31
- Manual override: no

## Fixes

- perf(excel-import): avoid verifier source rereads (f64db341)

## 0.40.31

- Release type: patch
- Previous libraries tag: libraries-v0.40.30
- Manual override: no

## Fixes

- perf(headless): replay numeric undo batches from typed history (553b29e3)
- perf(excel-import): skip redundant calcchain inspection (6e6e03a4)

## Internal runtime changes

- refactor(excel-import): split large simple helpers (397fd290)
- test(headless): expect packed numeric undo history (38acf175)

## 0.40.30

- Release type: patch
- Previous libraries tag: libraries-v0.40.29
- Manual override: no

## Fixes

- perf(xlsx): store rare arena integers sparsely (1a52eb87)

## 0.40.29

- Release type: patch
- Previous libraries tag: libraries-v0.40.28
- Manual override: no

## Fixes

- perf(xlsx): pack small integers in import arena (4fa2d029)

## 0.40.28

- Release type: patch
- Previous libraries tag: libraries-v0.40.27
- Manual override: no

## Fixes

- perf(xlsx): compact headless verifier imports (9347c803)

## 0.40.27

- Release type: patch
- Previous libraries tag: libraries-v0.40.26
- Manual override: no

## Fixes

- perf(headless): lazily summarize aggregate pages (1c7cdea2)
- perf(xlsx): stream package artifact imports (74033d79)

## 0.40.26

- Release type: patch
- Previous libraries tag: libraries-v0.40.25
- Manual override: no

## Fixes

- perf(xlsx): stream auto-filter metadata (bfdc18e2)
- perf(headless): avoid singleton hydrated formula bindings (18b998f3)

## Internal runtime changes

- refactor(excel-import): split worksheet stream cell readers (9a2588b9)

## 0.40.25

- Release type: patch
- Previous libraries tag: libraries-v0.40.24
- Manual override: no

## Fixes

- perf(xlsx): stream column metadata (dc088b51)

## 0.40.24

- Release type: patch
- Previous libraries tag: libraries-v0.40.23
- Manual override: no

## Fixes

- perf(xlsx): stream conditional format metadata (2a33eb95)

## 0.40.23

- Release type: patch
- Previous libraries tag: libraries-v0.40.22
- Manual override: no

## Fixes

- perf(xlsx): pool cached formula records (dbf6365c)
- perf(xlsx): release plain shared string tables (820fde93)
- fix(xlsx): restore large workbook export path (b5232805)
- perf(xlsx): resolve shared strings before materialization (04e4e70a)
- perf(headless): skip duplicate batch literal classification (73d49f21)
- perf(xlsx): stream highly compressed zip entries (15caba30)
- perf(xlsx): count headless data validations (514431f1)

## 0.40.22

- Release type: patch
- Previous libraries tag: libraries-v0.40.21
- Manual override: no

## Fixes

- perf(core): fast-bind hydrated aggregate formulas (f6ead970)
- perf(xlsx): stream small PowerPivot packages (baf6dc71)

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
- perf(workbook): harden headless engine benchmark gates (b0e399e9)
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
- perf(workbook): harden headless engine benchmark gates (b0e399e9)
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
- perf(workbook): harden headless engine benchmark gates (b0e399e9)
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
- perf(workbook): harden headless engine benchmark gates (b0e399e9)
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
- perf(workbook): harden headless engine benchmark gates (b0e399e9)
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
- perf(workbook): harden headless engine benchmark gates (b0e399e9)
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
