---
title: Verify npm provenance for @bilig/headless
published: true
description: How to verify the published @bilig/headless package before adopting it in a Node service or agent tool.
tags: npm, provenance, typescript, security, node
canonical_url: https://proompteng.github.io/bilig/npm-provenance-package-trust.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Verify npm Provenance For `@bilig/headless`

Production adoption starts before the first import. For a service runtime or
agent tool, the package needs to be traceable to source, release CI, and a
specific GitHub commit.

`@bilig/headless@0.18.3` is published with npm registry signatures and SLSA
provenance attestations. npm reports:

```sh
npm view @bilig/headless@0.18.3 version dist.attestations dist.signatures --json
```

The important signal is that `dist.attestations.provenance.predicateType` is
`https://slsa.dev/provenance/v1` and that `dist.signatures` is non-empty.

## Verify After Install

From a clean project:

```sh
mkdir bilig-package-trust
cd bilig-package-trust
npm init -y
npm install @bilig/headless@0.18.3
npm audit signatures
```

Expected result for the current dependency tree:

```text
audited 31 packages in 0s

31 packages have verified registry signatures

10 packages have verified attestations
```

Use this as a package-integrity check, not as an application-security claim.
You still need workflow fixtures, rollback, and formula compatibility gates for
your own WorkPaper-backed service.

## Release Path

Runtime packages are released by `.github/workflows/headless-package.yml`.
The workflow:

- verifies the runtime package chain;
- checks publishable package metadata with `pnpm publish:runtime:check`;
- requires Forgejo and GitHub `main` to agree before publishing;
- uses `id-token: write` for GitHub Actions OIDC;
- publishes through `scripts/publish-runtime-package-set.ts` with
  `npm publish ... --provenance`.

npm documents trusted publishing as an OIDC flow that avoids long-lived npm
tokens and can automatically generate provenance for public packages published
from public repositories:

- <https://docs.npmjs.com/trusted-publishers/>
- <https://docs.npmjs.com/viewing-package-provenance/>

OpenSSF Scorecard is another useful consumer-side signal for evaluating
dependency risk:

- <https://scorecard.dev/>

This repository runs the official OpenSSF Scorecard action on every `main`
update and on a weekly schedule. Results are published to the public Scorecard
API, exposed through the README badge, and uploaded as SARIF to GitHub code
scanning so dependency evaluators can inspect repository posture separately
from npm package provenance.

The GitHub trust surface also includes CodeQL analysis for the
JavaScript/TypeScript codebase and Dependabot version updates for npm, GitHub
Actions, and the root Dockerfile. Those checks do not replace package
provenance, but they make vulnerability discovery and dependency drift visible
before a production adopter has to ask for it.

## What This Does Not Prove

Package provenance does not prove that a workbook workflow is correct, complete,
or safe for every production domain.

Before adopting `@bilig/headless` for customer-critical work, also run:

- the [90-second npm eval](try-bilig-headless-in-node.md);
- the [quote approval WorkPaper API proof](quote-approval-workpaper-api.md);
- the [production adoption checklist](production-adoption-checklist-headless-workpaper.md);
- the [compatibility limits](where-bilig-is-not-excel-compatible-yet.md).

The package-trust question is: "Did this package come from the expected source
and release path?" The production-readiness question is: "Does this exact
workflow have fixtures, rollback, and compatibility evidence?"
