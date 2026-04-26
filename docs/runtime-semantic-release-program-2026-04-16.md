# Runtime Semantic Release Program

Status: implemented

Date: 2026-04-16

## Objective

Replace the current aligned library package release versioning logic with a production release system that:

- derives semver from Conventional Commits instead of always patch-bumping
- publishes the aligned library package set to npm from GitHub Actions
- creates authoritative release tags and releases for each published version
- generates stable changelog content from the same release metadata
- respects the repo rule that Forgejo `origin` remains the source of truth

This program applies to the aligned library package set:

- `@bilig/protocol`
- `@bilig/workbook-domain`
- `@bilig/wasm-kernel`
- `@bilig/formula`
- `@bilig/core`
- `@bilig/headless`

Implemented in:

- [/Users/gregkonush/github.com/bilig/scripts/plan-runtime-release.ts](/Users/gregkonush/github.com/bilig/scripts/plan-runtime-release.ts)
- [/Users/gregkonush/github.com/bilig/scripts/check-runtime-commit-policy.ts](/Users/gregkonush/github.com/bilig/scripts/check-runtime-commit-policy.ts)
- [/Users/gregkonush/github.com/bilig/scripts/sync-runtime-release-metadata.ts](/Users/gregkonush/github.com/bilig/scripts/sync-runtime-release-metadata.ts)
- [/Users/gregkonush/github.com/bilig/.github/workflows/headless-package.yml](/Users/gregkonush/github.com/bilig/.github/workflows/headless-package.yml)
- [/Users/gregkonush/github.com/bilig/.github/workflows/ci.yml](/Users/gregkonush/github.com/bilig/.github/workflows/ci.yml)

## Current State

The current runtime release path is centered on:

- [/Users/gregkonush/github.com/bilig/.github/workflows/headless-package.yml](/Users/gregkonush/github.com/bilig/.github/workflows/headless-package.yml)
- [/Users/gregkonush/github.com/bilig/scripts/runtime-package-set.ts](/Users/gregkonush/github.com/bilig/scripts/runtime-package-set.ts)
- [/Users/gregkonush/github.com/bilig/scripts/next-runtime-release-version.ts](/Users/gregkonush/github.com/bilig/scripts/next-runtime-release-version.ts)
- [/Users/gregkonush/github.com/bilig/scripts/publish-runtime-package-set.ts](/Users/gregkonush/github.com/bilig/scripts/publish-runtime-package-set.ts)

Today, the next library package version is chosen as follows:

1. If the package manifest version is ahead of npm, use the manifest version.
2. Otherwise, increment the published version by one patch.

This means:

- commit history does not affect version choice
- `feat:` does not trigger a minor bump
- breaking changes do not trigger a major bump
- changelog notes are not generated from release commits
- current package changelog claims are misleading

The current workflow also creates a GitHub release and tag, but the repo policy says Forgejo `origin` is the source of truth. A release tag that only exists on GitHub is not acceptable as the canonical release boundary.

## Problems To Fix

1. Version selection is wrong for public package consumers.
   `fix`, `feat`, and breaking changes all collapse into a patch bump unless a human manually edits package versions.

2. Git tags are not modeled as the authoritative release boundary.
   The current flow derives the next version from manifests plus npm registry state instead of from release history in git.

3. GitHub-native release state is ahead of repo truth.
   The current workflow is willing to create release metadata on GitHub without first proving that Forgejo and GitHub have the same release boundary.

4. Changelog ownership is unclear.
   `packages/headless/CHANGELOG.md` claims release-please ownership, but the repo does not actually run release-please.

5. The aligned library package set is special.
   These six packages are version-locked as one train. The release system must treat them as one release unit, not as six unrelated npm packages.

## Decision

Adopt a repo-owned semantic runtime release pipeline, executed by GitHub Actions, with Conventional Commits as the version source of truth and git tags as the release boundary.

The important design choice is this:

- use GitHub Actions as the executor
- keep release logic in repo-owned scripts
- do not hand authoritative control of release state to a GitHub-native release PR bot

This is the right tradeoff for `bilig` because:

- the source of truth is Forgejo `origin`, not GitHub
- the aligned library package set is custom and aligned
- the publish path already uses custom staging and pack/publish logic
- we need deterministic control over tags, changelog content, and sync behavior between remotes

## Chosen Model

### Release Boundary

The release boundary is the most recent reachable annotated tag matching:

- `libraries-vX.Y.Z`

This tag becomes the only input for "what was the last release?".

The next release version is derived from commits on `main` after the last matching tag and before the current release commit.

### Version Derivation

Version bump rules:

- `fix:` -> patch
- `feat:` -> minor
- any `!` or `BREAKING CHANGE:` footer -> major
- `docs`, `chore`, `test`, `build`, `ci`, `style`, and `refactor` without breaking markers -> no release by default

If multiple commits exist since the last library release tag, the strongest bump wins:

- major > minor > patch

If no runtime-affecting commits exist since the last tag, no runtime release is cut.

### Commit Scope Policy

Because the aligned library package set is one train, the release planner must consider both commit metadata and touched paths.

The planner should treat a commit as runtime-affecting when either of the following is true:

- the commit touches one of the runtime package directories
- the commit touches release/publish/build files that change the published runtime artifacts or release semantics

Initial runtime-affecting path set:

- `packages/protocol/**`
- `packages/workbook-domain/**`
- `packages/wasm-kernel/**`
- `packages/formula/**`
- `packages/core/**`
- `packages/headless/**`
- `scripts/runtime-package-set.ts`
- `scripts/publish-runtime-package-set.ts`
- `scripts/check-package-publish.ts`
- `scripts/gen-formula-dominance-snapshot.ts`
- `scripts/gen-workpaper-hyperformula-audit.ts`
- `scripts/gen-workpaper-benchmark-baseline.ts`
- `scripts/gen-workpaper-vs-hyperformula-benchmark.ts`
- `scripts/gen-workpaper-vs-hyperformula-benchmark.ts`
- `scripts/workpaper-external-smoke.ts`
- `.github/workflows/headless-package.yml`

This keeps release choice tied to real published surface, not just commit wording.

## Recommended Tooling

### Use

- `commitlint` for Conventional Commit enforcement
- `conventional-recommended-bump` or `@semantic-release/commit-analyzer` for bump calculation
- `conventional-changelog` or `@semantic-release/release-notes-generator` for release notes
- existing repo publish scripts for actual npm pack/publish

### Do Not Adopt As The Source Of Truth

- `release-please-action` as the primary runtime release owner
- `semantic-release` as a black-box monorepo publisher
- `changesets` for the aligned library package train

Why:

- `release-please` is good at commit-driven release PRs, but its natural model is GitHub-centered release state. That is a mismatch for this repo.
- `semantic-release` is good at fully automatic releases, but it is a poor fit for a custom aligned package set plus Forgejo source-of-truth constraints.
- `changesets` is excellent for monorepos, but it intentionally moves version intent into checked-in changeset files instead of deriving it from commit semantics.

## Release Architecture

### 1. Commit Policy

Add a required CI gate that enforces Conventional Commit release inputs.

Requirements:

- commits merged to `main` that are intended to affect the aligned library package set must use Conventional Commits
- if the team uses squash merges, the PR title must also satisfy the same format
- breaking changes must use either `!` or a `BREAKING CHANGE:` footer

Implementation:

- add commit/PR-title lint in CI
- fail fast before the publish workflow is ever asked to infer a version

### 2. Semantic Release Planner

Add a new repo-owned script:

- `scripts/plan-runtime-release.ts`

Responsibilities:

- locate the last `libraries-v*` tag
- collect commits since that tag
- determine whether each commit is runtime-affecting
- compute the strongest required bump
- produce `targetVersion`
- produce release notes markdown
- produce structured JSON outputs for downstream jobs

Example output shape:

```json
{
  "releaseNeeded": true,
  "lastTag": "libraries-v0.1.2",
  "targetVersion": "0.2.0",
  "bump": "minor",
  "commits": [
    {
      "sha": "abc123",
      "subject": "feat(core): add ...",
      "releaseType": "minor",
      "runtimeAffecting": true
    }
  ],
  "notesMarkdown": "## 0.2.0\n..."
}
```

### 3. Publish Workflow Refactor

Refactor [/Users/gregkonush/github.com/bilig/.github/workflows/headless-package.yml](/Users/gregkonush/github.com/bilig/.github/workflows/headless-package.yml) into three clear phases:

1. `verify`
   keep the current quality gates and pack validation

2. `plan-release`
   run `scripts/plan-runtime-release.ts`
   expose:
   - `release_needed`
   - `target_version`
   - `tag_name`
   - `notes_path`

3. `publish-release`
   only run if `release_needed == true`
   perform tag creation, npm publish, and release creation

### 4. Tag Authority

The library release tag must exist on Forgejo `origin` and GitHub, with the same object and target SHA.

Required behavior:

- before creating the tag, verify that `origin/main` and `github/main` point at the same commit
- create the annotated release tag from that exact commit
- push the tag to Forgejo first or in the same release transaction as GitHub
- fail the workflow if the remotes are not aligned

The workflow must not create a release tag that exists only on GitHub.

### 5. GitHub Release

Create one GitHub Release per aligned library package version:

- tag: `libraries-vX.Y.Z`
- title: `Libraries vX.Y.Z`
- notes: planner-generated release notes

The GitHub Release is the external consumer changelog artifact for the published version.

### 6. npm Publish

Keep using [/Users/gregkonush/github.com/bilig/scripts/publish-runtime-package-set.ts](/Users/gregkonush/github.com/bilig/scripts/publish-runtime-package-set.ts).

Change only the version source:

- stop feeding it the old patch-auto-increment result
- feed it `targetVersion` from the semantic release planner

Keep:

- aligned package set behavior
- staged tarball rewriting
- internal dependency rewriting
- dist-tag behavior
- dry-run/manual override inputs

### 7. Changelog Strategy

There are two changelog layers.

External changelog:

- canonical source: GitHub Release notes for `libraries-vX.Y.Z`

Repo-visible changelog:

- generated by a follow-up release sync commit or PR against Forgejo

The repo-visible changelog should not block npm publish. It is metadata sync, not release authority.

Initial repo-visible change set:

- update `packages/headless/CHANGELOG.md`
- update aligned library package manifest versions to the released version
- optionally add a root `docs/runtime-package-releases.md`

This sync should be created as a Forgejo PR, not as an unreviewed GitHub-only branch mutation.

## Rollout Plan

### Phase 0: Cleanup Current Misstatements

- remove the false release-please claim from `packages/headless/CHANGELOG.md`
- rename the workflow internally from "headless package" semantics to "aligned library package set" semantics where appropriate

### Phase 1: Enforce Conventional Commits

- add CI validation for commit messages and PR titles
- document required release commit conventions in repo docs
- fail PRs that would reach `main` with nonconforming release semantics

Acceptance criteria:

- every runtime-affecting merge to `main` is machine-readable for release planning

### Phase 2: Introduce Semantic Release Planner In Dry-Run Mode

- add `scripts/plan-runtime-release.ts`
- add `pnpm publish:runtime:plan`
- run it in CI without publishing
- compare planner output against expected bumps on recent history

Acceptance criteria:

- planner is deterministic
- planner returns "no release" when only non-runtime or non-release commits exist
- planner returns correct major/minor/patch on known history samples

### Phase 3: Switch Publish Workflow To Planner Output

- wire the GitHub Actions publish workflow to planner output
- preserve `workflow_dispatch target_version` as emergency override
- stop using `next-runtime-release-version.ts` for automatic version choice

Acceptance criteria:

- normal `push` releases use semantic commit analysis
- manual override still works for emergency or backfill releases

### Phase 4: Fix Tag Authority

- add explicit remote alignment check between Forgejo and GitHub
- create and push library release tags in a way that guarantees both remotes converge
- fail release if remotes differ

Acceptance criteria:

- every published npm version has one matching git tag on both remotes
- no GitHub-only library release tags exist

### Phase 5: Changelog And Repo Metadata Sync

- add a post-publish Forgejo sync PR flow for:
  - aligned manifest versions
  - `packages/headless/CHANGELOG.md`
  - optional aggregated runtime changelog doc

Acceptance criteria:

- the published version is visible both in external release artifacts and in repo metadata
- changelog sync does not own release numbering

## Required Workflow Changes

### Replace

- semantic version choice in `scripts/next-runtime-release-version.ts`

### Keep

- `scripts/runtime-package-set.ts`
- `scripts/publish-runtime-package-set.ts`
- `scripts/check-package-publish.ts`
- current build, smoke, and pack validation gates

### Add

- `scripts/plan-runtime-release.ts`
- `pnpm publish:runtime:plan`
- commit/PR-title lint workflow or job
- release sync workflow or script for Forgejo metadata PRs

## Acceptance Criteria

The program is complete when all of the following are true:

- runtime releases derive major/minor/patch from Conventional Commits
- the aligned library package set still publishes as one train
- the release boundary is the last `libraries-vX.Y.Z` git tag
- GitHub Actions only publishes when Forgejo and GitHub point at the same source commit
- npm publish, git tag, and GitHub Release all refer to the same version and commit
- package and docs changelog claims no longer mention tools the repo does not actually use
- manual emergency override remains available

## Rollback

If the semantic planner proves unstable:

- disable automatic publish on `push`
- keep `workflow_dispatch target_version`
- continue using the existing publish script
- cut releases manually while planner issues are fixed

Rollback must never silently fall back to patch-bumping without an explicit operator decision.

## Recommendation Summary

This repo should not adopt a generic GitHub-native release bot as the release authority.

The right production design is:

- Conventional Commit enforcement
- repo-owned semantic release planning
- existing aligned package publish script
- authoritative git tags across both remotes
- GitHub Release notes as external changelog
- Forgejo sync PRs for repo-visible version/changelog metadata

That gives the repo correct semantic versioning without giving up control over an unusual but intentional release model.
