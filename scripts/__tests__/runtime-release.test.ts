import { describe, expect, it } from 'vitest'

import { bumpVersion, isRuntimeAffectingPath, parseConventionalCommit, releaseTypeForConventionalCommit } from '../runtime-release.ts'

describe('runtime release helpers', () => {
  it('parses standard conventional commits', () => {
    const parsed = parseConventionalCommit({
      subject: 'feat(core): add runtime planner',
      body: '',
    })

    expect(parsed).toEqual({
      type: 'feat',
      scope: 'core',
      description: 'add runtime planner',
      breaking: false,
    })
  })

  it('detects breaking changes from bang markers and footer markers', () => {
    const bang = parseConventionalCommit({
      subject: 'feat(core)!: replace runtime release flow',
      body: '',
    })
    const footer = parseConventionalCommit({
      subject: 'fix(core): preserve publish ordering',
      body: 'BREAKING CHANGE: old release path removed',
    })

    expect(bang?.breaking).toBe(true)
    expect(footer?.breaking).toBe(true)
  })

  it('maps conventional commit kinds to semantic release types', () => {
    expect(
      releaseTypeForConventionalCommit({
        type: 'fix',
        scope: null,
        description: 'repair package metadata',
        breaking: false,
      }),
    ).toBe('patch')

    expect(
      releaseTypeForConventionalCommit({
        type: 'feat',
        scope: null,
        description: 'add runtime release planner',
        breaking: false,
      }),
    ).toBe('minor')

    expect(
      releaseTypeForConventionalCommit({
        type: 'refactor',
        scope: null,
        description: 'shuffle internal helpers',
        breaking: false,
      }),
    ).toBe('none')

    expect(
      releaseTypeForConventionalCommit({
        type: 'chore',
        scope: null,
        description: 'drop old release flow',
        breaking: true,
      }),
    ).toBe('major')
  })

  it('bumps semantic versions correctly', () => {
    expect(bumpVersion('0.1.2', 'patch')).toBe('0.1.3')
    expect(bumpVersion('0.1.2', 'minor')).toBe('0.2.0')
    expect(bumpVersion('0.1.2', 'major')).toBe('1.0.0')
  })

  it('matches runtime-affecting publish paths', () => {
    expect(isRuntimeAffectingPath('packages/core/src/index.ts')).toBe(true)
    expect(isRuntimeAffectingPath('scripts/publish-runtime-package-set.ts')).toBe(true)
    expect(isRuntimeAffectingPath('apps/web/src/App.tsx')).toBe(false)
  })
})
