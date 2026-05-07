# Security Policy

Security reports are handled separately from normal bug reports so sensitive
details do not become public before a fix is available.

## Supported Versions

Security fixes target the current `main` branch and the latest published
`@bilig/headless` runtime package set on npm. Older prerelease or unpublished
workspace states are not treated as supported release lines.

## Reporting A Vulnerability

Use GitHub's private vulnerability reporting flow from the repository security
page when it is available. If that flow is not visible for your account, open a
minimal GitHub issue that asks for a private maintainer contact and do not
include exploit details, secrets, private workbook data, or reproduction
artifacts in the public issue.

Please include:

- affected package or app
- affected version, commit, or npm package version
- impact and attack scenario
- minimal reproduction steps or a private proof artifact
- whether the report involves secret exposure, arbitrary code execution,
  formula evaluation, workbook persistence, import/export, sync transport, or
  agent execution

## Response Expectations

The maintainer response target is:

- initial triage within `7` days
- a fix, mitigation, or status update within `30` days for confirmed reports
- coordinated disclosure after a patched release or documented mitigation is
  available

If a report is not a security issue, it will be redirected to the normal GitHub
issue tracker.
