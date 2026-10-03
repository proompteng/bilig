# Security Policy

Security reports are handled separately from normal bug reports so sensitive
details do not become public before a fix is available.

`bilig` handles workbook data, formulas, import/export fixtures, local
persistence, browser runtime state, service-side WorkPaper documents, sync
transport, and agent-facing APIs.

## Supported Versions

Security fixes target the current `main` branch and the latest published
`@bilig/workpaper` runtime package set on npm. Older prerelease or unpublished
workspace states are not treated as supported release lines.

## Reporting A Vulnerability

Use GitHub's private vulnerability reporting flow from the repository security
page when it is available. If that flow is not visible for your account, email
the maintainer at <security@proompteng.ai>.

Do not include exploit details, secrets, private workbook data, customer data,
tokens, credentials, or reproduction artifacts in public issues.

Please include:

- affected package or app
- affected version, commit, or npm package version
- impact and attack scenario
- minimal reproduction steps or a private proof artifact using synthetic data
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

## Public Disclosure

Please wait for maintainer confirmation before publishing details. Public
issues are fine for ordinary correctness bugs, formula parity gaps, import/export
fidelity gaps, documentation fixes, and reproducible crashes that do not expose
private data or bypass security boundaries.
