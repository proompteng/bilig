#!/usr/bin/env bun

import { existsSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import * as ts from 'typescript'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../packages/excel-import/src/index.js'
import {
  requiresWorkbookAgentOwnerReview,
  resolveWorkbookAgentReviewDisposition,
} from '../packages/agent-api/src/workbook-agent-execution-policy.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'

export interface SecurityPostureControl {
  readonly id: string
  readonly category: 'formula-sandbox' | 'import-safety' | 'agent-permissions' | 'runtime-hardening'
  readonly required: boolean
  readonly passed: boolean
  readonly coveredControls: string[]
  readonly evidence: string
  readonly findings: string[]
}

export interface SecurityPostureScorecard {
  readonly schemaVersion: 1
  readonly suite: 'security-posture'
  readonly generatedAt: string
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-security-posture-scorecard.ts'
    readonly formulaRuntimeScanRoots: string[]
    readonly importImplementation: 'packages/excel-import/src/index.ts'
    readonly agentPolicyImplementation: 'packages/agent-api/src/workbook-agent-execution-policy.ts'
    readonly runtimePackageGate: 'pnpm publish:runtime:check'
  }
  readonly summary: {
    readonly allRequiredControlsPassed: boolean
    readonly formulaSandboxPassed: boolean
    readonly importSafetyPassed: boolean
    readonly agentPermissionPolicyPassed: boolean
    readonly runtimePackageHardeningPassed: boolean
    readonly coveredControls: string[]
    readonly uncoveredControls: string[]
    readonly externalGoogleSheetsEvidence: 'not-captured'
    readonly externalMicrosoftExcelEvidence: 'not-captured'
  }
  readonly controls: SecurityPostureControl[]
}

interface DynamicCodeFinding {
  readonly file: string
  readonly line: number
  readonly kind: string
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'security-posture-scorecard.json')
const formulaRuntimeScanRoots = ['packages/formula/src', 'packages/core/src/formula', 'packages/core/src/engine/services'] as const
const requiredControlIds = [
  'formula-runtime-no-dynamic-code-execution',
  'xlsx-import-macro-non-execution',
  'shared-agent-owner-review',
  'runtime-publish-package-hardening',
] as const
const coveredControlOrder = [
  'formula.noEval',
  'formula.noFunctionConstructor',
  'formula.noNodeProcessExecution',
  'xlsx.macroWarning',
  'xlsx.noMacroPayloadExport',
  'agent.sharedMediumHighRiskOwnerReview',
  'runtime.publishManifest',
  'runtime.noSourceInTarballs',
  'runtime.alignedPackageSet',
] as const
const uncoveredControls = [
  'browser.contentSecurityPolicy',
  'dependency.vulnerabilityAudit',
  'deployment.runtimeNetworkPolicy',
  'externalSheetsExcelSecurityComparison',
] as const
const disallowedImportModules = new Set(['node:child_process', 'child_process', 'node:vm', 'vm', 'bun:ffi'])
const formulaRuntimeServiceFileNames = new Set([
  'direct-formula-index-collection.ts',
  'direct-formula-recalc-helpers.ts',
  'formula-binding-service.ts',
  'formula-evaluation-service.ts',
  'formula-graph-service.ts',
  'formula-initialization-service.ts',
  'formula-template-normalization-service.ts',
  'operation-direct-formula-deltas.ts',
  'operation-direct-formula-values.ts',
  'runtime-column-store-service.ts',
])

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(`Security posture scorecard is missing. Run: bun scripts/gen-security-posture-scorecard.ts`)
    }
    const scorecard = parseSecurityPostureScorecard(JSON.parse(readFileSync(outputPath, 'utf8')) as unknown)
    validateSecurityPostureScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = buildSecurityPostureScorecard()
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export function buildSecurityPostureScorecard(generatedAt = new Date().toISOString()): SecurityPostureScorecard {
  const controls = [
    buildFormulaRuntimeSandboxControl(),
    buildXlsxImportSafetyControl(),
    buildAgentPermissionPolicyControl(),
    buildRuntimePackageHardeningControl(),
  ]
  const coveredControlSet = new Set(controls.flatMap((control) => control.coveredControls))
  const coveredControls = coveredControlOrder.filter((control) => coveredControlSet.has(control))

  return {
    schemaVersion: 1,
    suite: 'security-posture',
    generatedAt,
    source: {
      artifactGenerator: 'scripts/gen-security-posture-scorecard.ts',
      formulaRuntimeScanRoots: [...formulaRuntimeScanRoots],
      importImplementation: 'packages/excel-import/src/index.ts',
      agentPolicyImplementation: 'packages/agent-api/src/workbook-agent-execution-policy.ts',
      runtimePackageGate: 'pnpm publish:runtime:check',
    },
    summary: {
      allRequiredControlsPassed: controls.filter((control) => control.required).every((control) => control.passed),
      formulaSandboxPassed: requiredControl(controls, 'formula-runtime-no-dynamic-code-execution').passed,
      importSafetyPassed: requiredControl(controls, 'xlsx-import-macro-non-execution').passed,
      agentPermissionPolicyPassed: requiredControl(controls, 'shared-agent-owner-review').passed,
      runtimePackageHardeningPassed: requiredControl(controls, 'runtime-publish-package-hardening').passed,
      coveredControls,
      uncoveredControls: [...uncoveredControls],
      externalGoogleSheetsEvidence: 'not-captured',
      externalMicrosoftExcelEvidence: 'not-captured',
    },
    controls,
  }
}

function buildFormulaRuntimeSandboxControl(): SecurityPostureControl {
  const files = collectFormulaRuntimeFiles()
  const findings = files.flatMap(scanDynamicCodeFindings)
  return securityControl({
    id: 'formula-runtime-no-dynamic-code-execution',
    category: 'formula-sandbox',
    passed: findings.length === 0,
    coveredControls: ['formula.noEval', 'formula.noFunctionConstructor', 'formula.noNodeProcessExecution'],
    evidence: `Scanned ${String(files.length)} production formula/runtime TypeScript files for eval, Function constructors, process execution imports, and spawn calls.`,
    findings: findings.map((finding) => `${finding.file}:${String(finding.line)} ${finding.kind}`),
  })
}

function buildXlsxImportSafetyControl(): SecurityPostureControl {
  const macroWorkbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(macroWorkbook, XLSX.utils.aoa_to_sheet([['safe value']]), 'Sheet1')
  macroWorkbook.vbaraw = new Uint8Array([1, 2, 3, 4])
  const macroBytes = XLSX.write(macroWorkbook, { bookType: 'xlsm', type: 'buffer', bookVBA: true }) as unknown
  const importedMacroWorkbook = importXlsx(toUint8Array(macroBytes), 'macro.xlsm')
  const exportedBytes = exportXlsx(createSafeExportSnapshot())
  const importedExportedWorkbook = importXlsx(exportedBytes, 'safe.xlsx')
  const macroWarningPassed = importedMacroWorkbook.warnings.includes('Macros were ignored during XLSX import.')
  const noMacroExportPassed = !importedExportedWorkbook.warnings.includes('Macros were ignored during XLSX import.')

  return securityControl({
    id: 'xlsx-import-macro-non-execution',
    category: 'import-safety',
    passed: macroWarningPassed && noMacroExportPassed,
    coveredControls: ['xlsx.macroWarning', 'xlsx.noMacroPayloadExport'],
    evidence: 'Generated XLSM bytes with a VBA payload, verified import warns and supported XLSX export does not emit macro payloads.',
    findings: [
      ...(macroWarningPassed ? [] : ['macro-enabled workbook import did not emit the expected warning']),
      ...(noMacroExportPassed ? [] : ['supported XLSX export emitted a macro warning on re-import']),
    ],
  })
}

function buildAgentPermissionPolicyControl(): SecurityPostureControl {
  const sharedMediumReview =
    requiresWorkbookAgentOwnerReview({ scope: 'shared', riskClass: 'medium' }) &&
    resolveWorkbookAgentReviewDisposition({
      scope: 'shared',
      executionPolicy: 'autoApplyAll',
      riskClass: 'medium',
    }) === 'reviewQueue'
  const sharedHighReview =
    requiresWorkbookAgentOwnerReview({ scope: 'shared', riskClass: 'high' }) &&
    resolveWorkbookAgentReviewDisposition({
      scope: 'shared',
      executionPolicy: 'autoApplySafe',
      riskClass: 'high',
    }) === 'reviewQueue'

  return securityControl({
    id: 'shared-agent-owner-review',
    category: 'agent-permissions',
    passed: sharedMediumReview && sharedHighReview,
    coveredControls: ['agent.sharedMediumHighRiskOwnerReview'],
    evidence:
      'Executable policy checks require owner review for shared medium/high-risk workbook command bundles even under auto-apply policies.',
    findings: [
      ...(sharedMediumReview ? [] : ['shared medium-risk autoApplyAll bundle did not route to owner review']),
      ...(sharedHighReview ? [] : ['shared high-risk autoApplySafe bundle did not route to owner review']),
    ],
  })
}

function buildRuntimePackageHardeningControl(): SecurityPostureControl {
  const packageJson = readFileSync(join(rootDir, 'package.json'), 'utf8')
  const publishCheck = readFileSync(join(rootDir, 'scripts', 'check-package-publish.ts'), 'utf8')
  const runtimePackageSet = readFileSync(join(rootDir, 'scripts', 'runtime-package-set.ts'), 'utf8')
  const hasRuntimeGate =
    packageJson.includes('"publish:runtime:check"') &&
    packageJson.includes(
      '--require-aligned packages/protocol packages/workbook-domain packages/wasm-kernel packages/formula packages/core packages/headless',
    )
  const rejectsSource = publishCheck.includes('tarball must not contain source files') && publishCheck.includes('workspace:*')
  const alignedRuntimePackages =
    runtimePackageSet.includes('RUNTIME_PACKAGE_DIRS') &&
    runtimePackageSet.includes('assertAlignedVersions') &&
    runtimePackageSet.includes('packages/formula') &&
    runtimePackageSet.includes('packages/core') &&
    runtimePackageSet.includes('packages/headless')

  return securityControl({
    id: 'runtime-publish-package-hardening',
    category: 'runtime-hardening',
    passed: hasRuntimeGate && rejectsSource && alignedRuntimePackages,
    coveredControls: ['runtime.publishManifest', 'runtime.noSourceInTarballs', 'runtime.alignedPackageSet'],
    evidence:
      'Runtime publish check requires aligned package versions and validates packed tarballs omit source/test artifacts and workspace ranges.',
    findings: [
      ...(hasRuntimeGate ? [] : ['publish:runtime:check does not require the aligned runtime package set']),
      ...(rejectsSource ? [] : ['check-package-publish.ts does not reject source files or workspace ranges']),
      ...(alignedRuntimePackages ? [] : ['runtime-package-set.ts does not enforce aligned runtime package versions']),
    ],
  })
}

function securityControl(input: {
  readonly id: SecurityPostureControl['id']
  readonly category: SecurityPostureControl['category']
  readonly passed: boolean
  readonly coveredControls: readonly string[]
  readonly evidence: string
  readonly findings: readonly string[]
}): SecurityPostureControl {
  return {
    id: input.id,
    category: input.category,
    required: true,
    passed: input.passed,
    coveredControls: [...input.coveredControls],
    evidence: input.evidence,
    findings: [...input.findings],
  }
}

function collectFormulaRuntimeFiles(): string[] {
  const files = new Set<string>()
  for (const scanRoot of formulaRuntimeScanRoots) {
    collectTsFiles(join(rootDir, scanRoot), files)
  }
  return [...files].filter((file) => isProductionRuntimeFile(file)).toSorted()
}

function collectTsFiles(currentPath: string, output: Set<string>): void {
  const stats = statSync(currentPath)
  if (stats.isDirectory()) {
    for (const entry of readdirSync(currentPath)) {
      collectTsFiles(join(currentPath, entry), output)
    }
    return
  }
  if (currentPath.endsWith('.ts')) {
    output.add(currentPath)
  }
}

function isProductionRuntimeFile(filePath: string): boolean {
  const repoPath = toRepoPath(filePath)
  if (repoPath.includes('/__tests__/') || repoPath.includes('/generated/')) {
    return false
  }
  if (!repoPath.startsWith('packages/core/src/engine/services/')) {
    return true
  }
  const basename = repoPath.slice(repoPath.lastIndexOf('/') + 1)
  return formulaRuntimeServiceFileNames.has(basename)
}

function scanDynamicCodeFindings(filePath: string): DynamicCodeFinding[] {
  const source = readFileSync(filePath, 'utf8')
  const sourceFile = ts.createSourceFile(filePath, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS)
  const findings: DynamicCodeFinding[] = []
  const addFinding = (node: ts.Node, kind: string): void => {
    const position = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile))
    findings.push({
      file: toRepoPath(filePath),
      line: position.line + 1,
      kind,
    })
  }
  const visit = (node: ts.Node): void => {
    if (
      ts.isImportDeclaration(node) &&
      ts.isStringLiteral(node.moduleSpecifier) &&
      disallowedImportModules.has(node.moduleSpecifier.text)
    ) {
      addFinding(node, `disallowed import ${node.moduleSpecifier.text}`)
    } else if (ts.isCallExpression(node)) {
      if (ts.isIdentifier(node.expression) && node.expression.text === 'eval') {
        addFinding(node, 'eval call')
      } else if (ts.isIdentifier(node.expression) && node.expression.text === 'Function') {
        addFinding(node, 'Function constructor call')
      } else if (isDisallowedRequireCall(node)) {
        addFinding(node, 'disallowed require call')
      } else if (isDisallowedProcessExecutionCall(node.expression)) {
        addFinding(node, 'process execution call')
      }
    } else if (ts.isNewExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === 'Function') {
      addFinding(node, 'new Function constructor')
    }
    ts.forEachChild(node, visit)
  }
  visit(sourceFile)
  return findings
}

function isDisallowedRequireCall(node: ts.CallExpression): boolean {
  if (!ts.isIdentifier(node.expression) || node.expression.text !== 'require') {
    return false
  }
  const [specifier] = node.arguments
  return specifier !== undefined && ts.isStringLiteral(specifier) && disallowedImportModules.has(specifier.text)
}

function isDisallowedProcessExecutionCall(expression: ts.Expression): boolean {
  if (ts.isIdentifier(expression)) {
    return expression.text === 'spawn' || expression.text === 'spawnSync' || expression.text === 'execFile' || expression.text === 'exec'
  }
  if (ts.isPropertyAccessExpression(expression)) {
    const name = expression.name.text
    const receiver = expression.expression.getText()
    return receiver === 'Bun' && name === 'spawn'
  }
  return false
}

function createSafeExportSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'security-export',
    },
    sheets: [
      {
        id: 1,
        name: 'Sheet1',
        order: 0,
        cells: [{ address: 'A1', value: 'safe value' }],
      },
    ],
  }
}

function toUint8Array(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) {
    return new Uint8Array(value)
  }
  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value)
  }
  throw new Error('Expected workbook writer to return bytes')
}

function requiredControl(controls: readonly SecurityPostureControl[], id: string): SecurityPostureControl {
  const entry = controls.find((control) => control.id === id)
  if (!entry) {
    throw new Error(`Security posture scorecard is missing required control: ${id}`)
  }
  return entry
}

export function parseSecurityPostureScorecard(value: unknown): SecurityPostureScorecard {
  const record = toRecord(value, 'security posture scorecard')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'security-posture') {
    throw new Error('Unexpected security posture scorecard header')
  }
  const source = recordField(record, 'source', 'security posture source')
  const summary = recordField(record, 'summary', 'security posture summary')
  return {
    schemaVersion: 1,
    suite: 'security-posture',
    generatedAt: stringField(record, 'generatedAt', 'security posture generatedAt'),
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-security-posture-scorecard.ts'),
      formulaRuntimeScanRoots: stringArrayField(source, 'formulaRuntimeScanRoots', 'security posture formulaRuntimeScanRoots'),
      importImplementation: literalField(source, 'importImplementation', 'packages/excel-import/src/index.ts'),
      agentPolicyImplementation: literalField(
        source,
        'agentPolicyImplementation',
        'packages/agent-api/src/workbook-agent-execution-policy.ts',
      ),
      runtimePackageGate: literalField(source, 'runtimePackageGate', 'pnpm publish:runtime:check'),
    },
    summary: {
      allRequiredControlsPassed: booleanField(summary, 'allRequiredControlsPassed', 'security posture allRequiredControlsPassed'),
      formulaSandboxPassed: booleanField(summary, 'formulaSandboxPassed', 'security posture formulaSandboxPassed'),
      importSafetyPassed: booleanField(summary, 'importSafetyPassed', 'security posture importSafetyPassed'),
      agentPermissionPolicyPassed: booleanField(summary, 'agentPermissionPolicyPassed', 'security posture agentPermissionPolicyPassed'),
      runtimePackageHardeningPassed: booleanField(
        summary,
        'runtimePackageHardeningPassed',
        'security posture runtimePackageHardeningPassed',
      ),
      coveredControls: stringArrayField(summary, 'coveredControls', 'security posture coveredControls'),
      uncoveredControls: stringArrayField(summary, 'uncoveredControls', 'security posture uncoveredControls'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'not-captured'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'not-captured'),
    },
    controls: arrayField(record, 'controls', 'security posture controls').map(parseSecurityPostureControl),
  }
}

function parseSecurityPostureControl(value: unknown): SecurityPostureControl {
  const record = toRecord(value, 'security posture control')
  return {
    id: stringField(record, 'id', 'security posture control id'),
    category: parseSecurityPostureCategory(stringField(record, 'category', 'security posture control category')),
    required: booleanField(record, 'required', 'security posture control required'),
    passed: booleanField(record, 'passed', 'security posture control passed'),
    coveredControls: stringArrayField(record, 'coveredControls', 'security posture control coveredControls'),
    evidence: stringField(record, 'evidence', 'security posture control evidence'),
    findings: stringArrayField(record, 'findings', 'security posture control findings'),
  }
}

export function validateSecurityPostureScorecard(scorecard: SecurityPostureScorecard): void {
  for (const id of requiredControlIds) {
    const control = requiredControl(scorecard.controls, id)
    if (!control.required) {
      throw new Error(`Security posture scorecard required control is not marked required: ${id}`)
    }
    if (!control.passed) {
      throw new Error(`Security posture scorecard contains a failed required control: ${id}`)
    }
  }
  if (!scorecard.summary.allRequiredControlsPassed) {
    throw new Error('Security posture scorecard summary reports failed required controls')
  }
  for (const control of coveredControlOrder) {
    if (!scorecard.summary.coveredControls.includes(control)) {
      throw new Error(`Security posture scorecard is missing covered control: ${control}`)
    }
  }
  for (const control of uncoveredControls) {
    if (!scorecard.summary.uncoveredControls.includes(control)) {
      throw new Error(`Security posture scorecard is missing uncovered control disclosure: ${control}`)
    }
  }
}

function parseSecurityPostureCategory(value: string): SecurityPostureControl['category'] {
  if (value === 'formula-sandbox' || value === 'import-safety' || value === 'agent-permissions' || value === 'runtime-hardening') {
    return value
  }
  throw new Error(`Unexpected security posture category: ${value}`)
}

function logResult(mode: 'check' | 'write', scorecard: SecurityPostureScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredControlsPassed: scorecard.summary.allRequiredControlsPassed,
        coveredControls: scorecard.summary.coveredControls.length,
        uncoveredControls: scorecard.summary.uncoveredControls.length,
      },
      null,
      2,
    ),
  )
}

function recordField(value: Record<string, unknown>, field: string, name: string): Record<string, unknown> {
  return toRecord(value[field], name)
}

function arrayField(value: Record<string, unknown>, field: string, name: string): unknown[] {
  const fieldValue = value[field]
  if (!Array.isArray(fieldValue)) {
    throw new Error(`Expected ${name} to be an array`)
  }
  return fieldValue
}

function stringArrayField(value: Record<string, unknown>, field: string, name: string): string[] {
  const fieldValue = arrayField(value, field, name)
  if (!fieldValue.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected ${name} to contain only strings`)
  }
  return fieldValue
}

function stringField(value: Record<string, unknown>, field: string, name: string): string {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${name} to be a string`)
  }
  return fieldValue
}

function booleanField(value: Record<string, unknown>, field: string, name: string): boolean {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${name} to be a boolean`)
  }
  return fieldValue
}

function literalField<const T extends string>(value: Record<string, unknown>, field: string, expected: T): T {
  if (value[field] !== expected) {
    throw new Error(`Expected ${field} to be ${expected}`)
  }
  return expected
}

function toRecord(value: unknown, name: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${name} to be an object`)
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}

function toRepoPath(path: string): string {
  return relative(rootDir, path)
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'security-posture-scorecard-'))
  const tempFilePath = join(tempDir, 'scorecard.json')
  writeFileSync(tempFilePath, serializedJson)
  const oxfmtPath = join(rootDir, 'node_modules', '.bin', 'oxfmt')

  const formatResult = Bun.spawnSync([oxfmtPath, '--write', tempFilePath], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (formatResult.exitCode !== 0) {
    rmSync(tempDir, { recursive: true, force: true })
    throw new Error(`Unable to format generated security posture scorecard: ${new TextDecoder().decode(formatResult.stderr).trim()}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
