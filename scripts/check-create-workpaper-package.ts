#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'

type PackageManifest = Readonly<Record<string, unknown>>

const repoRoot = resolve(import.meta.dir, '..')
const packageDir = join(repoRoot, 'packages', 'create-workpaper')
const packageName = '@bilig/create-workpaper'
const packDir = join(repoRoot, 'build', 'create-workpaper-package')
const generatedDir = join(packDir, 'generated')
const args = new Set(process.argv.slice(2))
const requirePublished = args.has('--require-published')
const starterWorkpaperPath = './pricing.workpaper.json'
const existingRepoWorkpaperPath = './.bilig/pricing.workpaper.json'

function fail(message: string): never {
  throw new Error(message)
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    fail(message)
  }
}

function readJson(path: string): PackageManifest {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  assert(isRecord(parsed), `${path} must contain a JSON object`)
  return parsed
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function stringArray(value: unknown): readonly string[] {
  assert(Array.isArray(value) && value.every((entry) => typeof entry === 'string'), 'expected string array')
  return value
}

function run(command: string, commandArgs: readonly string[], options: { readonly allowFailure?: boolean; readonly cwd?: string } = {}) {
  const result = spawnSync(command, commandArgs, {
    cwd: options.cwd ?? repoRoot,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  if (result.status !== 0 && !options.allowFailure) {
    const stderr = result.stderr.trim()
    const stdout = result.stdout.trim()
    fail(`${command} ${commandArgs.join(' ')} failed\n${stderr || stdout}`)
  }
  return result
}

function assertManifest(manifest: PackageManifest): string {
  assert(manifest.name === packageName, `packages/create-workpaper/package.json.name must be ${packageName}`)
  assert(typeof manifest.version === 'string' && /^\d+\.\d+\.\d+$/u.test(manifest.version), 'package version must be stable semver')
  assert(
    manifest.description === 'Create a runnable Bilig WorkPaper starter for Node services and tool integrations.',
    'package description must keep the starter promise specific',
  )
  assert(manifest.homepage === 'https://proompteng.github.io/bilig/', 'package homepage must point to public docs')
  assert(manifest.license === 'MIT', 'package license must be MIT')
  assert(manifest.type === 'module', 'package type must be module')

  const keywords = stringArray(manifest.keywords)
  for (const keyword of [
    'formula-engine',
    'formula-workbook',
    'mcp-tools',
    'node-services',
    'server-side-formulas',
    'typescript',
    'workbook-api',
    'workbook-formulas',
    'workbook-runtime',
    'workpaper',
    'workpaper-json',
    'workpaper-runtime',
  ]) {
    assert(keywords.includes(keyword), `package keywords must include ${keyword}`)
  }
  for (const keyword of ['spreadsheet-formulas', 'spreadsheet-automation', 'xlsx']) {
    assert(!keywords.includes(keyword), `package keywords must not include stale keyword ${keyword}`)
  }

  assert(isRecord(manifest.bugs), 'package bugs must be an object')
  assert(manifest.bugs['url'] === 'https://github.com/proompteng/bilig/issues', 'package bugs.url must point to GitHub issues')

  assert(isRecord(manifest.repository), 'package repository must be an object')
  assert(manifest.repository['type'] === 'git', 'package repository.type must be git')
  assert(manifest.repository['url'] === 'git+https://github.com/proompteng/bilig.git', 'package repository.url must point to GitHub')
  assert(manifest.repository['directory'] === 'packages/create-workpaper', 'package repository.directory must be packages/create-workpaper')

  assert(isRecord(manifest.bin), 'package bin must be an object')
  assert(manifest.bin['create-workpaper'] === 'bin/create-bilig-workpaper.js', 'package bin must expose create-workpaper')

  const files = stringArray(manifest.files)
  for (const included of ['agent-overlay', 'bin', 'template', 'README.md']) {
    assert(files.includes(included), `package files must include ${included}`)
  }

  assert(isRecord(manifest.publishConfig), 'package publishConfig must be an object')
  assert(manifest.publishConfig['access'] === 'public', 'package publishConfig.access must be public')

  assert(isRecord(manifest.engines), 'package engines must be an object')
  assert(manifest.engines['node'] === '>=22.0.0', 'package engines.node must stay Node 22 compatible')

  assert(isRecord(manifest.scripts), 'package scripts must be an object')
  assert(manifest.scripts['smoke'] === 'node ./bin/create-bilig-workpaper.js --help', 'package smoke script must exercise the CLI')

  return manifest.version
}

function assertDocs(): void {
  const readme = readFileSync(join(packageDir, 'README.md'), 'utf8')
  const docs = readFileSync(join(repoRoot, 'docs', 'create-bilig-workpaper.md'), 'utf8')
  const agentDocs = readFileSync(join(repoRoot, 'docs', 'agent-adoption-kit.md'), 'utf8')
  const templateSource = readFileSync(join(packageDir, 'template', 'src', 'index.ts'), 'utf8')
  for (const [label, source] of [
    ['packages/create-workpaper/README.md', readme],
    ['docs/create-bilig-workpaper.md', docs],
  ] as const) {
    assert(source.includes('The restored-formula fix is in source'), `${label} must guard the broken 0.164.11 starter`)
    assert(source.includes('npm create @bilig/workpaper@latest'), `${label} must document the scoped npm create path`)
    assert(source.includes('@bilig/create-workpaper'), `${label} must include the published package name`)
    assert(source.includes('--agent'), `${label} must document the agent-ready starter path`)
    assert(source.includes('--add-agent'), `${label} must document the existing-repo agent overlay path`)
  }
  assert(readme.includes('verified: true') || readme.includes('"verified": true'), 'starter README must show the verification output')
  assert(docs.includes('verified: true') || docs.includes('"verified": true'), 'starter docs must show the verification output')
  assert(
    templateSource.includes('assertSmokeOutput(output)') &&
      templateSource.indexOf('assertSmokeOutput(output)') < templateSource.indexOf('console.log(JSON.stringify(output, null, 2))'),
    'starter smoke output must be verified before it is printed',
  )
  assert(templateSource.includes('buildA1WorkPaper'), 'starter template must use the A1 WorkPaper facade')
  assert(templateSource.includes('editManyAndReadback'), 'starter template must use atomic editManyAndReadback proof')
  assert(templateSource.includes('validateFormula'), 'starter template must validate formulas through the A1 facade')
  assert(!templateSource.includes('getSheetId'), 'starter template must not teach sheet-id lookup')
  assert(!templateSource.includes('setCellContents({ sheet'), 'starter template must not teach zero-based cell writes')
  assert(!templateSource.includes('nextStep'), 'starter smoke output must not include CTA metadata in the JSON proof')
  assert(!templateSource.includes('https://github.com/proompteng/bilig/stargazers'), 'starter smoke output must not include the star link')
  assert(
    !templateSource.includes('https://github.com/proompteng/bilig/subscription'),
    'starter smoke output must not include the release-watch link',
  )
  assert(
    readme.includes('https://github.com/proompteng/bilig/discussions/new?category=general'),
    'starter README must include the adoption-feedback link',
  )
  assert(!readme.includes('https://github.com/proompteng/bilig/stargazers'), 'starter README must not include the star link')
  assert(
    readme.includes('https://github.com/proompteng/bilig/subscription'),
    'starter README must include the release-watch link after proof guidance',
  )
  assert(
    docs.includes('https://github.com/proompteng/bilig/discussions/new?category=general'),
    'starter docs must include the adoption-feedback link',
  )
  assert(!docs.includes('https://github.com/proompteng/bilig/stargazers'), 'starter docs must not include the star link')
  assert(
    docs.includes('https://github.com/proompteng/bilig/subscription'),
    'starter docs must include the release-watch link after proof guidance',
  )
  assert(readme.includes('agent:verify'), 'starter README must document the agent verification script')
  assert(docs.includes('agent:verify'), 'starter docs must document the agent verification script')
  assert(readme.includes('Kiro'), 'starter README must document the Kiro agent overlay')
  assert(docs.includes('.kiro/settings/mcp.json'), 'starter docs must document the Kiro MCP config')
  assert(readme.includes('Trae') && readme.includes('.trae/mcp.json'), 'starter README must document the Trae MCP config')
  assert(docs.includes('.trae/mcp.json'), 'starter docs must document the Trae MCP config')
  assert(agentDocs.includes('.trae/rules/bilig-workpaper.md'), 'agent handoff checklist must document the Trae project rule')
  assert(readme.includes('Zed') && readme.includes('.zed/settings.json'), 'starter README must document the Zed MCP config')
  assert(docs.includes('.zed/settings.json'), 'starter docs must document the Zed MCP config')
  assert(agentDocs.includes('.zed/settings.json'), 'agent handoff checklist must document the Zed MCP config')
  assert(agentDocs.includes('.kiro/steering/bilig-workpaper.md'), 'agent handoff checklist must document the Kiro steering file')
  assert(
    readme.includes('bilig-evaluate --door agent-mcp --scenario revenue-plan --json'),
    'starter README must document the revenue-plan agent evaluator command',
  )
  assert(
    docs.includes('bilig-evaluate --door agent-mcp --scenario revenue-plan --json'),
    'starter docs must document the revenue-plan agent evaluator command',
  )
  assert(
    readme.includes('SUMIF') && readme.includes('XLOOKUP') && readme.includes('FILTER'),
    'starter README must disclose the formula families covered by the revenue-plan evaluator',
  )
  assert(
    docs.includes('SUMIF') && docs.includes('XLOOKUP') && docs.includes('FILTER'),
    'starter docs must disclose the formula families covered by the revenue-plan evaluator',
  )
  assert(
    readme.includes('npm run mcp:challenge') && readme.includes('lower-level JSON-RPC transcript'),
    'starter README must keep the raw MCP challenge as a diagnostic, not the primary proof',
  )
  assert(
    docs.includes('npm run mcp:challenge') && docs.includes('lower-level JSON-RPC'),
    'starter docs must keep the raw MCP challenge as a diagnostic, not the primary proof',
  )
  assert(readme.includes(existingRepoWorkpaperPath), 'starter README must document the existing-repo WorkPaper state path')
  assert(docs.includes(existingRepoWorkpaperPath), 'starter docs must document the existing-repo WorkPaper state path')
  assert(agentDocs.includes(existingRepoWorkpaperPath), 'agent handoff checklist must document the existing-repo WorkPaper state path')
}

function assertPackedTarball(): void {
  rmSync(packDir, { force: true, recursive: true })
  mkdirSync(packDir, { recursive: true })
  run('pnpm', ['--filter', packageName, 'pack', '--pack-destination', packDir])

  const tarballs = readdirSync(packDir).filter((entry) => entry.endsWith('.tgz'))
  assert(tarballs.length === 1, `expected one packed tarball in ${packDir}, found ${String(tarballs.length)}`)
  const tarball = join(packDir, tarballs[0] ?? fail('missing tarball'))
  const tarOutput = run('tar', ['-tf', tarball]).stdout
  const entries = new Set(tarOutput.trim().split('\n').filter(Boolean))
  for (const entry of [
    'package/package.json',
    'package/README.md',
    'package/agent-overlay/.agents/skills/bilig-workpaper/SKILL.md',
    'package/agent-overlay/.aider.conf.yml',
    'package/agent-overlay/.claude/commands/bilig-workpaper-proof.md',
    'package/agent-overlay/.claude/skills/bilig-workpaper/SKILL.md',
    'package/agent-overlay/.clinerules/bilig-workpaper.md',
    'package/agent-overlay/.continue/mcpServers/bilig-workpaper.yaml',
    'package/agent-overlay/.continue/rules/bilig-workpaper.md',
    'package/agent-overlay/.devin/rules/bilig-workpaper.md',
    'package/agent-overlay/.kiro/settings/mcp.json',
    'package/agent-overlay/.kiro/steering/bilig-workpaper.md',
    'package/agent-overlay/.mcp.json',
    'package/agent-overlay/.cursor/mcp.json',
    'package/agent-overlay/.cursor/rules/bilig-workpaper.mdc',
    'package/agent-overlay/.roo/mcp.json',
    'package/agent-overlay/.roo/rules/bilig-workpaper.md',
    'package/agent-overlay/.trae/mcp.json',
    'package/agent-overlay/.trae/rules/bilig-workpaper.md',
    'package/agent-overlay/.zed/settings.json',
    'package/agent-overlay/.github/copilot-instructions.md',
    'package/agent-overlay/.github/instructions/bilig-workpaper.instructions.md',
    'package/agent-overlay/.github/prompts/bilig-workpaper-proof.prompt.md',
    'package/agent-overlay/.junie/mcp/mcp.json',
    'package/agent-overlay/.opencode/agents/bilig-workpaper.md',
    'package/agent-overlay/.vscode/mcp.json',
    'package/agent-overlay/.windsurf/rules/bilig-workpaper.md',
    'package/agent-overlay/AGENTS.md',
    'package/agent-overlay/CLAUDE.md',
    'package/agent-overlay/CONVENTIONS.md',
    'package/agent-overlay/GEMINI.md',
    'package/agent-overlay/README.md',
    'package/agent-overlay/mcp/bilig-workpaper.mcp.json',
    'package/agent-overlay/opencode.jsonc',
    'package/agent-overlay/package.json',
    'package/bin/create-bilig-workpaper.js',
    'package/template/package.json',
    'package/template/README.md',
    'package/template/src/index.ts',
    'package/template/tsconfig.json',
  ]) {
    assert(entries.has(entry), `${basename(tarball)} is missing ${entry}`)
  }

  for (const entry of entries) {
    assert(!entry.includes('node_modules/'), `${basename(tarball)} must not include node_modules`)
    assert(!entry.includes('.cache/'), `${basename(tarball)} must not include .cache output`)
  }
}

function assertGeneratedStarters(): void {
  rmSync(generatedDir, { force: true, recursive: true })
  mkdirSync(generatedDir, { recursive: true })

  const cliPath = join(packageDir, 'bin', 'create-bilig-workpaper.js')
  const serviceDir = join(generatedDir, 'service-demo')
  const agentDir = join(generatedDir, 'agent-demo')
  const existingDir = join(generatedDir, 'existing-demo')
  const dotExistingDir = join(generatedDir, 'dot-existing-demo')

  run('node', [cliPath, serviceDir])
  run('node', [cliPath, agentDir, '--agent'])
  mkdirSync(existingDir, { recursive: true })
  writeFileSync(
    join(existingDir, 'package.json'),
    `${JSON.stringify({ name: 'existing-demo', private: true, scripts: { test: 'node --test' } }, null, 2)}\n`,
  )
  writeFileSync(join(existingDir, 'README.md'), '# Existing project\n')
  writeFileSync(join(existingDir, 'src-index.ts'), 'console.log("keep me")\n')
  run('node', [cliPath, existingDir, '--add-agent'])
  mkdirSync(dotExistingDir, { recursive: true })
  writeFileSync(
    join(dotExistingDir, 'package.json'),
    `${JSON.stringify({ name: '@demo/pricing-rules-service', private: true }, null, 2)}\n`,
  )
  writeFileSync(join(dotExistingDir, 'README.md'), '# Dot existing project\n')
  run('node', [cliPath, '.', '--add-agent'], { cwd: dotExistingDir })

  const serviceManifest = readJson(join(serviceDir, 'package.json'))
  assert(isRecord(serviceManifest.scripts), 'generated service package scripts must be an object')
  assert(serviceManifest.scripts['smoke'] === 'tsx src/index.ts', 'generated service starter must keep the smoke script')
  assert(serviceManifest.scripts['agent:verify'] === undefined, 'generated service starter must not include agent-only scripts')
  assertGeneratedDependencyVersions(serviceManifest, version, 'service')

  run('pnpm', ['--filter', '@bilig/workpaper', 'build'])
  const serviceScopedModules = join(serviceDir, 'node_modules', '@bilig')
  mkdirSync(serviceScopedModules, { recursive: true })
  symlinkSync(
    join(repoRoot, 'packages', 'workpaper'),
    join(serviceScopedModules, 'workpaper'),
    process.platform === 'win32' ? 'junction' : 'dir',
  )
  const serviceSmokeResult = run('pnpm', ['exec', 'tsx', 'src/index.ts'], { cwd: serviceDir })
  const serviceSmoke: unknown = JSON.parse(serviceSmokeResult.stdout)
  assert(isRecord(serviceSmoke), 'generated service starter smoke must return a JSON object')
  assert(serviceSmoke.verified === true, 'generated service starter smoke must return verified true')
  assert(isRecord(serviceSmoke.edit), 'generated service starter smoke must return edit evidence')
  assert(isRecord(serviceSmoke.edit.checks), 'generated service starter smoke must return edit checks')
  assert(serviceSmoke.edit.checks.formulasPersisted === true, 'generated service starter must prove formulas persisted')
  assert(serviceSmoke.edit.checks.restoredMatchesAfter === true, 'generated service starter must prove restored readback')

  const agentManifest = readJson(join(agentDir, 'package.json'))
  assert(isRecord(agentManifest.scripts), 'generated agent package scripts must be an object')
  assertGeneratedDependencyVersions(agentManifest, version, 'agent')
  assert(
    agentManifest.scripts['agent:verify'] === 'npm run smoke && npm run agent:evaluate:basic && npm run agent:evaluate',
    'generated agent starter must verify the service smoke, basic evaluator, and revenue-plan evaluator paths',
  )
  assert(
    agentManifest.scripts['agent:evaluate'] === 'bilig-evaluate --door agent-mcp --scenario revenue-plan --json',
    'generated agent starter must expose the revenue-plan agent evaluator script',
  )
  assert(
    agentManifest.scripts['agent:evaluate:basic'] === 'bilig-evaluate --door agent-mcp --json',
    'generated agent starter must keep the basic agent evaluator script',
  )
  assert(
    agentManifest.scripts['mcp:challenge'] === 'bilig-mcp-challenge --json',
    'generated agent starter must keep the raw MCP challenge as a JSON diagnostic script',
  )
  assert(
    agentManifest.scripts['mcp:server'] === `bilig-workpaper-mcp --workpaper ${starterWorkpaperPath} --init-demo-workpaper --writable`,
    'generated agent starter must include the file-backed MCP server script',
  )

  for (const expected of [
    'AGENTS.md',
    'CLAUDE.md',
    'GEMINI.md',
    '.claude/commands/bilig-workpaper-proof.md',
    '.claude/skills/bilig-workpaper/SKILL.md',
    '.agents/skills/bilig-workpaper/SKILL.md',
    '.aider.conf.yml',
    '.clinerules/bilig-workpaper.md',
    '.continue/mcpServers/bilig-workpaper.yaml',
    '.continue/rules/bilig-workpaper.md',
    '.devin/rules/bilig-workpaper.md',
    '.kiro/settings/mcp.json',
    '.kiro/steering/bilig-workpaper.md',
    '.mcp.json',
    '.cursor/mcp.json',
    '.cursor/rules/bilig-workpaper.mdc',
    '.roo/mcp.json',
    '.roo/rules/bilig-workpaper.md',
    '.trae/mcp.json',
    '.trae/rules/bilig-workpaper.md',
    '.zed/settings.json',
    '.github/copilot-instructions.md',
    '.github/instructions/bilig-workpaper.instructions.md',
    '.github/prompts/bilig-workpaper-proof.prompt.md',
    '.junie/mcp/mcp.json',
    '.opencode/agents/bilig-workpaper.md',
    '.vscode/mcp.json',
    '.windsurf/rules/bilig-workpaper.md',
    'CONVENTIONS.md',
    'mcp/bilig-workpaper.mcp.json',
    'opencode.jsonc',
  ]) {
    assert(existsSync(join(agentDir, expected)), `generated agent starter is missing ${expected}`)
  }

  const generatedAgentInstructions = readFileSync(join(agentDir, 'AGENTS.md'), 'utf8')
  assert(
    generatedAgentInstructions.includes('set_cell_contents_and_readback'),
    'generated agent starter must teach the composite MCP edit/readback tool',
  )
  assert(
    generatedAgentInstructions.includes('bilig-evaluate --door agent-mcp --scenario revenue-plan --json'),
    'generated agent starter must teach the revenue-plan evaluator command',
  )
  assert(
    readFileSync(join(agentDir, '.aider.conf.yml'), 'utf8').includes('- CONVENTIONS.md'),
    'generated agent starter Aider config must load CONVENTIONS.md',
  )

  for (const expected of [
    '.agents/skills/bilig-workpaper/SKILL.md',
    '.mcp.json',
    '.cursor/mcp.json',
    '.continue/mcpServers/bilig-workpaper.yaml',
    '.junie/mcp/mcp.json',
    '.kiro/settings/mcp.json',
    '.kiro/steering/bilig-workpaper.md',
    '.roo/mcp.json',
    '.roo/rules/bilig-workpaper.md',
    '.trae/mcp.json',
    '.trae/rules/bilig-workpaper.md',
    '.zed/settings.json',
    '.vscode/mcp.json',
    'mcp/bilig-workpaper.mcp.json',
    'CONVENTIONS.md',
    'README.md',
  ]) {
    const generatedSource = readFileSync(join(agentDir, expected), 'utf8')
    assert(generatedSource.includes(starterWorkpaperPath), `generated agent starter ${expected} must use the starter WorkPaper path`)
    assert(!generatedSource.includes('__WORKPAPER_PATH__'), `generated agent starter ${expected} must render WorkPaper path placeholders`)
  }

  const existingManifest = readJson(join(existingDir, 'package.json'))
  assert(existingManifest.name === 'existing-demo', 'existing-repo overlay must preserve package name')
  assert(isRecord(existingManifest.scripts), 'existing-repo package scripts must be an object')
  assert(existingManifest.scripts['test'] === 'node --test', 'existing-repo overlay must preserve existing scripts')
  assert(existingManifest.scripts['agent:verify'] === undefined, 'existing-repo overlay must not mutate package scripts')
  assert(existingManifest.scripts['mcp:server'] === undefined, 'existing-repo overlay must not add package scripts')
  assert(existingManifest.dependencies === undefined, 'existing-repo overlay must not mutate dependencies')
  assert(
    readFileSync(join(existingDir, 'README.md'), 'utf8') === '# Existing project\n',
    'existing-repo overlay must not overwrite README.md',
  )
  assert(
    readFileSync(join(existingDir, 'src-index.ts'), 'utf8') === 'console.log("keep me")\n',
    'existing-repo overlay must not touch unrelated app files',
  )
  assert(!existsSync(join(existingDir, 'src', 'index.ts')), 'existing-repo overlay must not copy the starter template')
  assert(existsSync(join(existingDir, 'BILIG_WORKPAPER.md')), 'existing-repo overlay must write BILIG_WORKPAPER.md')
  assert(existsSync(join(existingDir, 'AGENTS.md')), 'existing-repo overlay must write AGENTS.md when absent')
  assert(
    existsSync(join(existingDir, '.claude', 'skills', 'bilig-workpaper', 'SKILL.md')),
    'existing-repo overlay must write the Claude Code project skill',
  )
  assert(
    existsSync(join(existingDir, '.agents', 'skills', 'bilig-workpaper', 'SKILL.md')),
    'existing-repo overlay must write the OpenHands project skill',
  )
  assert(existsSync(join(existingDir, 'CONVENTIONS.md')), 'existing-repo overlay must write Aider conventions')
  assert(existsSync(join(existingDir, '.aider.conf.yml')), 'existing-repo overlay must write Aider config')
  assert(
    readFileSync(join(existingDir, '.aider.conf.yml'), 'utf8').includes('- CONVENTIONS.md'),
    'existing-repo overlay Aider config must load CONVENTIONS.md',
  )
  assert(
    existsSync(join(existingDir, '.opencode', 'agents', 'bilig-workpaper.md')),
    'existing-repo overlay must write the OpenCode project agent',
  )
  assert(
    existsSync(join(existingDir, '.github', 'instructions', 'bilig-workpaper.instructions.md')),
    'existing-repo overlay must write path-specific Copilot instructions',
  )
  assert(
    readFileSync(join(existingDir, '.mcp.json'), 'utf8').includes('npm') &&
      readFileSync(join(existingDir, '.mcp.json'), 'utf8').includes('exec') &&
      readFileSync(join(existingDir, '.mcp.json'), 'utf8').includes('@bilig/workpaper@latest'),
    'existing-repo MCP config must use direct npm exec instead of project scripts',
  )
  assert(!existsSync(join(existingDir, '.bilig')), 'existing-repo overlay must not create WorkPaper state before the MCP server runs')
  for (const expected of [
    '.mcp.json',
    '.agents/skills/bilig-workpaper/SKILL.md',
    'CONVENTIONS.md',
    '.cursor/mcp.json',
    '.junie/mcp/mcp.json',
    '.kiro/settings/mcp.json',
    '.kiro/steering/bilig-workpaper.md',
    '.roo/mcp.json',
    '.roo/rules/bilig-workpaper.md',
    '.trae/mcp.json',
    '.trae/rules/bilig-workpaper.md',
    '.zed/settings.json',
    '.opencode/agents/bilig-workpaper.md',
    '.vscode/mcp.json',
    'mcp/bilig-workpaper.mcp.json',
    'opencode.jsonc',
    'BILIG_WORKPAPER.md',
    'AGENTS.md',
  ]) {
    const generatedSource = readFileSync(join(existingDir, expected), 'utf8')
    assert(generatedSource.includes(existingRepoWorkpaperPath), `existing-repo overlay ${expected} must use the hidden WorkPaper path`)
    assert(!generatedSource.includes(starterWorkpaperPath), `existing-repo overlay ${expected} must not use the root WorkPaper path`)
    assert(!generatedSource.includes('__WORKPAPER_PATH__'), `existing-repo overlay ${expected} must render WorkPaper path placeholders`)
  }

  const existingAgentNotes = readFileSync(join(existingDir, 'AGENTS.md'), 'utf8')
  writeFileSync(join(existingDir, 'AGENTS.md'), '# Existing agent policy\n')
  run('node', [cliPath, existingDir, '--add-agent'])
  assert(
    readFileSync(join(existingDir, 'AGENTS.md'), 'utf8') === '# Existing agent policy\n',
    'existing-repo overlay must skip existing agent files by default',
  )
  const installSummary = readFileSync(join(existingDir, 'BILIG_WORKPAPER_INSTALL.md'), 'utf8')
  assert(
    installSummary.includes('AGENTS.md') &&
      installSummary.includes('BILIG_WORKPAPER.md') &&
      installSummary.includes('bilig-evaluate --door agent-mcp --scenario revenue-plan --json'),
    'existing-repo overlay must write a handoff summary when it skips existing agent files',
  )
  run('node', [cliPath, existingDir, '--add-agent'])
  assert(
    readFileSync(join(existingDir, 'BILIG_WORKPAPER_INSTALL.md'), 'utf8') === installSummary,
    'existing-repo skipped-file handoff summary must be idempotent',
  )
  run('node', [cliPath, existingDir, '--add-agent', '--force'])
  assert(
    readFileSync(join(existingDir, 'AGENTS.md'), 'utf8') === existingAgentNotes,
    'existing-repo overlay --force must overwrite overlay files',
  )
  assert(readFileSync(join(existingDir, 'README.md'), 'utf8') === '# Existing project\n', '--force must still not overwrite README.md')
  assert(
    JSON.stringify(readJson(join(existingDir, 'package.json'))) === JSON.stringify(existingManifest),
    '--force must still not mutate package.json',
  )
  assert(
    readFileSync(join(dotExistingDir, 'BILIG_WORKPAPER.md'), 'utf8').startsWith('# `pricing-rules-service`'),
    'existing-repo dot overlay must derive the project label from package.json instead of rendering "."',
  )
  assert(
    readFileSync(join(dotExistingDir, 'README.md'), 'utf8') === '# Dot existing project\n',
    'existing-repo dot overlay must not overwrite README.md',
  )
  assert(
    JSON.stringify(readJson(join(dotExistingDir, 'package.json'))) ===
      JSON.stringify({ name: '@demo/pricing-rules-service', private: true }),
    'existing-repo dot overlay must not mutate package.json',
  )
  assert(
    !existsSync(join(dotExistingDir, '.bilig')),
    'existing-repo dot overlay must not create WorkPaper state before the MCP server runs',
  )
}

function assertGeneratedDependencyVersions(manifest: PackageManifest, version: string, label: string): void {
  assert(isRecord(manifest.dependencies), `generated ${label} package dependencies must be an object`)
  assert(manifest.dependencies['@bilig/workpaper'] === version, `generated ${label} package must pin @bilig/workpaper to ${version}`)
  assert(isRecord(manifest.devDependencies), `generated ${label} package devDependencies must be an object`)
  for (const dependencyName of ['@types/node', 'tsx', 'typescript']) {
    const dependencyVersion = manifest.devDependencies[dependencyName]
    assert(
      typeof dependencyVersion === 'string' && dependencyVersion !== 'latest' && /^\d+\.\d+\.\d+$/u.test(dependencyVersion),
      `generated ${label} package must pin ${dependencyName} instead of using latest`,
    )
  }
}

function assertPublishedVersion(version: string): void {
  const result = run('npm', ['view', `${packageName}@${version}`, 'version', '--json'], { allowFailure: true })
  if (result.status === 0 && result.stdout.includes(version)) {
    console.log(`${packageName}@${version} is published on npm`)
    return
  }

  const message = `${packageName}@${version} is not published on npm yet`
  if (requirePublished) {
    fail(message)
  }
  console.warn(`${message}; run with --require-published in release verification after the npm package exists.`)
}

const manifest = readJson(join(packageDir, 'package.json'))
const version = assertManifest(manifest)

assert(existsSync(join(packageDir, 'bin', 'create-bilig-workpaper.js')), 'CLI entrypoint must exist')
assert(existsSync(join(packageDir, 'template', 'src', 'index.ts')), 'starter template source must exist')
assert(existsSync(join(packageDir, 'agent-overlay', 'AGENTS.md')), 'agent starter overlay must exist')
assertDocs()
assertGeneratedStarters()
assertPackedTarball()
assertPublishedVersion(version)

console.log(`${packageName}@${version} package checks passed`)
