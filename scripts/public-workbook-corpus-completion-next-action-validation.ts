import { publicCorpusStopMarkerOverrideEnvVar, publicCorpusStopMarkerOverrideFlag } from './public-workbook-corpus-cli.ts'
import { pnpmScriptName } from './public-workbook-corpus-completion-audit-helpers.ts'
import type { PublicWorkbookCorpusCompletionAudit } from './public-workbook-corpus-completion-audit-types.ts'

export function validatePublicWorkbookCorpusAuditNextActions(args: {
  readonly audit: PublicWorkbookCorpusCompletionAudit
  readonly packageScripts: ReadonlyMap<string, string>
}): string[] {
  const findings: string[] = []
  for (const action of args.audit.nextActions) {
    if (!action.id.trim()) {
      findings.push('next action is missing an id')
    }
    if (!Number.isFinite(action.priority) || action.priority <= 0) {
      findings.push(`${action.id} next action has invalid priority`)
    }
    if (!action.reason.trim()) {
      findings.push(`${action.id} next action is missing a reason`)
    }
    if (action.commands.length === 0) {
      findings.push(`${action.id} next action has no commands`)
    }
    for (const command of action.commands) {
      if (args.audit.completionVerdict.nextCorpusRunRequiresExplicitResume && bypassesActiveStopMarker(command)) {
        findings.push(`${action.id} runnable next action command bypasses the active corpus stop marker: ${command}`)
      }
      validatePnpmScriptCommand(`${action.id} next action command`, command, args.packageScripts, findings)
    }
    for (const command of action.blockedCommands) {
      validatePnpmScriptCommand(`${action.id} blocked command`, command, args.packageScripts, findings)
    }
  }
  return findings
}

function validatePnpmScriptCommand(label: string, command: string, packageScripts: ReadonlyMap<string, string>, findings: string[]): void {
  const scriptName = pnpmScriptName(command)
  if (!scriptName) {
    findings.push(`${label} is not a pnpm package script: ${command}`)
  } else if (!packageScripts.has(scriptName)) {
    findings.push(`${label} references missing package script: ${scriptName}`)
  }
}

function bypassesActiveStopMarker(command: string): boolean {
  return command.includes(publicCorpusStopMarkerOverrideFlag) || command.includes(`${publicCorpusStopMarkerOverrideEnvVar}=1`)
}
