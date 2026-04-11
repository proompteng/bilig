import type { WorkbookAgentWorkflowRun } from "@bilig/contracts";

export interface WorkbookAgentFeatureFlags {
  readonly sharedThreadsEnabled: boolean;
  readonly workflowRunnerEnabled: boolean;
  readonly autoApplyLowRiskEnabled: boolean;
  readonly formulaWorkflowFamilyEnabled: boolean;
  readonly formattingWorkflowFamilyEnabled: boolean;
  readonly importWorkflowFamilyEnabled: boolean;
  readonly rollupWorkflowFamilyEnabled: boolean;
  readonly structuralWorkflowFamilyEnabled: boolean;
  readonly allowlistedUserIds: readonly string[];
  readonly allowlistedDocumentIds: readonly string[];
}

function parseBooleanEnv(value: string | undefined, fallback: boolean): boolean {
  if (!value) {
    return fallback;
  }
  const normalized = value.trim().toLowerCase();
  if (normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "on") {
    return true;
  }
  if (normalized === "false" || normalized === "0" || normalized === "no" || normalized === "off") {
    return false;
  }
  return fallback;
}

function parseCsvEnv(value: string | undefined): string[] {
  if (!value) {
    return [];
  }
  return value
    .split(",")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
}

export function resolveWorkbookAgentFeatureFlags(
  env: NodeJS.ProcessEnv = process.env,
): WorkbookAgentFeatureFlags {
  return {
    sharedThreadsEnabled: parseBooleanEnv(env["BILIG_AGENT_SHARED_THREADS_ENABLED"], true),
    workflowRunnerEnabled: parseBooleanEnv(env["BILIG_AGENT_WORKFLOW_RUNNER_ENABLED"], true),
    autoApplyLowRiskEnabled: parseBooleanEnv(env["BILIG_AGENT_AUTO_APPLY_LOW_RISK_ENABLED"], true),
    formulaWorkflowFamilyEnabled: parseBooleanEnv(
      env["BILIG_AGENT_FORMULA_WORKFLOWS_ENABLED"],
      true,
    ),
    formattingWorkflowFamilyEnabled: parseBooleanEnv(
      env["BILIG_AGENT_FORMATTING_WORKFLOWS_ENABLED"],
      true,
    ),
    importWorkflowFamilyEnabled: parseBooleanEnv(env["BILIG_AGENT_IMPORT_WORKFLOWS_ENABLED"], true),
    rollupWorkflowFamilyEnabled: parseBooleanEnv(env["BILIG_AGENT_ROLLUP_WORKFLOWS_ENABLED"], true),
    structuralWorkflowFamilyEnabled: parseBooleanEnv(
      env["BILIG_AGENT_STRUCTURAL_WORKFLOWS_ENABLED"],
      true,
    ),
    allowlistedUserIds: parseCsvEnv(env["BILIG_AGENT_ALLOWLIST_USERS"]),
    allowlistedDocumentIds: parseCsvEnv(env["BILIG_AGENT_ALLOWLIST_DOCUMENTS"]),
  };
}

export function isWorkbookAgentRolloutAllowed(
  featureFlags: Pick<WorkbookAgentFeatureFlags, "allowlistedUserIds" | "allowlistedDocumentIds">,
  input: { documentId: string; userId: string },
): boolean {
  const hasUserAllowlist = featureFlags.allowlistedUserIds.length > 0;
  const hasDocumentAllowlist = featureFlags.allowlistedDocumentIds.length > 0;
  if (!hasUserAllowlist && !hasDocumentAllowlist) {
    return true;
  }
  return (
    featureFlags.allowlistedUserIds.includes(input.userId) ||
    featureFlags.allowlistedDocumentIds.includes(input.documentId)
  );
}

export type WorkbookAgentWorkflowFamily =
  | "report"
  | "formula"
  | "formatting"
  | "import"
  | "rollup"
  | "structural";

export function getWorkbookAgentWorkflowFamily(
  workflowTemplate: WorkbookAgentWorkflowRun["workflowTemplate"],
): WorkbookAgentWorkflowFamily {
  switch (workflowTemplate) {
    case "summarizeWorkbook":
    case "summarizeCurrentSheet":
    case "describeRecentChanges":
    case "traceSelectionDependencies":
    case "explainSelectionCell":
    case "searchWorkbookQuery":
      return "report";
    case "findFormulaIssues":
    case "highlightFormulaIssues":
      return "formula";
    case "highlightCurrentSheetOutliers":
      return "formatting";
    case "normalizeCurrentSheetHeaders":
    case "normalizeCurrentSheetNumberFormats":
    case "normalizeCurrentSheetWhitespace":
    case "fillCurrentSheetFormulasDown":
      return "import";
    case "createCurrentSheetRollup":
      return "rollup";
    case "createSheet":
    case "renameCurrentSheet":
    case "hideCurrentRow":
    case "hideCurrentColumn":
    case "unhideCurrentRow":
    case "unhideCurrentColumn":
      return "structural";
  }
}
