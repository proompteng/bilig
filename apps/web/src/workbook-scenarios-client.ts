import {
  workbookScenarioCreateRequestSchema,
  workbookScenarioResponseSchema,
  type WorkbookScenarioResponse,
} from "@bilig/zero-sync";

interface WorkbookScenarioCreateInput {
  readonly documentId: string;
  readonly name: string;
  readonly sheetName?: string;
  readonly address?: string;
  readonly viewport?: {
    readonly rowStart: number;
    readonly rowEnd: number;
    readonly colStart: number;
    readonly colEnd: number;
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

async function readErrorMessage(response: Response): Promise<string> {
  try {
    const payload = (await response.json()) as unknown;
    if (isRecord(payload) && typeof payload["message"] === "string") {
      return payload["message"];
    }
    if (
      isRecord(payload) &&
      isRecord(payload["error"]) &&
      typeof payload["error"]["message"] === "string"
    ) {
      return payload["error"]["message"];
    }
  } catch {}
  return `Workbook scenario request failed with status ${response.status}`;
}

export async function createWorkbookScenarioRequest(
  input: WorkbookScenarioCreateInput,
): Promise<WorkbookScenarioResponse> {
  const payload = workbookScenarioCreateRequestSchema.parse({
    name: input.name,
    ...(input.sheetName ? { sheetName: input.sheetName } : {}),
    ...(input.address ? { address: input.address } : {}),
    ...(input.viewport ? { viewport: input.viewport } : {}),
  });
  const response = await fetch(`/v2/documents/${encodeURIComponent(input.documentId)}/scenarios`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    throw new Error(await readErrorMessage(response));
  }
  return workbookScenarioResponseSchema.parse(await response.json());
}

export async function deleteWorkbookScenarioRequest(input: {
  readonly documentId: string;
  readonly scenarioDocumentId: string;
}): Promise<void> {
  const response = await fetch(
    `/v2/documents/${encodeURIComponent(input.documentId)}/scenarios/${encodeURIComponent(input.scenarioDocumentId)}`,
    {
      method: "DELETE",
    },
  );
  if (!response.ok) {
    throw new Error(await readErrorMessage(response));
  }
}
