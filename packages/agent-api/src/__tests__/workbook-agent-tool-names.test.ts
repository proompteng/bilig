import { describe, expect, it } from "vitest";
import {
  WORKBOOK_AGENT_TOOL_NAMES,
  isWorkbookAgentToolName,
  normalizeWorkbookAgentToolName,
} from "../workbook-agent-tool-names.js";

describe("workbook agent tool names", () => {
  it("uses clean unprefixed canonical tool names", () => {
    expect(WORKBOOK_AGENT_TOOL_NAMES.readWorkbook).toBe("read_workbook");
    expect(WORKBOOK_AGENT_TOOL_NAMES.startWorkflow).toBe("start_workflow");
    expect(WORKBOOK_AGENT_TOOL_NAMES.createSheet).toBe("create_sheet");
    expect(WORKBOOK_AGENT_TOOL_NAMES.insertRows).toBe("insert_rows");
    expect(WORKBOOK_AGENT_TOOL_NAMES.deleteColumns).toBe("delete_columns");
  });

  it("normalizes legacy prefixed tool names for compatibility", () => {
    expect(normalizeWorkbookAgentToolName("bilig_read_workbook")).toBe(
      WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
    );
    expect(normalizeWorkbookAgentToolName("bilig.read_workbook")).toBe(
      WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
    );
    expect(normalizeWorkbookAgentToolName("bilig_start_workflow")).toBe(
      WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
    );
    expect(normalizeWorkbookAgentToolName("bilig_insert_rows")).toBe(
      WORKBOOK_AGENT_TOOL_NAMES.insertRows,
    );
    expect(normalizeWorkbookAgentToolName("bilig.delete_columns")).toBe(
      WORKBOOK_AGENT_TOOL_NAMES.deleteColumns,
    );
    expect(isWorkbookAgentToolName("bilig_search_workbook")).toBe(true);
    expect(isWorkbookAgentToolName("insert_rows")).toBe(true);
    expect(isWorkbookAgentToolName("search_workbook")).toBe(true);
  });
});
