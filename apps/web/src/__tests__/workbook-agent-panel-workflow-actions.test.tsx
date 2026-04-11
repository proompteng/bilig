// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WorkflowActionStrip } from "../workbook-agent-panel-workflow-actions.js";

describe("WorkflowActionStrip", () => {
  beforeEach(() => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
  });

  afterEach(() => {
    vi.restoreAllMocks();
    document.body.innerHTML = "";
  });

  it("starts create-sheet workflows with the provided sheet name", async () => {
    const onStartNamedWorkflow = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkflowActionStrip
          disabled={false}
          isStartingWorkflow={false}
          onStartNamedWorkflow={onStartNamedWorkflow}
          onStartSearchWorkflow={vi.fn()}
          onStartStructuralWorkflow={vi.fn()}
          onStartWorkflow={vi.fn()}
        />,
      );
    });

    const input = host.querySelector("[data-testid='workbook-agent-structural-sheet-name-input']");
    const button = host.querySelector("[data-testid='workbook-agent-workflow-start-createSheet']");
    expect(input instanceof HTMLInputElement).toBe(true);
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(input instanceof HTMLInputElement) || !(button instanceof HTMLButtonElement)) {
        throw new Error("Structural workflow controls not found");
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, "set") : null;
      if (typeof valueSetter !== "function") {
        throw new Error("Input value setter not found");
      }
      Reflect.apply(valueSetter, input, ["Forecast"]);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      button.click();
    });

    expect(onStartNamedWorkflow).toHaveBeenCalledWith("createSheet", "Forecast");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts hide-current-row workflows without extra input", async () => {
    const onStartStructuralWorkflow = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkflowActionStrip
          disabled={false}
          isStartingWorkflow={false}
          onStartNamedWorkflow={vi.fn()}
          onStartSearchWorkflow={vi.fn()}
          onStartStructuralWorkflow={onStartStructuralWorkflow}
          onStartWorkflow={vi.fn()}
        />,
      );
    });

    const button = host.querySelector("[data-testid='workbook-agent-workflow-start-hideCurrentRow']");
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Structural workflow button not found");
      }
      button.click();
    });

    expect(onStartStructuralWorkflow).toHaveBeenCalledWith("hideCurrentRow");

    await act(async () => {
      root.unmount();
    });
  });

  it("starts highlight-formula workflows from the quick action list", async () => {
    const onStartWorkflow = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkflowActionStrip
          disabled={false}
          isStartingWorkflow={false}
          onStartNamedWorkflow={vi.fn()}
          onStartSearchWorkflow={vi.fn()}
          onStartStructuralWorkflow={vi.fn()}
          onStartWorkflow={onStartWorkflow}
        />,
      );
    });

    const button = host.querySelector(
      "[data-testid='workbook-agent-workflow-start-highlightFormulaIssues']",
    );
    expect(button instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error("Highlight formula workflow button not found");
      }
      button.click();
    });

    expect(onStartWorkflow).toHaveBeenCalledWith("highlightFormulaIssues");

    await act(async () => {
      root.unmount();
    });
  });
});
