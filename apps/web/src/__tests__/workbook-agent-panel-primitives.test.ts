import { describe, expect, it } from "vitest";
import {
  agentPanelComposerSendButtonClass,
  agentPanelDisclosureChevronClass,
  agentPanelDisclosureContentClass,
  agentPanelDisclosureSummaryClass,
  agentPanelDisclosureTriggerClass,
} from "../workbook-agent-panel-primitives.js";

describe("workbook agent panel primitives", () => {
  it("keeps the composer send button circular", () => {
    const className = agentPanelComposerSendButtonClass();

    expect(className).toContain("h-8");
    expect(className).toContain("w-8");
    expect(className).toContain("rounded-full");
  });

  it("keeps disclosure rows tightly aligned around the chevron", () => {
    expect(agentPanelDisclosureTriggerClass()).toContain("flex");
    expect(agentPanelDisclosureTriggerClass()).not.toContain("grid");
    expect(agentPanelDisclosureTriggerClass()).toContain("items-start");
    expect(agentPanelDisclosureContentClass()).toContain("flex-1");
    expect(agentPanelDisclosureContentClass()).toContain("flex-wrap");
    expect(agentPanelDisclosureSummaryClass()).toContain("whitespace-normal");
    expect(agentPanelDisclosureSummaryClass()).not.toContain("truncate");
    expect(agentPanelDisclosureChevronClass({ open: false })).toContain("size-4");
    expect(agentPanelDisclosureChevronClass({ open: false })).toContain("mt-0.5");
  });
});
