import { describe, expect, it } from "vitest";
import {
  agentPanelComposerSendButtonClass,
  agentPanelDisclosureChevronClass,
  agentPanelDisclosureContentClass,
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
    expect(agentPanelDisclosureContentClass()).toContain("flex-1");
    expect(agentPanelDisclosureContentClass()).toContain("gap-1.5");
    expect(agentPanelDisclosureChevronClass({ open: false })).toContain("size-4");
  });
});
