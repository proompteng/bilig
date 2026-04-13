// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { AssistantProgressRow } from "../workbook-agent-panel-history.js";

afterEach(() => {
  document.body.innerHTML = "";
});

describe("workbook agent panel history", () => {
  it("renders the progress row with assistant body typography", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<AssistantProgressRow />);
    });

    const progressRow = host.querySelector("[data-testid='workbook-agent-progress-row']");
    expect(progressRow?.textContent).toBe("Thinking");
    expect(progressRow?.textContent).not.toBe("T h i n k i n g");

    const textContainer = progressRow?.firstElementChild;
    expect(textContainer?.className).toContain("text-[13px]");
    expect(textContainer?.className).toContain("leading-[1.65]");
    expect(textContainer?.className).toContain("text-[var(--wb-text-subtle)]");

    await act(async () => {
      root.unmount();
    });
  });
});
