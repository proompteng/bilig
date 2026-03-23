// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it, vi } from "vitest";
import { CellEditorOverlay } from "@bilig/grid";

describe("grid overlay editing", () => {
  it("commits directional moves and cancels with escape", async () => {
    (
      globalThis as typeof globalThis & {
        IS_REACT_ACT_ENVIRONMENT?: boolean;
      }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const onChange = vi.fn();
    const onCommit = vi.fn();
    const onCancel = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CellEditorOverlay
          key="first"
          label="Sheet1!B1"
          onCancel={onCancel}
          onChange={onChange}
          onCommit={onCommit}
          resolvedValue="20"
          style={{ left: 40, top: 60, width: 180 }}
          value="=A1*2"
        />,
      );
    });

    const input = host.querySelector<HTMLInputElement>("[aria-label='Sheet1!B1 editor']");
    expect(input).not.toBeNull();

    await act(async () => {
      input?.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, key: "Tab" }));
    });

    expect(onCommit).toHaveBeenCalledWith([1, 0]);

    await act(async () => {
      root.render(
        <CellEditorOverlay
          key="second"
          label="Sheet1!B1"
          onCancel={onCancel}
          onChange={onChange}
          onCommit={onCommit}
          resolvedValue="42"
          style={{ left: 40, top: 60, width: 180 }}
          value="42"
        />,
      );
    });

    const nextInput = host.querySelector<HTMLInputElement>("[aria-label='Sheet1!B1 editor']");
    expect(nextInput).not.toBeNull();

    await act(async () => {
      nextInput?.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, key: "Escape" }));
    });

    expect(onCancel).toHaveBeenCalledTimes(1);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
