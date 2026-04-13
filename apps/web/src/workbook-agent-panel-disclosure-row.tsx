import { type ReactNode, useState } from "react";
import { Collapsible } from "@base-ui/react/collapsible";
import { ScrollArea } from "@base-ui/react/scroll-area";
import { ChevronRight } from "lucide-react";
import { cn } from "./cn.js";
import {
  agentPanelDisclosureBadgeClass,
  agentPanelDisclosureBodyClass,
  agentPanelDisclosureBodyCardClass,
  agentPanelDisclosureChevronClass,
  agentPanelDisclosureContentClass,
  agentPanelDisclosureFrameClass,
  agentPanelDisclosureLabelClass,
  agentPanelDisclosurePanelClass,
  agentPanelDisclosureSummaryClass,
  agentPanelDisclosureTriggerClass,
  agentPanelDisclosureViewportClass,
  agentPanelScrollAreaScrollbarClass,
  agentPanelScrollAreaThumbClass,
} from "./workbook-agent-panel-primitives.js";

export function WorkbookAgentDisclosureRow(props: {
  readonly id: string;
  readonly label: string;
  readonly summary?: string | null;
  readonly badge?: ReactNode;
  readonly triggerLabel: {
    readonly expanded: string;
    readonly collapsed: string;
  };
  readonly triggerTestId: string;
  readonly panelTestId?: string;
  readonly labelClassName?: string;
  readonly summaryClassName?: string;
  readonly children: ReactNode;
}) {
  const [open, setOpen] = useState(false);
  const summary = props.summary?.trim() ? props.summary.trim() : null;
  const viewportTestId = props.panelTestId ? `${props.panelTestId}-viewport` : undefined;

  return (
    <Collapsible.Root
      open={open}
      onOpenChange={(nextOpen) => {
        setOpen(nextOpen);
      }}
    >
      <div className={agentPanelDisclosureFrameClass({ open })}>
        <Collapsible.Trigger
          aria-label={open ? props.triggerLabel.expanded : props.triggerLabel.collapsed}
          className={agentPanelDisclosureTriggerClass()}
          data-testid={props.triggerTestId}
          type="button"
        >
          <ChevronRight className={agentPanelDisclosureChevronClass({ open })} />
          <div className={agentPanelDisclosureContentClass({ open })}>
            <span className={cn(agentPanelDisclosureLabelClass(), props.labelClassName)}>
              {props.label}
            </span>
            {summary ? (
              <span
                className={cn(agentPanelDisclosureSummaryClass({ open }), props.summaryClassName)}
              >
                {summary}
              </span>
            ) : null}
          </div>
          {props.badge ? (
            <div className={agentPanelDisclosureBadgeClass()}>{props.badge}</div>
          ) : null}
        </Collapsible.Trigger>
        <Collapsible.Panel
          className={agentPanelDisclosurePanelClass()}
          data-testid={props.panelTestId}
        >
          <ScrollArea.Root className="relative overflow-hidden">
            <ScrollArea.Viewport
              className={agentPanelDisclosureViewportClass()}
              data-testid={viewportTestId}
            >
              <ScrollArea.Content className={agentPanelDisclosureBodyClass()}>
                <div className={agentPanelDisclosureBodyCardClass()}>{props.children}</div>
              </ScrollArea.Content>
            </ScrollArea.Viewport>
            <ScrollArea.Scrollbar
              className={agentPanelScrollAreaScrollbarClass()}
              keepMounted
              orientation="vertical"
            >
              <ScrollArea.Thumb className={agentPanelScrollAreaThumbClass()} />
            </ScrollArea.Scrollbar>
          </ScrollArea.Root>
        </Collapsible.Panel>
      </div>
    </Collapsible.Root>
  );
}
