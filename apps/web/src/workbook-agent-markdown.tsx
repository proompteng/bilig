import { Fragment, type ReactNode } from "react";
import { marked, type Token, type Tokens, type TokensList } from "marked";
import { cn } from "./cn.js";

type WorkbookMarkdownTone = "default" | "muted";

function normalizeMarkdownHeading(text: string): string {
  return text.trim().replaceAll(/\s+/g, " ").toLowerCase();
}

function stripLeadingMarkdownTitle(markdown: string, title: string | null | undefined): string {
  if (!title) {
    return markdown;
  }
  const lines = markdown.trimStart().split("\n");
  const firstLine = lines[0]?.match(/^#{1,6}\s+(.+)$/);
  if (
    !firstLine ||
    normalizeMarkdownHeading(firstLine[1] ?? "") !== normalizeMarkdownHeading(title)
  ) {
    return markdown;
  }
  return lines.slice(1).join("\n").trimStart();
}

function sanitizeHref(href: string | undefined): string | null {
  if (!href) {
    return null;
  }
  const normalized = href.trim();
  if (normalized.length === 0) {
    return null;
  }
  if (
    normalized.startsWith("#") ||
    normalized.startsWith("/") ||
    normalized.startsWith("http://") ||
    normalized.startsWith("https://") ||
    normalized.startsWith("mailto:")
  ) {
    return normalized;
  }
  return null;
}

function tokenText(token: Token): string {
  return "text" in token && typeof token.text === "string" ? token.text : token.raw;
}

function renderInlineTokens(tokens: readonly Token[] | undefined, keyPrefix: string): ReactNode[] {
  if (!tokens || tokens.length === 0) {
    return [];
  }
  return tokens.map((token, index) => {
    const key = `${keyPrefix}-${String(index)}`;
    switch (token.type) {
      case "strong":
        return (
          <strong key={key} className="font-semibold text-[var(--wb-text)]">
            {renderInlineTokens(token.tokens, key)}
          </strong>
        );
      case "em":
        return (
          <em key={key} className="italic">
            {renderInlineTokens(token.tokens, key)}
          </em>
        );
      case "codespan":
        return (
          <code
            key={key}
            className="rounded-[calc(var(--wb-radius-control)-4px)] bg-[var(--wb-surface-subtle)] px-1.5 py-0.5 font-mono text-[0.95em] text-[var(--wb-text)]"
          >
            {token.text ?? token.raw ?? ""}
          </code>
        );
      case "br":
        return <br key={key} />;
      case "del":
        return (
          <del key={key} className="line-through">
            {renderInlineTokens(token.tokens, key)}
          </del>
        );
      case "link": {
        const href = sanitizeHref(token.href);
        const content = renderInlineTokens(token.tokens, key);
        if (!href) {
          return <Fragment key={key}>{content}</Fragment>;
        }
        const external = href.startsWith("http://") || href.startsWith("https://");
        return (
          <a
            key={key}
            className="font-medium text-[var(--wb-accent)] underline decoration-[var(--wb-border-strong)] underline-offset-2 hover:brightness-[0.95]"
            href={href}
            rel={external ? "noreferrer" : undefined}
            target={external ? "_blank" : undefined}
            title={token.title ?? undefined}
          >
            {content}
          </a>
        );
      }
      case "text":
        return token.tokens && token.tokens.length > 0 ? (
          <Fragment key={key}>{renderInlineTokens(token.tokens, key)}</Fragment>
        ) : (
          <Fragment key={key}>{token.text ?? token.raw ?? ""}</Fragment>
        );
      case "escape":
      case "html":
      default:
        return <Fragment key={key}>{tokenText(token)}</Fragment>;
    }
  });
}

function renderBlockTokens(
  tokens: readonly Token[] | undefined,
  keyPrefix: string,
  tone: WorkbookMarkdownTone,
): ReactNode[] {
  if (!tokens || tokens.length === 0) {
    return [];
  }
  const bodyClass = tone === "muted" ? "text-[var(--wb-text-muted)]" : "text-[var(--wb-text)]";
  return tokens.flatMap((token, index) => {
    const key = `${keyPrefix}-${String(index)}`;
    switch (token.type) {
      case "space":
        return [];
      case "heading": {
        const depth = token.depth ?? 1;
        const className =
          depth <= 2
            ? `text-[13px] font-semibold ${bodyClass}`
            : `text-[12px] font-semibold ${bodyClass}`;
        return (
          <div key={key} className={className}>
            {renderInlineTokens(token.tokens, key)}
          </div>
        );
      }
      case "paragraph":
        return (
          <p key={key} className={`min-w-0 break-words text-[13px] leading-6 ${bodyClass}`}>
            {renderInlineTokens(token.tokens, key)}
          </p>
        );
      case "text":
        return (
          <p key={key} className={`min-w-0 break-words text-[13px] leading-6 ${bodyClass}`}>
            {renderInlineTokens(token.tokens, key)}
          </p>
        );
      case "blockquote":
        return (
          <blockquote
            key={key}
            className={cn(
              "border-l-2 border-[var(--wb-border-strong)] pl-3",
              tone === "muted" ? "text-[var(--wb-text-muted)]" : "text-[var(--wb-text-subtle)]",
            )}
          >
            <div className="flex flex-col gap-2">{renderBlockTokens(token.tokens, key, tone)}</div>
          </blockquote>
        );
      case "list": {
        const ListTag = token.ordered ? "ol" : "ul";
        return (
          <ListTag
            key={key}
            className={cn(
              "ml-5 flex flex-col gap-2 text-[13px] leading-6",
              bodyClass,
              token.ordered ? "list-decimal" : "list-disc",
            )}
            start={token.ordered && typeof token.start === "number" ? token.start : undefined}
          >
            {token.items?.map((item: Tokens.ListItem) => (
              <li key={`${key}-item-${item.raw}`} className="min-w-0 break-words pl-1">
                <div className="flex flex-col gap-2">
                  {renderBlockTokens(item.tokens, `${key}-item-${item.raw}`, tone)}
                </div>
              </li>
            ))}
          </ListTag>
        );
      }
      case "code":
        return (
          <pre
            key={key}
            className="overflow-x-auto rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-3 py-2 font-mono text-[12px] leading-6 text-[var(--wb-text)]"
          >
            <code>{token.text ?? ""}</code>
          </pre>
        );
      case "hr":
        return <div key={key} className="border-t border-[var(--wb-border)]" />;
      case "html":
      default:
        return (
          <p key={key} className={`min-w-0 break-words text-[13px] leading-6 ${bodyClass}`}>
            {tokenText(token)}
          </p>
        );
    }
  });
}

export function WorkbookAgentMarkdown(props: {
  readonly markdown: string;
  readonly className?: string;
  readonly tone?: WorkbookMarkdownTone;
  readonly title?: string | null;
}) {
  const normalized = stripLeadingMarkdownTitle(props.markdown, props.title).trim();
  if (normalized.length === 0) {
    return null;
  }
  const tone = props.tone ?? "default";
  const tokens: TokensList = marked.lexer(normalized, { gfm: true, breaks: true });
  return (
    <div className={cn("flex flex-col gap-2", props.className)}>
      {renderBlockTokens(tokens, "markdown", tone)}
    </div>
  );
}
