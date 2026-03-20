export type TokenKind =
  | "number"
  | "identifier"
  | "quotedIdentifier"
  | "string"
  | "lparen"
  | "rparen"
  | "comma"
  | "colon"
  | "bang"
  | "plus"
  | "minus"
  | "star"
  | "slash"
  | "caret"
  | "percent"
  | "ampersand"
  | "eq"
  | "neq"
  | "gt"
  | "gte"
  | "lt"
  | "lte"
  | "eof";

export interface Token {
  kind: TokenKind;
  value: string;
}

export function lexFormula(input: string): Token[] {
  const source = input.trim();
  const tokens: Token[] = [];
  let index = 0;

  while (index < source.length) {
    const char = source[index]!;

    if (/\s/.test(char)) {
      index += 1;
      continue;
    }

    if (/[0-9.]/.test(char)) {
      let end = index + 1;
      while (end < source.length && /[0-9.]/.test(source[end]!)) end += 1;
      tokens.push({ kind: "number", value: source.slice(index, end) });
      index = end;
      continue;
    }

    if (char === "\"") {
      let end = index + 1;
      let value = "";
      while (end < source.length) {
        const current = source[end]!;
        const next = source[end + 1];
        if (current === "\"" && next === "\"") {
          value += "\"";
          end += 2;
          continue;
        }
        if (current === "\"") {
          break;
        }
        value += current;
        end += 1;
      }
      tokens.push({ kind: "string", value });
      index = end + 1;
      continue;
    }

    if (char === "'") {
      let end = index + 1;
      let value = "";
      while (end < source.length) {
        const current = source[end]!;
        const next = source[end + 1];
        if (current === "'" && next === "'") {
          value += "'";
          end += 2;
          continue;
        }
        if (current === "'") {
          break;
        }
        value += current;
        end += 1;
      }
      tokens.push({ kind: "quotedIdentifier", value });
      index = end + 1;
      continue;
    }

    if (/[A-Za-z_$]/.test(char)) {
      let end = index + 1;
      while (end < source.length && /[A-Za-z0-9_.$]/.test(source[end]!)) end += 1;
      tokens.push({ kind: "identifier", value: source.slice(index, end) });
      index = end;
      continue;
    }

    const two = source.slice(index, index + 2);
    if (two === ">=") {
      tokens.push({ kind: "gte", value: two });
      index += 2;
      continue;
    }
    if (two === "<=") {
      tokens.push({ kind: "lte", value: two });
      index += 2;
      continue;
    }
    if (two === "<>") {
      tokens.push({ kind: "neq", value: two });
      index += 2;
      continue;
    }

    const singleMap: Record<string, TokenKind> = {
      "(": "lparen",
      ")": "rparen",
      ",": "comma",
      ":": "colon",
      "!": "bang",
      "+": "plus",
      "-": "minus",
      "*": "star",
      "/": "slash",
      "^": "caret",
      "%": "percent",
      "&": "ampersand",
      "=": "eq",
      ">": "gt",
      "<": "lt"
    };

    const kind = singleMap[char];
    if (!kind) {
      throw new Error(`Unexpected token '${char}' in formula '${input}'`);
    }

    tokens.push({ kind, value: char });
    index += 1;
  }

  tokens.push({ kind: "eof", value: "" });
  return tokens;
}
