import { ErrorCode, type CellValue } from "@bilig/protocol";
import type { TextBuiltin } from "./text.js";

interface TextCoreBuiltinDeps {
  error: (code: ErrorCode) => CellValue;
  stringResult: (value: string) => CellValue;
  booleanResult: (value: boolean) => CellValue;
  firstError: (args: readonly (CellValue | undefined)[]) => CellValue | undefined;
  coerceText: (value: CellValue) => string;
  coerceNumber: (value: CellValue) => number | undefined;
}

const bahtDigitWords = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];
const bahtScaleWords = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"];
const maxBahtTextSatang = 999_999_999_999_999;

const halfWidthKanaToFullWidthMap = new Map<string, string>([
  ["｡", "。"],
  ["｢", "「"],
  ["｣", "」"],
  ["､", "、"],
  ["･", "・"],
  ["ｦ", "ヲ"],
  ["ｧ", "ァ"],
  ["ｨ", "ィ"],
  ["ｩ", "ゥ"],
  ["ｪ", "ェ"],
  ["ｫ", "ォ"],
  ["ｬ", "ャ"],
  ["ｭ", "ュ"],
  ["ｮ", "ョ"],
  ["ｯ", "ッ"],
  ["ｰ", "ー"],
  ["ｱ", "ア"],
  ["ｲ", "イ"],
  ["ｳ", "ウ"],
  ["ｴ", "エ"],
  ["ｵ", "オ"],
  ["ｶ", "カ"],
  ["ｷ", "キ"],
  ["ｸ", "ク"],
  ["ｹ", "ケ"],
  ["ｺ", "コ"],
  ["ｻ", "サ"],
  ["ｼ", "シ"],
  ["ｽ", "ス"],
  ["ｾ", "セ"],
  ["ｿ", "ソ"],
  ["ﾀ", "タ"],
  ["ﾁ", "チ"],
  ["ﾂ", "ツ"],
  ["ﾃ", "テ"],
  ["ﾄ", "ト"],
  ["ﾅ", "ナ"],
  ["ﾆ", "ニ"],
  ["ﾇ", "ヌ"],
  ["ﾈ", "ネ"],
  ["ﾉ", "ノ"],
  ["ﾊ", "ハ"],
  ["ﾋ", "ヒ"],
  ["ﾌ", "フ"],
  ["ﾍ", "ヘ"],
  ["ﾎ", "ホ"],
  ["ﾏ", "マ"],
  ["ﾐ", "ミ"],
  ["ﾑ", "ム"],
  ["ﾒ", "メ"],
  ["ﾓ", "モ"],
  ["ﾔ", "ヤ"],
  ["ﾕ", "ユ"],
  ["ﾖ", "ヨ"],
  ["ﾗ", "ラ"],
  ["ﾘ", "リ"],
  ["ﾙ", "ル"],
  ["ﾚ", "レ"],
  ["ﾛ", "ロ"],
  ["ﾜ", "ワ"],
  ["ﾝ", "ン"],
]);

const halfWidthVoicedKanaToFullWidthMap = new Map<string, string>([
  ["ｳﾞ", "ヴ"],
  ["ｶﾞ", "ガ"],
  ["ｷﾞ", "ギ"],
  ["ｸﾞ", "グ"],
  ["ｹﾞ", "ゲ"],
  ["ｺﾞ", "ゴ"],
  ["ｻﾞ", "ザ"],
  ["ｼﾞ", "ジ"],
  ["ｽﾞ", "ズ"],
  ["ｾﾞ", "ゼ"],
  ["ｿﾞ", "ゾ"],
  ["ﾀﾞ", "ダ"],
  ["ﾁﾞ", "ヂ"],
  ["ﾂﾞ", "ヅ"],
  ["ﾃﾞ", "デ"],
  ["ﾄﾞ", "ド"],
  ["ﾊﾞ", "バ"],
  ["ﾋﾞ", "ビ"],
  ["ﾌﾞ", "ブ"],
  ["ﾍﾞ", "ベ"],
  ["ﾎﾞ", "ボ"],
  ["ﾊﾟ", "パ"],
  ["ﾋﾟ", "ピ"],
  ["ﾌﾟ", "プ"],
  ["ﾍﾟ", "ペ"],
  ["ﾎﾟ", "ポ"],
]);

const fullWidthKanaToHalfWidthMap = new Map<string, string>([
  ...[...halfWidthKanaToFullWidthMap.entries()].map(([half, full]) => [full, half] as const),
  ...[...halfWidthVoicedKanaToFullWidthMap.entries()].map(([half, full]) => [full, half] as const),
]);

function bahtSegmentText(digits: string): string {
  const normalized = digits.replace(/^0+(?=\d)/u, "");
  if (normalized === "" || /^0+$/u.test(normalized)) {
    return "";
  }

  let output = "";
  let hasHigherNonZero = false;
  const length = normalized.length;
  for (let index = 0; index < length; index += 1) {
    const digit = normalized.charCodeAt(index) - 48;
    if (digit === 0) {
      continue;
    }

    const position = length - index - 1;
    if (position === 0) {
      output += digit === 1 && hasHigherNonZero ? "เอ็ด" : bahtDigitWords[digit]!;
    } else if (position === 1) {
      output += digit === 1 ? "สิบ" : digit === 2 ? "ยี่สิบ" : `${bahtDigitWords[digit]!}สิบ`;
    } else {
      output += `${bahtDigitWords[digit]!}${bahtScaleWords[position]!}`;
    }
    hasHigherNonZero = true;
  }
  return output;
}

function bahtIntegerText(digits: string): string {
  const normalized = digits.replace(/^0+(?=\d)/u, "");
  if (normalized === "" || /^0+$/u.test(normalized)) {
    return bahtDigitWords[0]!;
  }
  if (normalized.length > 6) {
    const head = bahtIntegerText(normalized.slice(0, -6));
    const tail = bahtSegmentText(normalized.slice(-6));
    return `${head}ล้าน${tail}`;
  }
  return bahtSegmentText(normalized) || bahtDigitWords[0]!;
}

function bahtTextFromNumber(value: number, deps: TextCoreBuiltinDeps): CellValue {
  if (!Number.isFinite(value)) {
    return deps.error(ErrorCode.Value);
  }

  const absolute = Math.abs(value);
  const scaled = Math.round(absolute * 100);
  if (!Number.isSafeInteger(scaled) || scaled > maxBahtTextSatang) {
    return deps.error(ErrorCode.Value);
  }

  const baht = Math.trunc(scaled / 100);
  const satang = scaled % 100;
  const prefix = value < 0 ? "ลบ" : "";
  const bahtText = bahtIntegerText(String(baht));
  if (satang === 0) {
    return deps.stringResult(`${prefix}${bahtText}บาทถ้วน`);
  }
  return deps.stringResult(`${prefix}${bahtText}บาท${bahtSegmentText(String(satang))}สตางค์`);
}

function excelTrim(input: string): string {
  let start = 0;
  let end = input.length;

  while (start < end && input.charCodeAt(start) === 32) {
    start += 1;
  }
  while (end > start && input.charCodeAt(end - 1) === 32) {
    end -= 1;
  }

  return input.slice(start, end).replace(/ {2,}/g, " ");
}

function stripControlCharacters(input: string): string {
  let output = "";
  for (let index = 0; index < input.length; index += 1) {
    const char = input.charCodeAt(index);
    if ((char >= 0 && char <= 31) || char === 127) {
      continue;
    }
    output += input[index] ?? "";
  }
  return output;
}

function isLeadingSurrogate(codeUnit: number): boolean {
  return codeUnit >= 0xd800 && codeUnit <= 0xdbff;
}

function isTrailingSurrogate(codeUnit: number): boolean {
  return codeUnit >= 0xdc00 && codeUnit <= 0xdfff;
}

function toJapaneseFullWidth(input: string): string {
  let output = "";
  for (let index = 0; index < input.length; index += 1) {
    const char = input[index]!;
    const code = input.charCodeAt(index);

    if (isLeadingSurrogate(code) && index + 1 < input.length) {
      const nextCode = input.charCodeAt(index + 1);
      if (isTrailingSurrogate(nextCode)) {
        output += input.slice(index, index + 2);
        index += 1;
        continue;
      }
    }

    if (code === 0x20) {
      output += "\u3000";
      continue;
    }
    if (code >= 0x21 && code <= 0x7e) {
      output += String.fromCharCode(code + 0xfee0);
      continue;
    }

    const next = input[index + 1];
    if (next !== undefined) {
      const voiced = halfWidthVoicedKanaToFullWidthMap.get(char + next);
      if (voiced !== undefined) {
        output += voiced;
        index += 1;
        continue;
      }
    }

    const mapped = halfWidthKanaToFullWidthMap.get(char);
    output += mapped ?? char;
  }
  return output;
}

function toJapaneseHalfWidth(input: string): string {
  let output = "";
  for (let index = 0; index < input.length; index += 1) {
    const char = input[index]!;
    const code = input.charCodeAt(index);

    if (isLeadingSurrogate(code) && index + 1 < input.length) {
      const nextCode = input.charCodeAt(index + 1);
      if (isTrailingSurrogate(nextCode)) {
        output += input.slice(index, index + 2);
        index += 1;
        continue;
      }
    }

    if (code === 0x3000) {
      output += " ";
      continue;
    }
    if (code >= 0xff01 && code <= 0xff5e) {
      output += String.fromCharCode(code - 0xfee0);
      continue;
    }

    const mapped = fullWidthKanaToHalfWidthMap.get(char);
    output += mapped ?? char;
  }
  return output;
}

function toTitleCase(input: string): string {
  let result = "";
  let capitalizeNext = true;

  for (let index = 0; index < input.length; index += 1) {
    const char = input[index] ?? "";
    const code = char.charCodeAt(0);
    const isAlpha = (code >= 65 && code <= 90) || (code >= 97 && code <= 122);

    if (!isAlpha) {
      capitalizeNext = true;
      result += char;
      continue;
    }

    result += capitalizeNext ? char.toUpperCase() : char.toLowerCase();
    capitalizeNext = false;
  }

  return result;
}

export function createTextCoreBuiltins(deps: TextCoreBuiltinDeps): Record<string, TextBuiltin> {
  return {
    CLEAN: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [textValue] = args;
      if (textValue === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(stripControlCharacters(deps.coerceText(textValue)));
    },
    ASC: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [textValue] = args;
      if (textValue === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(toJapaneseHalfWidth(deps.coerceText(textValue)));
    },
    JIS: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [textValue] = args;
      if (textValue === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(toJapaneseFullWidth(deps.coerceText(textValue)));
    },
    DBCS: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [textValue] = args;
      if (textValue === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(toJapaneseFullWidth(deps.coerceText(textValue)));
    },
    BAHTTEXT: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [value] = args;
      if (value === undefined) {
        return deps.error(ErrorCode.Value);
      }
      const numeric = deps.coerceNumber(value);
      return numeric === undefined ? deps.error(ErrorCode.Value) : bahtTextFromNumber(numeric, deps);
    },
    PHONETIC: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [value] = args;
      if (value === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(deps.coerceText(value));
    },
    CONCATENATE: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      if (args.length === 0) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(args.map(deps.coerceText).join(""));
    },
    CONCAT: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      if (args.length === 0) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(args.map(deps.coerceText).join(""));
    },
    PROPER: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [textValue] = args;
      if (textValue === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(toTitleCase(deps.coerceText(textValue)));
    },
    EXACT: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [leftValue, rightValue] = args;
      if (leftValue === undefined || rightValue === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.booleanResult(deps.coerceText(leftValue) === deps.coerceText(rightValue));
    },
    TRIM: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [value] = args;
      if (value === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(excelTrim(deps.coerceText(value)));
    },
    UPPER: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [value] = args;
      if (value === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(deps.coerceText(value).toUpperCase());
    },
    LOWER: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [value] = args;
      if (value === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return deps.stringResult(deps.coerceText(value).toLowerCase());
    },
  };
}
