import { ValueTag } from "./protocol";
import { EXCEL_SECONDS_PER_DAY } from "./date-finance";
import { scalarText, trimAsciiWhitespace } from "./text-codec";

function toNumberExactValue(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

function parseAsciiIntegerSegment(input: string, startIndex: i32, endIndex: i32): i32 {
  let start = startIndex;
  let end = endIndex;
  while (start < end && input.charCodeAt(start) <= 32) {
    start += 1;
  }
  while (end > start && input.charCodeAt(end - 1) <= 32) {
    end -= 1;
  }
  if (start >= end) {
    return i32.MIN_VALUE;
  }

  let value = 0;
  for (let index = start; index < end; index += 1) {
    const char = input.charCodeAt(index);
    if (char < 48 || char > 57) {
      return i32.MIN_VALUE;
    }
    value = value * 10 + (char - 48);
  }
  return value;
}

export function parseTimeValueText(input: string): f64 {
  const text = trimAsciiWhitespace(input);
  if (text.length == 0) {
    return NaN;
  }

  let hasMeridiem = false;
  let hasPm = false;
  let coreEnd = text.length;
  if (coreEnd >= 2) {
    const secondLast = text.charCodeAt(coreEnd - 2);
    const last = text.charCodeAt(coreEnd - 1);
    const isAm = (secondLast == 65 || secondLast == 97) && (last == 77 || last == 109);
    const isPm = (secondLast == 80 || secondLast == 112) && (last == 77 || last == 109);
    if (isAm || isPm) {
      let whitespaceStart = coreEnd - 2;
      while (whitespaceStart > 0 && text.charCodeAt(whitespaceStart - 1) <= 32) {
        whitespaceStart -= 1;
      }
      if (whitespaceStart == coreEnd - 2) {
        return NaN;
      }
      hasMeridiem = true;
      hasPm = isPm;
      coreEnd = whitespaceStart;
    }
  }

  let firstColon = -1;
  let secondColon = -1;
  for (let index = 0; index < coreEnd; index += 1) {
    if (text.charCodeAt(index) != 58) {
      continue;
    }
    if (firstColon < 0) {
      firstColon = index;
      continue;
    }
    if (secondColon < 0) {
      secondColon = index;
      continue;
    }
    return NaN;
  }
  if (firstColon < 0) {
    return NaN;
  }

  const hour = parseAsciiIntegerSegment(text, 0, firstColon);
  const minute = parseAsciiIntegerSegment(
    text,
    firstColon + 1,
    secondColon < 0 ? coreEnd : secondColon,
  );
  const second = secondColon < 0 ? 0 : parseAsciiIntegerSegment(text, secondColon + 1, coreEnd);
  if (hour == i32.MIN_VALUE || minute == i32.MIN_VALUE || second == i32.MIN_VALUE) {
    return NaN;
  }
  if (minute < 0 || minute > 59 || second < 0 || second > 59) {
    return NaN;
  }

  let normalizedHour = hour;
  if (hasMeridiem) {
    if (hour < 1 || hour > 12) {
      return NaN;
    }
    if (hour == 12) {
      normalizedHour = hasPm ? 12 : 0;
    } else if (hasPm) {
      normalizedHour = hour + 12;
    }
  } else if (hour == 24 && minute == 0 && second == 0) {
    normalizedHour = 0;
  } else if (hour < 0 || hour > 23) {
    return NaN;
  }

  return <f64>(normalizedHour * 3600 + minute * 60 + second) / <f64>EXCEL_SECONDS_PER_DAY;
}

export function parseNumericText(input: string): f64 {
  const text = trimAsciiWhitespace(input);
  if (text.length == 0) {
    return 0;
  }

  let index = 0;
  let sign = 1.0;
  const first = text.charCodeAt(index);
  if (first == 43) {
    index += 1;
  } else if (first == 45) {
    sign = -1.0;
    index += 1;
  }

  let value = 0.0;
  let digitCount = 0;
  while (index < text.length) {
    const char = text.charCodeAt(index);
    if (char < 48 || char > 57) {
      break;
    }
    value = value * 10.0 + <f64>(char - 48);
    digitCount += 1;
    index += 1;
  }

  if (index < text.length && text.charCodeAt(index) == 46) {
    index += 1;
    let factor = 0.1;
    while (index < text.length) {
      const char = text.charCodeAt(index);
      if (char < 48 || char > 57) {
        break;
      }
      value += <f64>(char - 48) * factor;
      factor *= 0.1;
      digitCount += 1;
      index += 1;
    }
  }

  if (digitCount == 0) {
    return NaN;
  }

  if (index < text.length) {
    const exponentMarker = text.charCodeAt(index);
    if (exponentMarker == 69 || exponentMarker == 101) {
      index += 1;
      let exponentSign = 1;
      if (index < text.length) {
        const exponentPrefix = text.charCodeAt(index);
        if (exponentPrefix == 43) {
          index += 1;
        } else if (exponentPrefix == 45) {
          exponentSign = -1;
          index += 1;
        }
      }

      let exponent = 0;
      let exponentDigits = 0;
      while (index < text.length) {
        const char = text.charCodeAt(index);
        if (char < 48 || char > 57) {
          break;
        }
        exponent = exponent * 10 + (char - 48);
        exponentDigits += 1;
        index += 1;
      }
      if (exponentDigits == 0) {
        return NaN;
      }
      value *= Math.pow(10.0, <f64>(exponentSign * exponent));
    }
  }

  if (index != text.length) {
    return NaN;
  }

  const parsed = sign * value;
  if (parsed == Infinity || parsed == -Infinity) {
    return NaN;
  }
  return parsed;
}

export function coerceScalarNumberLikeText(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): f64 {
  if (tag == ValueTag.String) {
    const text = scalarText(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    return text == null ? NaN : parseNumericText(text);
  }
  return toNumberExactValue(tag, value);
}

export function firstUnicodeCodePoint(text: string): i32 {
  if (text.length == 0) {
    return -1;
  }
  const first = text.charCodeAt(0);
  if ((first & 0xfc00) == 0xd800 && text.length > 1) {
    const next = text.charCodeAt(1);
    if ((next & 0xfc00) == 0xdc00) {
      return 0x10000 + ((first - 0xd800) << 10) + (next - 0xdc00);
    }
  }
  return first;
}

export function stringFromUnicodeCodePoint(codePoint: i32): string {
  if (codePoint <= 0xffff) {
    return String.fromCharCode(codePoint);
  }
  const adjusted = codePoint - 0x10000;
  const high = 0xd800 + (adjusted >> 10);
  const low = 0xdc00 + (adjusted & 0x3ff);
  return String.fromCharCode(high) + String.fromCharCode(low);
}

export function stripControlCharacters(text: string): string {
  let output = "";
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    if ((code >= 0 && code <= 31) || code == 127) {
      continue;
    }
    output += String.fromCharCode(code);
  }
  return output;
}

function isLeadingSurrogate(codeUnit: i32): bool {
  return codeUnit >= 0xd800 && codeUnit <= 0xdbff;
}

function isTrailingSurrogate(codeUnit: i32): bool {
  return codeUnit >= 0xdc00 && codeUnit <= 0xdfff;
}

function halfWidthKanaToFullWidthChar(code: i32): string | null {
  switch (code) {
    case 0xff61:
      return "。";
    case 0xff62:
      return "「";
    case 0xff63:
      return "」";
    case 0xff64:
      return "、";
    case 0xff65:
      return "・";
    case 0xff66:
      return "ヲ";
    case 0xff67:
      return "ァ";
    case 0xff68:
      return "ィ";
    case 0xff69:
      return "ゥ";
    case 0xff6a:
      return "ェ";
    case 0xff6b:
      return "ォ";
    case 0xff6c:
      return "ャ";
    case 0xff6d:
      return "ュ";
    case 0xff6e:
      return "ョ";
    case 0xff6f:
      return "ッ";
    case 0xff70:
      return "ー";
    case 0xff71:
      return "ア";
    case 0xff72:
      return "イ";
    case 0xff73:
      return "ウ";
    case 0xff74:
      return "エ";
    case 0xff75:
      return "オ";
    case 0xff76:
      return "カ";
    case 0xff77:
      return "キ";
    case 0xff78:
      return "ク";
    case 0xff79:
      return "ケ";
    case 0xff7a:
      return "コ";
    case 0xff7b:
      return "サ";
    case 0xff7c:
      return "シ";
    case 0xff7d:
      return "ス";
    case 0xff7e:
      return "セ";
    case 0xff7f:
      return "ソ";
    case 0xff80:
      return "タ";
    case 0xff81:
      return "チ";
    case 0xff82:
      return "ツ";
    case 0xff83:
      return "テ";
    case 0xff84:
      return "ト";
    case 0xff85:
      return "ナ";
    case 0xff86:
      return "ニ";
    case 0xff87:
      return "ヌ";
    case 0xff88:
      return "ネ";
    case 0xff89:
      return "ノ";
    case 0xff8a:
      return "ハ";
    case 0xff8b:
      return "ヒ";
    case 0xff8c:
      return "フ";
    case 0xff8d:
      return "ヘ";
    case 0xff8e:
      return "ホ";
    case 0xff8f:
      return "マ";
    case 0xff90:
      return "ミ";
    case 0xff91:
      return "ム";
    case 0xff92:
      return "メ";
    case 0xff93:
      return "モ";
    case 0xff94:
      return "ヤ";
    case 0xff95:
      return "ユ";
    case 0xff96:
      return "ヨ";
    case 0xff97:
      return "ラ";
    case 0xff98:
      return "リ";
    case 0xff99:
      return "ル";
    case 0xff9a:
      return "レ";
    case 0xff9b:
      return "ロ";
    case 0xff9c:
      return "ワ";
    case 0xff9d:
      return "ン";
    default:
      return null;
  }
}

function halfWidthVoicedKanaToFullWidthChar(baseCode: i32, markCode: i32): string | null {
  if (markCode == 0xff9e) {
    switch (baseCode) {
      case 0xff73:
        return "ヴ";
      case 0xff76:
        return "ガ";
      case 0xff77:
        return "ギ";
      case 0xff78:
        return "グ";
      case 0xff79:
        return "ゲ";
      case 0xff7a:
        return "ゴ";
      case 0xff7b:
        return "ザ";
      case 0xff7c:
        return "ジ";
      case 0xff7d:
        return "ズ";
      case 0xff7e:
        return "ゼ";
      case 0xff7f:
        return "ゾ";
      case 0xff80:
        return "ダ";
      case 0xff81:
        return "ヂ";
      case 0xff82:
        return "ヅ";
      case 0xff83:
        return "デ";
      case 0xff84:
        return "ド";
      case 0xff8a:
        return "バ";
      case 0xff8b:
        return "ビ";
      case 0xff8c:
        return "ブ";
      case 0xff8d:
        return "ベ";
      case 0xff8e:
        return "ボ";
      default:
        return null;
    }
  }
  if (markCode == 0xff9f) {
    switch (baseCode) {
      case 0xff8a:
        return "パ";
      case 0xff8b:
        return "ピ";
      case 0xff8c:
        return "プ";
      case 0xff8d:
        return "ペ";
      case 0xff8e:
        return "ポ";
      default:
        return null;
    }
  }
  return null;
}

function fullWidthKanaToHalfWidthChar(code: i32): string | null {
  switch (code) {
    case 0x3002:
      return "｡";
    case 0x300c:
      return "｢";
    case 0x300d:
      return "｣";
    case 0x3001:
      return "､";
    case 0x30fb:
      return "･";
    case 0x30f2:
      return "ｦ";
    case 0x30a1:
      return "ｧ";
    case 0x30a3:
      return "ｨ";
    case 0x30a5:
      return "ｩ";
    case 0x30a7:
      return "ｪ";
    case 0x30a9:
      return "ｫ";
    case 0x30e3:
      return "ｬ";
    case 0x30e5:
      return "ｭ";
    case 0x30e7:
      return "ｮ";
    case 0x30c3:
      return "ｯ";
    case 0x30fc:
      return "ｰ";
    case 0x30a2:
      return "ｱ";
    case 0x30a4:
      return "ｲ";
    case 0x30a6:
      return "ｳ";
    case 0x30a8:
      return "ｴ";
    case 0x30aa:
      return "ｵ";
    case 0x30ab:
      return "ｶ";
    case 0x30ad:
      return "ｷ";
    case 0x30af:
      return "ｸ";
    case 0x30b1:
      return "ｹ";
    case 0x30b3:
      return "ｺ";
    case 0x30b5:
      return "ｻ";
    case 0x30b7:
      return "ｼ";
    case 0x30b9:
      return "ｽ";
    case 0x30bb:
      return "ｾ";
    case 0x30bd:
      return "ｿ";
    case 0x30bf:
      return "ﾀ";
    case 0x30c1:
      return "ﾁ";
    case 0x30c4:
      return "ﾂ";
    case 0x30c6:
      return "ﾃ";
    case 0x30c8:
      return "ﾄ";
    case 0x30ca:
      return "ﾅ";
    case 0x30cb:
      return "ﾆ";
    case 0x30cc:
      return "ﾇ";
    case 0x30cd:
      return "ﾈ";
    case 0x30ce:
      return "ﾉ";
    case 0x30cf:
      return "ﾊ";
    case 0x30d2:
      return "ﾋ";
    case 0x30d5:
      return "ﾌ";
    case 0x30d8:
      return "ﾍ";
    case 0x30db:
      return "ﾎ";
    case 0x30de:
      return "ﾏ";
    case 0x30df:
      return "ﾐ";
    case 0x30e0:
      return "ﾑ";
    case 0x30e1:
      return "ﾒ";
    case 0x30e2:
      return "ﾓ";
    case 0x30e4:
      return "ﾔ";
    case 0x30e6:
      return "ﾕ";
    case 0x30e8:
      return "ﾖ";
    case 0x30e9:
      return "ﾗ";
    case 0x30ea:
      return "ﾘ";
    case 0x30eb:
      return "ﾙ";
    case 0x30ec:
      return "ﾚ";
    case 0x30ed:
      return "ﾛ";
    case 0x30ef:
      return "ﾜ";
    case 0x30f3:
      return "ﾝ";
    case 0x30f4:
      return "ｳﾞ";
    case 0x30ac:
      return "ｶﾞ";
    case 0x30ae:
      return "ｷﾞ";
    case 0x30b0:
      return "ｸﾞ";
    case 0x30b2:
      return "ｹﾞ";
    case 0x30b4:
      return "ｺﾞ";
    case 0x30b6:
      return "ｻﾞ";
    case 0x30b8:
      return "ｼﾞ";
    case 0x30ba:
      return "ｽﾞ";
    case 0x30bc:
      return "ｾﾞ";
    case 0x30be:
      return "ｿﾞ";
    case 0x30c0:
      return "ﾀﾞ";
    case 0x30c2:
      return "ﾁﾞ";
    case 0x30c5:
      return "ﾂﾞ";
    case 0x30c7:
      return "ﾃﾞ";
    case 0x30c9:
      return "ﾄﾞ";
    case 0x30d0:
      return "ﾊﾞ";
    case 0x30d3:
      return "ﾋﾞ";
    case 0x30d6:
      return "ﾌﾞ";
    case 0x30d9:
      return "ﾍﾞ";
    case 0x30dc:
      return "ﾎﾞ";
    case 0x30d1:
      return "ﾊﾟ";
    case 0x30d4:
      return "ﾋﾟ";
    case 0x30d7:
      return "ﾌﾟ";
    case 0x30da:
      return "ﾍﾟ";
    case 0x30dd:
      return "ﾎﾟ";
    default:
      return null;
  }
}

export function toJapaneseFullWidth(text: string): string {
  let output = "";
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index) & 0xffff;
    if (isLeadingSurrogate(code) && index + 1 < text.length) {
      const nextCode = text.charCodeAt(index + 1) & 0xffff;
      if (isTrailingSurrogate(nextCode)) {
        output += String.fromCharCode(code) + String.fromCharCode(nextCode);
        index += 1;
        continue;
      }
    }
    if (code == 0x20) {
      output += "　";
      continue;
    }
    if (code >= 0x21 && code <= 0x7e) {
      output += String.fromCharCode(code + 0xfee0);
      continue;
    }
    if (index + 1 < text.length) {
      const voiced = halfWidthVoicedKanaToFullWidthChar(code, text.charCodeAt(index + 1) & 0xffff);
      if (voiced != null) {
        output += voiced;
        index += 1;
        continue;
      }
    }
    const mapped = halfWidthKanaToFullWidthChar(code);
    output += mapped != null ? mapped : String.fromCharCode(code);
  }
  return output;
}

export function toJapaneseHalfWidth(text: string): string {
  let output = "";
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index) & 0xffff;
    if (isLeadingSurrogate(code) && index + 1 < text.length) {
      const nextCode = text.charCodeAt(index + 1) & 0xffff;
      if (isTrailingSurrogate(nextCode)) {
        output += String.fromCharCode(code) + String.fromCharCode(nextCode);
        index += 1;
        continue;
      }
    }
    if (code == 0x3000) {
      output += " ";
      continue;
    }
    if (code >= 0xff01 && code <= 0xff5e) {
      output += String.fromCharCode(code - 0xfee0);
      continue;
    }
    const mapped = fullWidthKanaToHalfWidthChar(code);
    output += mapped != null ? mapped : String.fromCharCode(code);
  }
  return output;
}
