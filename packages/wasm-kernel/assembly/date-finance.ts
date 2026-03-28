import { ValueTag } from "./protocol";

function toNumberExact(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

function truncToInt(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>numeric;
}

function floorDiv(a: i32, b: i32): i32 {
  let quotient = a / b;
  const remainder = a % b;
  if (remainder != 0 && remainder > 0 != b > 0) {
    quotient -= 1;
  }
  return quotient;
}

function daysFromCivil(year: i32, month: i32, day: i32): i32 {
  let adjustedYear = year;
  if (month <= 2) {
    adjustedYear -= 1;
  }
  const era = adjustedYear >= 0 ? adjustedYear / 400 : (adjustedYear - 399) / 400;
  const yearOfEra = adjustedYear - era * 400;
  const shiftedMonth = month + (month > 2 ? -3 : 9);
  const dayOfYear = (153 * shiftedMonth + 2) / 5 + day - 1;
  const dayOfEra = yearOfEra * 365 + yearOfEra / 4 - yearOfEra / 100 + dayOfYear;
  return era * 146097 + dayOfEra - 719468;
}

function civilYear(days: i32): i32 {
  const shifted = days + 719468;
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097;
  const dayOfEra = shifted - era * 146097;
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365;
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100);
  const monthPrime = (5 * dayOfYear + 2) / 153;
  const month = monthPrime + (monthPrime < 10 ? 3 : -9);
  return yearOfEra + era * 400 + (month <= 2 ? 1 : 0);
}

function civilMonth(days: i32): i32 {
  const shifted = days + 719468;
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097;
  const dayOfEra = shifted - era * 146097;
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365;
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100);
  const monthPrime = (5 * dayOfYear + 2) / 153;
  return monthPrime + (monthPrime < 10 ? 3 : -9);
}

function civilDay(days: i32): i32 {
  const shifted = days + 719468;
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097;
  const dayOfEra = shifted - era * 146097;
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365;
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100);
  const monthPrime = (5 * dayOfYear + 2) / 153;
  return dayOfYear - (153 * monthPrime + 2) / 5 + 1;
}

const EXCEL_EPOCH_DAYS: i32 = -25568;
export const EXCEL_SECONDS_PER_DAY: i32 = 86400;

export function excelSerialWhole(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>Math.floor(numeric);
}

function daysInExcelMonth(year: i32, month: i32): i32 {
  if (year == 1900 && month == 2) {
    return 29;
  }
  const start = daysFromCivil(year, month, 1);
  const nextMonth = month == 12 ? 1 : month + 1;
  const nextYear = month == 12 ? year + 1 : year;
  const end = daysFromCivil(nextYear, nextMonth, 1);
  return end - start;
}

function isLeapYear(year: i32): bool {
  return year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
}

export function excelYearPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }
  if (whole == 60) {
    return 1900;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return civilYear(EXCEL_EPOCH_DAYS + adjustedWhole);
}

export function excelMonthPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }
  if (whole == 60) {
    return 2;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return civilMonth(EXCEL_EPOCH_DAYS + adjustedWhole);
}

export function excelDayPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }
  if (whole == 60) {
    return 29;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return civilDay(EXCEL_EPOCH_DAYS + adjustedWhole);
}

export function excelTimeSerial(
  hourTag: u8,
  hourValue: f64,
  minuteTag: u8,
  minuteValue: f64,
  secondTag: u8,
  secondValue: f64,
): f64 {
  const hourNumeric = toNumberExact(hourTag, hourValue);
  const minuteNumeric = toNumberExact(minuteTag, minuteValue);
  const secondNumeric = toNumberExact(secondTag, secondValue);
  if (isNaN(hourNumeric) || isNaN(minuteNumeric) || isNaN(secondNumeric)) {
    return NaN;
  }
  const hour = <f64>(<i32>hourNumeric);
  const minute = <f64>(<i32>minuteNumeric);
  const second = <f64>(<i32>secondNumeric);
  if (hour < 0 || minute < 0 || second < 0 || hour > 32767 || minute > 32767 || second > 32767) {
    return NaN;
  }
  let totalSeconds = hour * 3600.0 + minute * 60.0 + second;
  totalSeconds %= <f64>EXCEL_SECONDS_PER_DAY;
  if (totalSeconds < 0) {
    totalSeconds += <f64>EXCEL_SECONDS_PER_DAY;
  }
  return totalSeconds / <f64>EXCEL_SECONDS_PER_DAY;
}

export function excelSecondOfDay(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric) || numeric < 0) {
    return i32.MIN_VALUE;
  }
  const whole = Math.floor(numeric);
  let fraction = numeric - whole;
  if (fraction < 0) {
    fraction += 1.0;
  }
  let seconds = <i32>Math.floor(fraction * <f64>EXCEL_SECONDS_PER_DAY + 1e-9);
  if (seconds >= EXCEL_SECONDS_PER_DAY) {
    seconds = 0;
  }
  return seconds;
}

export function excelWeekdayFromSerial(tag: u8, value: f64, returnType: i32): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE || whole < 0) {
    return i32.MIN_VALUE;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  const sundayOne = (((adjustedWhole % 7) + 7) % 7) + 1;
  if (returnType == 1) {
    return sundayOne;
  }
  if (returnType == 3) {
    return sundayOne == 1 ? 6 : sundayOne - 2;
  }

  let startDay = 0;
  if (returnType == 2 || returnType == 11) {
    startDay = 2;
  } else if (returnType == 12) {
    startDay = 3;
  } else if (returnType == 13) {
    startDay = 4;
  } else if (returnType == 14) {
    startDay = 5;
  } else if (returnType == 15) {
    startDay = 6;
  } else if (returnType == 16) {
    startDay = 7;
  } else if (returnType == 17) {
    startDay = 1;
  } else {
    return i32.MIN_VALUE;
  }

  return ((sundayOne - startDay + 7) % 7) + 1;
}

export function excelWeeknumFromSerial(tag: u8, value: f64, returnType: i32): i32 {
  const year = excelYearPartFromSerial(tag, value);
  const month = excelMonthPartFromSerial(tag, value);
  const day = excelDayPartFromSerial(tag, value);
  if (year == i32.MIN_VALUE || month == i32.MIN_VALUE || day == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }

  let weekStartDay = 0;
  if (returnType == 1 || returnType == 17) {
    weekStartDay = 0;
  } else if (returnType == 2 || returnType == 11) {
    weekStartDay = 1;
  } else if (returnType == 12) {
    weekStartDay = 2;
  } else if (returnType == 13) {
    weekStartDay = 3;
  } else if (returnType == 14) {
    weekStartDay = 4;
  } else if (returnType == 15) {
    weekStartDay = 5;
  } else if (returnType == 16) {
    weekStartDay = 6;
  } else {
    return i32.MIN_VALUE;
  }

  const jan1Serial = excelDateSerial(
    <u8>ValueTag.Number,
    <f64>year,
    <u8>ValueTag.Number,
    1.0,
    <u8>ValueTag.Number,
    1.0,
  );
  if (isNaN(jan1Serial)) {
    return i32.MIN_VALUE;
  }

  let adjustedJan1 = <i32>Math.floor(jan1Serial);
  adjustedJan1 = adjustedJan1 < 60 ? adjustedJan1 : adjustedJan1 - 1;
  const jan1Weekday = ((adjustedJan1 % 7) + 7) % 7;
  const shift = (jan1Weekday - weekStartDay + 7) % 7;

  let dayOfYear = day;
  for (let currentMonth = 1; currentMonth < month; currentMonth += 1) {
    dayOfYear += daysInExcelMonth(year, currentMonth);
  }

  return <i32>Math.floor(<f64>(dayOfYear - 1 + shift) / 7.0) + 1;
}

export function excelDateSerial(
  yearTag: u8,
  yearValue: f64,
  monthTag: u8,
  monthValue: f64,
  dayTag: u8,
  dayValue: f64,
): f64 {
  let year = truncToInt(yearTag, yearValue);
  const month = truncToInt(monthTag, monthValue);
  const day = truncToInt(dayTag, dayValue);
  if (year == i32.MIN_VALUE || month == i32.MIN_VALUE || day == i32.MIN_VALUE) {
    return NaN;
  }
  if (year >= 0 && year <= 1899) {
    year += 1900;
  }
  if (year < 0 || year > 9999) {
    return NaN;
  }
  if (year == 1900 && month == 2 && day == 29) {
    return 60;
  }

  const zeroBasedMonth = month - 1;
  const monthQuotient = floorDiv(zeroBasedMonth, 12);
  const normalizedYear = year + monthQuotient;
  const normalizedMonthZero = zeroBasedMonth - monthQuotient * 12;
  if (normalizedYear < 0 || normalizedYear > 9999) {
    return NaN;
  }
  const days = daysFromCivil(normalizedYear, normalizedMonthZero + 1, 1) + (day - 1);
  let serial = days - EXCEL_EPOCH_DAYS;
  if (days >= daysFromCivil(1900, 3, 1)) {
    serial += 1;
  }
  return <f64>serial;
}

export function addMonthsExcelSerial(
  tag: u8,
  value: f64,
  offsetTag: u8,
  offsetValue: f64,
  endOfMonth: bool,
): f64 {
  const startYear = excelYearPartFromSerial(tag, value);
  const startMonth = excelMonthPartFromSerial(tag, value);
  const startDay = excelDayPartFromSerial(tag, value);
  const offset = truncToInt(offsetTag, offsetValue);
  if (
    startYear == i32.MIN_VALUE ||
    startMonth == i32.MIN_VALUE ||
    startDay == i32.MIN_VALUE ||
    offset == i32.MIN_VALUE
  ) {
    return NaN;
  }

  const totalMonths = startYear * 12 + (startMonth - 1) + offset;
  const shiftedYear = floorDiv(totalMonths, 12);
  const shiftedMonth = totalMonths - shiftedYear * 12 + 1;
  if (shiftedYear < 0 || shiftedYear > 9999) {
    return NaN;
  }

  const targetDay = endOfMonth
    ? daysInExcelMonth(shiftedYear, shiftedMonth)
    : min<i32>(startDay, daysInExcelMonth(shiftedYear, shiftedMonth));
  if (shiftedYear == 1900 && shiftedMonth == 2 && targetDay == 29) {
    return 60;
  }
  return excelDateSerial(
    <u8>ValueTag.Number,
    <f64>shiftedYear,
    <u8>ValueTag.Number,
    <f64>shiftedMonth,
    <u8>ValueTag.Number,
    <f64>targetDay,
  );
}

export function excelDatedifValue(startWhole: i32, endWhole: i32, unit: string): f64 {
  if (startWhole > endWhole) {
    return NaN;
  }

  const startYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
  const startMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
  const startDay = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
  const endYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
  const endMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
  const endDay = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
  if (
    startYear == i32.MIN_VALUE ||
    startMonth == i32.MIN_VALUE ||
    startDay == i32.MIN_VALUE ||
    endYear == i32.MIN_VALUE ||
    endMonth == i32.MIN_VALUE ||
    endDay == i32.MIN_VALUE
  ) {
    return NaN;
  }

  const totalDays = endWhole - startWhole;
  const totalMonths =
    (endYear - startYear) * 12 + (endMonth - startMonth) - (endDay < startDay ? 1 : 0);
  const totalYears =
    endYear -
    startYear -
    (endMonth < startMonth || (endMonth == startMonth && endDay < startDay) ? 1 : 0);

  if (unit == "D") return <f64>totalDays;
  if (unit == "M") return <f64>totalMonths;
  if (unit == "Y") return <f64>totalYears;
  if (unit == "YM") return <f64>(((totalMonths % 12) + 12) % 12);
  if (unit == "YD") {
    let comparisonYear = endYear;
    let comparison = excelDateSerial(
      <u8>ValueTag.Number,
      <f64>comparisonYear,
      <u8>ValueTag.Number,
      <f64>startMonth,
      <u8>ValueTag.Number,
      <f64>startDay,
    );
    if (isNaN(comparison) || comparison > <f64>endWhole) {
      comparisonYear -= 1;
      comparison = excelDateSerial(
        <u8>ValueTag.Number,
        <f64>comparisonYear,
        <u8>ValueTag.Number,
        <f64>startMonth,
        <u8>ValueTag.Number,
        <f64>startDay,
      );
    }
    return isNaN(comparison) ? NaN : <f64>(endWhole - <i32>Math.floor(comparison));
  }
  if (unit == "MD") {
    if (endDay >= startDay) {
      return <f64>(endDay - startDay);
    }
    const previousMonth = endMonth == 1 ? 12 : endMonth - 1;
    return <f64>(daysInExcelMonth(endYear, previousMonth) - startDay + endDay);
  }
  return NaN;
}

export function excelDays360Value(startWhole: i32, endWhole: i32, method: i32): f64 {
  const startYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
  const startMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
  const startDayRaw = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
  const endYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
  const endMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
  const endDayRaw = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
  if (
    startYear == i32.MIN_VALUE ||
    startMonth == i32.MIN_VALUE ||
    startDayRaw == i32.MIN_VALUE ||
    endYear == i32.MIN_VALUE ||
    endMonth == i32.MIN_VALUE ||
    endDayRaw == i32.MIN_VALUE
  ) {
    return NaN;
  }

  let startDay = startDayRaw;
  let endDay = endDayRaw;
  if (method == 0) {
    if (startDay == 31) {
      startDay = 30;
    }
    if (endDay == 31 && startDay >= 30) {
      endDay = 30;
    }
  } else {
    if (startDay == 31) {
      startDay = 30;
    }
    if (endDay == 31) {
      endDay = 30;
    }
  }

  return <f64>((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay));
}

export function excelYearfracValue(startWhole: i32, endWhole: i32, basis: i32): f64 {
  if (basis < 0 || basis > 4) {
    return NaN;
  }

  let start = startWhole;
  let end = endWhole;
  if (start > end) {
    const swapped = start;
    start = end;
    end = swapped;
  }

  let startYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>start);
  let startMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>start);
  let startDay = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>start);
  let endYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>end);
  let endMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>end);
  let endDay = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>end);
  if (
    startYear == i32.MIN_VALUE ||
    startMonth == i32.MIN_VALUE ||
    startDay == i32.MIN_VALUE ||
    endYear == i32.MIN_VALUE ||
    endMonth == i32.MIN_VALUE ||
    endDay == i32.MIN_VALUE
  ) {
    return NaN;
  }

  let totalDays = 0.0;
  if (basis == 0) {
    if (startDay == 31) {
      startDay -= 1;
    }
    if (startDay == 30 && endDay == 31) {
      endDay -= 1;
    } else if (startMonth == 2 && startDay == (isLeapYear(startYear) ? 29 : 28)) {
      startDay = 30;
      if (endMonth == 2 && endDay == (isLeapYear(endYear) ? 29 : 28)) {
        endDay = 30;
      }
    }
    totalDays = <f64>(
      ((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay))
    );
  } else if (basis == 1 || basis == 2 || basis == 3) {
    totalDays = <f64>(end - start);
  } else {
    if (startDay == 31) {
      startDay -= 1;
    }
    if (endDay == 31) {
      endDay -= 1;
    }
    totalDays = <f64>(
      ((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay))
    );
  }

  let daysInYear = 360.0;
  if (basis == 1) {
    if (startYear == endYear) {
      daysInYear = isLeapYear(startYear) ? 366.0 : 365.0;
    } else {
      const crossesMultipleYears =
        endYear != startYear + 1 ||
        endMonth < startMonth ||
        (endMonth == startMonth && endDay > startDay);
      if (crossesMultipleYears) {
        let total = 0.0;
        for (let year = startYear; year <= endYear; year += 1) {
          total += isLeapYear(year) ? 366.0 : 365.0;
        }
        daysInYear = total / <f64>(endYear - startYear + 1);
      } else {
        const startsInLeapYear =
          isLeapYear(startYear) && (startMonth < 2 || (startMonth == 2 && startDay <= 29));
        const endsInLeapYear =
          isLeapYear(endYear) && (endMonth > 2 || (endMonth == 2 && endDay == 29));
        daysInYear = startsInLeapYear || endsInLeapYear ? 366.0 : 365.0;
      }
    }
  } else if (basis == 3) {
    daysInYear = 365.0;
  }

  return totalDays / daysInYear;
}

export function securityAnnualizedYearfracValue(
  settlementWhole: i32,
  maturityWhole: i32,
  basis: i32,
): f64 {
  if (settlementWhole >= maturityWhole) {
    return NaN;
  }
  const years = excelYearfracValue(settlementWhole, maturityWhole, basis);
  return isNaN(years) || years <= 0.0 ? NaN : years;
}

export function treasuryBillDaysValue(settlementWhole: i32, maturityWhole: i32): f64 {
  if (settlementWhole >= maturityWhole) {
    return NaN;
  }
  const days = maturityWhole - settlementWhole;
  return days > 0 && days <= 365 ? <f64>days : NaN;
}

export function maturityIssueYearfracValue(
  issueWhole: i32,
  settlementWhole: i32,
  maturityWhole: i32,
  basis: i32,
): f64 {
  if (
    issueWhole >= settlementWhole ||
    issueWhole >= maturityWhole ||
    settlementWhole >= maturityWhole
  ) {
    return NaN;
  }
  return excelYearfracValue(issueWhole, maturityWhole, basis);
}

export function accruedIssueYearfracValue(
  issueWhole: i32,
  settlementWhole: i32,
  maturityWhole: i32,
  basis: i32,
): f64 {
  if (
    issueWhole >= settlementWhole ||
    issueWhole >= maturityWhole ||
    settlementWhole >= maturityWhole
  ) {
    return NaN;
  }
  return excelYearfracValue(issueWhole, settlementWhole, basis);
}

export function couponDaysByBasisValue(startWhole: i32, endWhole: i32, basis: i32): f64 {
  if (basis < 0 || basis > 4 || startWhole > endWhole) {
    return NaN;
  }
  if (basis == 0 || basis == 4) {
    const startYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
    const startMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
    const startDayRaw = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>startWhole);
    const endYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
    const endMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
    const endDayRaw = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>endWhole);
    if (
      startYear == i32.MIN_VALUE ||
      startMonth == i32.MIN_VALUE ||
      startDayRaw == i32.MIN_VALUE ||
      endYear == i32.MIN_VALUE ||
      endMonth == i32.MIN_VALUE ||
      endDayRaw == i32.MIN_VALUE
    ) {
      return NaN;
    }
    let startDay = startDayRaw;
    let endDay = endDayRaw;
    if (basis == 0) {
      if (startDay == 31) {
        startDay = 30;
      }
      if (startDay == 30 && endDay == 31) {
        endDay = 30;
      } else if (startMonth == 2 && startDay == (isLeapYear(startYear) ? 29 : 28)) {
        startDay = 30;
        if (endMonth == 2 && endDay == (isLeapYear(endYear) ? 29 : 28)) {
          endDay = 30;
        }
      }
    } else {
      if (startDay == 31) {
        startDay = 30;
      }
      if (endDay == 31) {
        endDay = 30;
      }
    }
    return <f64>((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay));
  }
  return <f64>(endWhole - startWhole);
}

export function couponDateFromMaturityValue(
  maturityWhole: i32,
  periodsBack: i32,
  frequency: i32,
): i32 {
  if (periodsBack < 0 || (frequency != 1 && frequency != 2 && frequency != 4)) {
    return i32.MIN_VALUE;
  }
  const stepMonths = 12 / frequency;
  const serial = addMonthsExcelSerial(
    <u8>ValueTag.Number,
    <f64>maturityWhole,
    <u8>ValueTag.Number,
    <f64>(-periodsBack * stepMonths),
    false,
  );
  return isNaN(serial) ? i32.MIN_VALUE : <i32>serial;
}

export function couponPeriodsRemainingValue(
  settlementWhole: i32,
  maturityWhole: i32,
  frequency: i32,
): i32 {
  if (settlementWhole >= maturityWhole || (frequency != 1 && frequency != 2 && frequency != 4)) {
    return i32.MIN_VALUE;
  }
  let periodsRemaining = 1;
  let previousCoupon = couponDateFromMaturityValue(maturityWhole, periodsRemaining, frequency);
  while (previousCoupon != i32.MIN_VALUE && previousCoupon > settlementWhole) {
    periodsRemaining += 1;
    previousCoupon = couponDateFromMaturityValue(maturityWhole, periodsRemaining, frequency);
  }
  return previousCoupon == i32.MIN_VALUE ? i32.MIN_VALUE : periodsRemaining;
}

export function couponPeriodDaysValue(
  previousCoupon: i32,
  nextCoupon: i32,
  basis: i32,
  frequency: i32,
): f64 {
  if (basis == 1) {
    return couponDaysByBasisValue(previousCoupon, nextCoupon, basis);
  }
  return <f64>(basis == 3 ? 365 : 360) / <f64>frequency;
}

export function oddLastPriceValue(
  settlementWhole: i32,
  maturityWhole: i32,
  lastInterestWhole: i32,
  rate: f64,
  yieldRate: f64,
  redemption: f64,
  frequency: i32,
  basis: i32,
): f64 {
  if (
    lastInterestWhole >= settlementWhole ||
    settlementWhole >= maturityWhole ||
    !isFinite(rate) ||
    !isFinite(yieldRate) ||
    !isFinite(redemption) ||
    rate < 0.0 ||
    yieldRate < 0.0 ||
    redemption <= 0.0 ||
    (frequency != 1 && frequency != 2 && frequency != 4) ||
    basis < 0 ||
    basis > 4
  ) {
    return NaN;
  }

  const stepMonths = 12 / frequency;
  let periodStart = lastInterestWhole;
  let accruedFraction = 0.0;
  let remainingFraction = 0.0;
  let totalFraction = 0.0;
  let iterations = 0;
  while (periodStart < maturityWhole && iterations < 32) {
    const normalEndRaw = addMonthsExcelSerial(
      <u8>ValueTag.Number,
      <f64>periodStart,
      <u8>ValueTag.Number,
      <f64>stepMonths,
      false,
    );
    if (isNaN(normalEndRaw)) {
      return NaN;
    }
    const normalEnd = <i32>normalEndRaw;
    if (normalEnd <= periodStart) {
      return NaN;
    }
    const actualEnd = normalEnd < maturityWhole ? normalEnd : maturityWhole;
    const normalDays = couponDaysByBasisValue(periodStart, normalEnd, basis);
    const countedDays = couponDaysByBasisValue(periodStart, actualEnd, basis);
    if (isNaN(normalDays) || isNaN(countedDays) || normalDays <= 0.0 || countedDays < 0.0) {
      return NaN;
    }
    totalFraction += countedDays / normalDays;

    if (settlementWhole > periodStart) {
      const accruedEnd = settlementWhole < actualEnd ? settlementWhole : actualEnd;
      const accruedDays = couponDaysByBasisValue(periodStart, accruedEnd, basis);
      if (isNaN(accruedDays) || accruedDays < 0.0) {
        return NaN;
      }
      accruedFraction += accruedDays / normalDays;
    }

    if (settlementWhole < actualEnd) {
      const remainingStart = settlementWhole > periodStart ? settlementWhole : periodStart;
      const remainingDays = couponDaysByBasisValue(remainingStart, actualEnd, basis);
      if (isNaN(remainingDays) || remainingDays < 0.0) {
        return NaN;
      }
      remainingFraction += remainingDays / normalDays;
    }

    periodStart = actualEnd;
    iterations += 1;
  }

  if (
    periodStart != maturityWhole ||
    iterations >= 32 ||
    remainingFraction <= 0.0 ||
    totalFraction <= 0.0
  ) {
    return NaN;
  }

  const coupon = (100.0 * rate) / <f64>frequency;
  const denominator = 1.0 + (yieldRate * remainingFraction) / <f64>frequency;
  if (!isFinite(denominator) || denominator <= 0.0) {
    return NaN;
  }
  return (redemption + coupon * totalFraction) / denominator - coupon * accruedFraction;
}

export function oddLastYieldValue(
  settlementWhole: i32,
  maturityWhole: i32,
  lastInterestWhole: i32,
  rate: f64,
  price: f64,
  redemption: f64,
  frequency: i32,
  basis: i32,
): f64 {
  if (
    lastInterestWhole >= settlementWhole ||
    settlementWhole >= maturityWhole ||
    !isFinite(rate) ||
    !isFinite(price) ||
    !isFinite(redemption) ||
    rate < 0.0 ||
    price <= 0.0 ||
    redemption <= 0.0 ||
    (frequency != 1 && frequency != 2 && frequency != 4) ||
    basis < 0 ||
    basis > 4
  ) {
    return NaN;
  }

  const stepMonths = 12 / frequency;
  let periodStart = lastInterestWhole;
  let accruedFraction = 0.0;
  let remainingFraction = 0.0;
  let totalFraction = 0.0;
  let iterations = 0;
  while (periodStart < maturityWhole && iterations < 32) {
    const normalEndRaw = addMonthsExcelSerial(
      <u8>ValueTag.Number,
      <f64>periodStart,
      <u8>ValueTag.Number,
      <f64>stepMonths,
      false,
    );
    if (isNaN(normalEndRaw)) {
      return NaN;
    }
    const normalEnd = <i32>normalEndRaw;
    if (normalEnd <= periodStart) {
      return NaN;
    }
    const actualEnd = normalEnd < maturityWhole ? normalEnd : maturityWhole;
    const normalDays = couponDaysByBasisValue(periodStart, normalEnd, basis);
    const countedDays = couponDaysByBasisValue(periodStart, actualEnd, basis);
    if (isNaN(normalDays) || isNaN(countedDays) || normalDays <= 0.0 || countedDays < 0.0) {
      return NaN;
    }
    totalFraction += countedDays / normalDays;

    if (settlementWhole > periodStart) {
      const accruedEnd = settlementWhole < actualEnd ? settlementWhole : actualEnd;
      const accruedDays = couponDaysByBasisValue(periodStart, accruedEnd, basis);
      if (isNaN(accruedDays) || accruedDays < 0.0) {
        return NaN;
      }
      accruedFraction += accruedDays / normalDays;
    }

    if (settlementWhole < actualEnd) {
      const remainingStart = settlementWhole > periodStart ? settlementWhole : periodStart;
      const remainingDays = couponDaysByBasisValue(remainingStart, actualEnd, basis);
      if (isNaN(remainingDays) || remainingDays < 0.0) {
        return NaN;
      }
      remainingFraction += remainingDays / normalDays;
    }

    periodStart = actualEnd;
    iterations += 1;
  }

  if (
    periodStart != maturityWhole ||
    iterations >= 32 ||
    remainingFraction <= 0.0 ||
    totalFraction <= 0.0
  ) {
    return NaN;
  }

  const coupon = (100.0 * rate) / <f64>frequency;
  const dirtyPrice = price + coupon * accruedFraction;
  if (!isFinite(dirtyPrice) || dirtyPrice <= 0.0) {
    return NaN;
  }
  const maturityValue = redemption + coupon * totalFraction;
  return ((maturityValue - dirtyPrice) / dirtyPrice) * (<f64>frequency / remainingFraction);
}

export function oddFirstPriceValue(
  settlementWhole: i32,
  maturityWhole: i32,
  issueWhole: i32,
  firstCouponWhole: i32,
  rate: f64,
  yieldRate: f64,
  redemption: f64,
  frequency: i32,
  basis: i32,
): f64 {
  if (
    issueWhole >= settlementWhole ||
    settlementWhole >= firstCouponWhole ||
    firstCouponWhole >= maturityWhole ||
    !isFinite(rate) ||
    !isFinite(yieldRate) ||
    !isFinite(redemption) ||
    rate < 0.0 ||
    redemption <= 0.0 ||
    (frequency != 1 && frequency != 2 && frequency != 4) ||
    basis < 0 ||
    basis > 4
  ) {
    return NaN;
  }

  const stepMonths = 12 / frequency;
  const actualStarts = new Array<i32>();
  const periodEnds = new Array<i32>();
  const normalDaysList = new Array<f64>();
  const countedDaysList = new Array<f64>();

  let periodEnd = firstCouponWhole;
  let iterations = 0;
  while (periodEnd > issueWhole && iterations < 64) {
    const normalStartRaw = addMonthsExcelSerial(
      <u8>ValueTag.Number,
      <f64>periodEnd,
      <u8>ValueTag.Number,
      <f64>-stepMonths,
      false,
    );
    if (isNaN(normalStartRaw)) {
      return NaN;
    }
    const normalStart = <i32>normalStartRaw;
    if (normalStart >= periodEnd) {
      return NaN;
    }
    const actualStart = normalStart > issueWhole ? normalStart : issueWhole;
    const normalDays = couponDaysByBasisValue(normalStart, periodEnd, basis);
    const countedDays = couponDaysByBasisValue(actualStart, periodEnd, basis);
    if (isNaN(normalDays) || isNaN(countedDays) || normalDays <= 0.0 || countedDays < 0.0) {
      return NaN;
    }
    actualStarts.push(actualStart);
    periodEnds.push(periodEnd);
    normalDaysList.push(normalDays);
    countedDaysList.push(countedDays);
    periodEnd = actualStart;
    iterations += 1;
  }

  if (periodEnd != issueWhole || iterations >= 64 || actualStarts.length == 0) {
    return NaN;
  }

  let accruedFraction = 0.0;
  let remainingFraction = 0.0;
  let totalFraction = 0.0;
  for (let index = actualStarts.length - 1; index >= 0; index -= 1) {
    const actualStart = unchecked(actualStarts[index]);
    const segmentEnd = unchecked(periodEnds[index]);
    const normalDays = unchecked(normalDaysList[index]);
    const countedDays = unchecked(countedDaysList[index]);
    totalFraction += countedDays / normalDays;

    if (settlementWhole > actualStart) {
      const accruedEnd = settlementWhole < segmentEnd ? settlementWhole : segmentEnd;
      const accruedDays = couponDaysByBasisValue(actualStart, accruedEnd, basis);
      if (isNaN(accruedDays) || accruedDays < 0.0) {
        return NaN;
      }
      accruedFraction += accruedDays / normalDays;
    }

    if (settlementWhole < segmentEnd) {
      const remainingStart = settlementWhole > actualStart ? settlementWhole : actualStart;
      const remainingDays = couponDaysByBasisValue(remainingStart, segmentEnd, basis);
      if (isNaN(remainingDays) || remainingDays < 0.0) {
        return NaN;
      }
      remainingFraction += remainingDays / normalDays;
    }
  }

  let regularPeriodsAfterFirst = 0;
  let couponDate = firstCouponWhole;
  while (couponDate < maturityWhole && regularPeriodsAfterFirst < 256) {
    const nextCouponRaw = addMonthsExcelSerial(
      <u8>ValueTag.Number,
      <f64>couponDate,
      <u8>ValueTag.Number,
      <f64>stepMonths,
      false,
    );
    if (isNaN(nextCouponRaw)) {
      return NaN;
    }
    const nextCouponDate = <i32>nextCouponRaw;
    if (nextCouponDate <= couponDate) {
      return NaN;
    }
    couponDate = nextCouponDate;
    regularPeriodsAfterFirst += 1;
  }

  if (
    couponDate != maturityWhole ||
    regularPeriodsAfterFirst <= 0 ||
    regularPeriodsAfterFirst >= 256 ||
    remainingFraction <= 0.0 ||
    totalFraction <= 0.0
  ) {
    return NaN;
  }

  const discountBase = 1.0 + yieldRate / <f64>frequency;
  if (!isFinite(discountBase) || discountBase <= 0.0) {
    return NaN;
  }
  const coupon = (100.0 * rate) / <f64>frequency;
  let price =
    (coupon * totalFraction) / Math.pow(discountBase, remainingFraction) - coupon * accruedFraction;

  for (let period = 1; period <= regularPeriodsAfterFirst; period += 1) {
    const exponent = remainingFraction + <f64>period;
    const cashflow = period == regularPeriodsAfterFirst ? redemption + coupon : coupon;
    price += cashflow / Math.pow(discountBase, exponent);
  }
  return price;
}

export function oddFirstYieldValue(
  settlementWhole: i32,
  maturityWhole: i32,
  issueWhole: i32,
  firstCouponWhole: i32,
  rate: f64,
  price: f64,
  redemption: f64,
  frequency: i32,
  basis: i32,
): f64 {
  if (
    issueWhole >= settlementWhole ||
    settlementWhole >= firstCouponWhole ||
    firstCouponWhole >= maturityWhole ||
    !isFinite(rate) ||
    !isFinite(price) ||
    !isFinite(redemption) ||
    rate < 0.0 ||
    price <= 0.0 ||
    redemption <= 0.0 ||
    (frequency != 1 && frequency != 2 && frequency != 4) ||
    basis < 0 ||
    basis > 4
  ) {
    return NaN;
  }

  let lower = -(<f64>frequency) + 1e-10;
  let upper = Math.max(1.0, rate * 2.0 + 0.1);
  let lowerPrice = oddFirstPriceValue(
    settlementWhole,
    maturityWhole,
    issueWhole,
    firstCouponWhole,
    rate,
    lower,
    redemption,
    frequency,
    basis,
  );
  let upperPrice = oddFirstPriceValue(
    settlementWhole,
    maturityWhole,
    issueWhole,
    firstCouponWhole,
    rate,
    upper,
    redemption,
    frequency,
    basis,
  );
  for (
    let iteration = 0;
    iteration < 200 &&
    (isNaN(lowerPrice) || isNaN(upperPrice) || lowerPrice < price || upperPrice > price);
    iteration += 1
  ) {
    if (isNaN(upperPrice) || upperPrice > price) {
      upper = upper * 2.0 + 1.0;
      upperPrice = oddFirstPriceValue(
        settlementWhole,
        maturityWhole,
        issueWhole,
        firstCouponWhole,
        rate,
        upper,
        redemption,
        frequency,
        basis,
      );
      continue;
    }
    lower = (lower - <f64>frequency) / 2.0;
    lowerPrice = oddFirstPriceValue(
      settlementWhole,
      maturityWhole,
      issueWhole,
      firstCouponWhole,
      rate,
      lower,
      redemption,
      frequency,
      basis,
    );
  }

  if (isNaN(lowerPrice) || isNaN(upperPrice) || lowerPrice < price || upperPrice > price) {
    return NaN;
  }

  let guess = Math.min(Math.max(rate, lower + 1e-8), upper - 1e-8);
  for (let iteration = 0; iteration < 200; iteration += 1) {
    const estimatedPrice = oddFirstPriceValue(
      settlementWhole,
      maturityWhole,
      issueWhole,
      firstCouponWhole,
      rate,
      guess,
      redemption,
      frequency,
      basis,
    );
    if (isNaN(estimatedPrice)) {
      return NaN;
    }
    const error = estimatedPrice - price;
    if (Math.abs(error) < 1e-14) {
      return guess;
    }

    const epsilon = Math.max(1e-7, Math.abs(guess) * 1e-6);
    const shiftedPrice = oddFirstPriceValue(
      settlementWhole,
      maturityWhole,
      issueWhole,
      firstCouponWhole,
      rate,
      guess + epsilon,
      redemption,
      frequency,
      basis,
    );
    const derivative = isNaN(shiftedPrice) ? NaN : (shiftedPrice - estimatedPrice) / epsilon;
    let nextGuess =
      isNaN(derivative) || !isFinite(derivative) || derivative == 0.0
        ? (lower + upper) / 2.0
        : guess - error / derivative;
    if (!isFinite(nextGuess) || nextGuess <= lower || nextGuess >= upper) {
      nextGuess = (lower + upper) / 2.0;
    }

    const boundedPrice = oddFirstPriceValue(
      settlementWhole,
      maturityWhole,
      issueWhole,
      firstCouponWhole,
      rate,
      nextGuess,
      redemption,
      frequency,
      basis,
    );
    if (isNaN(boundedPrice)) {
      return NaN;
    }
    if (boundedPrice > price) {
      lower = nextGuess;
    } else {
      upper = nextGuess;
    }
    guess = nextGuess;
    if (Math.abs(upper - lower) < 1e-14) {
      return (lower + upper) / 2.0;
    }
  }

  return (lower + upper) / 2.0;
}

export function couponPriceFromMetricsValue(
  periodsRemaining: i32,
  accruedDays: f64,
  daysToNextCoupon: f64,
  daysInPeriod: f64,
  rate: f64,
  yieldRate: f64,
  redemption: f64,
  frequency: i32,
): f64 {
  if (
    periodsRemaining < 1 ||
    !isFinite(accruedDays) ||
    !isFinite(daysToNextCoupon) ||
    !isFinite(daysInPeriod) ||
    !isFinite(rate) ||
    !isFinite(yieldRate) ||
    !isFinite(redemption) ||
    daysInPeriod <= 0.0 ||
    redemption <= 0.0
  ) {
    return NaN;
  }
  const coupon = (100.0 * rate) / <f64>frequency;
  const periodsToNextCoupon = daysToNextCoupon / daysInPeriod;
  if (periodsRemaining == 1) {
    const denominator = 1.0 + (yieldRate / <f64>frequency) * periodsToNextCoupon;
    return denominator <= 0.0
      ? NaN
      : (redemption + coupon) / denominator - coupon * (accruedDays / daysInPeriod);
  }
  const discountBase = 1.0 + yieldRate / <f64>frequency;
  if (discountBase <= 0.0) {
    return NaN;
  }
  let price = 0.0;
  for (let period = 1; period <= periodsRemaining; period += 1) {
    const periodsToCashflow = <f64>(period - 1) + periodsToNextCoupon;
    price += coupon / Math.pow(discountBase, periodsToCashflow);
  }
  price += redemption / Math.pow(discountBase, <f64>(periodsRemaining - 1) + periodsToNextCoupon);
  return price - coupon * (accruedDays / daysInPeriod);
}

export function solveCouponYieldValue(
  periodsRemaining: i32,
  accruedDays: f64,
  daysToNextCoupon: f64,
  daysInPeriod: f64,
  rate: f64,
  price: f64,
  redemption: f64,
  frequency: i32,
): f64 {
  if (periodsRemaining < 1 || !isFinite(price) || price <= 0.0) {
    return NaN;
  }
  const coupon = (100.0 * rate) / <f64>frequency;
  if (periodsRemaining == 1) {
    const dirtyPrice = price + coupon * (accruedDays / daysInPeriod);
    if (dirtyPrice <= 0.0 || daysToNextCoupon <= 0.0) {
      return NaN;
    }
    return (
      ((redemption + coupon) / dirtyPrice - 1.0) *
      <f64>frequency *
      (daysInPeriod / daysToNextCoupon)
    );
  }

  const targetPrice = price;
  let lower = -(<f64>frequency) + 1e-10;
  let upper = max<f64>(1.0, rate * 2.0 + 0.1);
  let lowerPrice = couponPriceFromMetricsValue(
    periodsRemaining,
    accruedDays,
    daysToNextCoupon,
    daysInPeriod,
    rate,
    lower,
    redemption,
    frequency,
  );
  let upperPrice = couponPriceFromMetricsValue(
    periodsRemaining,
    accruedDays,
    daysToNextCoupon,
    daysInPeriod,
    rate,
    upper,
    redemption,
    frequency,
  );
  for (
    let iteration = 0;
    iteration < 100 &&
    (isNaN(lowerPrice) ||
      isNaN(upperPrice) ||
      lowerPrice < targetPrice ||
      upperPrice > targetPrice);
    iteration += 1
  ) {
    if (isNaN(upperPrice) || upperPrice > targetPrice) {
      upper = upper * 2.0 + 1.0;
      upperPrice = couponPriceFromMetricsValue(
        periodsRemaining,
        accruedDays,
        daysToNextCoupon,
        daysInPeriod,
        rate,
        upper,
        redemption,
        frequency,
      );
      continue;
    }
    lower = (lower - <f64>frequency) / 2.0;
    lowerPrice = couponPriceFromMetricsValue(
      periodsRemaining,
      accruedDays,
      daysToNextCoupon,
      daysInPeriod,
      rate,
      lower,
      redemption,
      frequency,
    );
  }
  if (
    isNaN(lowerPrice) ||
    isNaN(upperPrice) ||
    lowerPrice < targetPrice ||
    upperPrice > targetPrice
  ) {
    return NaN;
  }

  let guess = min<f64>(max<f64>(rate, lower + 1e-8), upper - 1e-8);
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const estimatedPrice = couponPriceFromMetricsValue(
      periodsRemaining,
      accruedDays,
      daysToNextCoupon,
      daysInPeriod,
      rate,
      guess,
      redemption,
      frequency,
    );
    if (isNaN(estimatedPrice)) {
      return NaN;
    }
    const error = estimatedPrice - targetPrice;
    if (Math.abs(error) < 1e-12) {
      return guess;
    }
    const epsilon = max<f64>(1e-7, Math.abs(guess) * 1e-6);
    const shiftedPrice = couponPriceFromMetricsValue(
      periodsRemaining,
      accruedDays,
      daysToNextCoupon,
      daysInPeriod,
      rate,
      guess + epsilon,
      redemption,
      frequency,
    );
    const derivative = isNaN(shiftedPrice) ? NaN : (shiftedPrice - estimatedPrice) / epsilon;
    let nextGuess =
      !isFinite(derivative) || derivative == 0.0
        ? (lower + upper) / 2.0
        : guess - error / derivative;
    if (!isFinite(nextGuess) || nextGuess <= lower || nextGuess >= upper) {
      nextGuess = (lower + upper) / 2.0;
    }
    const boundedPrice = couponPriceFromMetricsValue(
      periodsRemaining,
      accruedDays,
      daysToNextCoupon,
      daysInPeriod,
      rate,
      nextGuess,
      redemption,
      frequency,
    );
    if (isNaN(boundedPrice)) {
      return NaN;
    }
    if (boundedPrice > targetPrice) {
      lower = nextGuess;
    } else {
      upper = nextGuess;
    }
    guess = nextGuess;
    if (Math.abs(upper - lower) < 1e-12) {
      return guess;
    }
  }
  return guess;
}

export function macaulayDurationValue(
  periodsRemaining: i32,
  accruedDays: f64,
  daysToNextCoupon: f64,
  daysInPeriod: f64,
  couponRate: f64,
  yieldRate: f64,
  frequency: i32,
): f64 {
  const price = couponPriceFromMetricsValue(
    periodsRemaining,
    accruedDays,
    daysToNextCoupon,
    daysInPeriod,
    couponRate,
    yieldRate,
    100.0,
    frequency,
  );
  if (isNaN(price) || price <= 0.0) {
    return NaN;
  }
  const discountBase = 1.0 + yieldRate / <f64>frequency;
  if (discountBase <= 0.0) {
    return NaN;
  }
  const coupon = (100.0 * couponRate) / <f64>frequency;
  const periodsToNextCoupon = daysToNextCoupon / daysInPeriod;
  let weightedPresentValue = 0.0;
  for (let period = 1; period <= periodsRemaining; period += 1) {
    const periodsToCashflow = <f64>(period - 1) + periodsToNextCoupon;
    const timeInYears = periodsToCashflow / <f64>frequency;
    const cashflow = period == periodsRemaining ? 100.0 + coupon : coupon;
    weightedPresentValue += (timeInYears * cashflow) / Math.pow(discountBase, periodsToCashflow);
  }
  return weightedPresentValue / price;
}

export function excelIsoWeeknumValue(whole: i32): i32 {
  if (whole < 0) {
    return i32.MIN_VALUE;
  }
  const weekday = excelWeekdayFromSerial(<u8>ValueTag.Number, <f64>whole, 2);
  if (weekday == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  const shiftedWhole = adjustedWhole + 4 - weekday;
  const shiftedDays = EXCEL_EPOCH_DAYS + shiftedWhole;
  const shiftedYear = civilYear(shiftedDays);
  const yearStart = daysFromCivil(shiftedYear, 1, 1);
  const dayOfYear = shiftedDays - yearStart + 1;
  return <i32>Math.floor(<f64>(dayOfYear - 1) / 7.0) + 1;
}

function fixedDecliningBalanceRate(cost: f64, salvage: f64, life: f64): f64 {
  if (
    !isFinite(cost) ||
    !isFinite(salvage) ||
    !isFinite(life) ||
    cost <= 0 ||
    salvage < 0 ||
    life <= 0
  ) {
    return NaN;
  }
  const ratio = salvage / cost;
  if (ratio < 0) {
    return NaN;
  }
  return Math.round((1.0 - Math.pow(ratio, 1.0 / life)) * 1000.0) / 1000.0;
}

export function dbDepreciation(cost: f64, salvage: f64, life: f64, period: f64, month: f64): f64 {
  const rate = fixedDecliningBalanceRate(cost, salvage, life);
  if (isNaN(rate) || month < 1 || month > 12 || period < 1 || period > life + 1) {
    return NaN;
  }

  let bookValue = cost;
  let depreciation = 0.0;
  for (let currentPeriod = 1.0; currentPeriod <= period; currentPeriod += 1.0) {
    const raw =
      currentPeriod == 1.0
        ? bookValue * rate * (month / 12.0)
        : currentPeriod == <f64>Math.floor(life) + 1.0
          ? bookValue * rate * ((12.0 - month) / 12.0)
          : bookValue * rate;
    depreciation = Math.min(Math.max(raw, 0.0), Math.max(0.0, bookValue - salvage));
    bookValue -= depreciation;
  }
  return depreciation;
}

function ddbPeriodDepreciation(
  bookValue: f64,
  salvage: f64,
  life: f64,
  factor: f64,
  remainingLife: f64,
  noSwitch: bool,
): f64 {
  const declining = (bookValue * factor) / life;
  const straightLine = remainingLife <= 0 ? 0.0 : (bookValue - salvage) / remainingLife;
  const base = noSwitch ? declining : Math.max(declining, straightLine);
  return Math.min(Math.max(base, 0.0), Math.max(0.0, bookValue - salvage));
}

export function ddbDepreciation(cost: f64, salvage: f64, life: f64, period: f64, factor: f64): f64 {
  if (
    !isFinite(cost) ||
    !isFinite(salvage) ||
    !isFinite(life) ||
    !isFinite(period) ||
    !isFinite(factor) ||
    cost <= 0 ||
    salvage < 0 ||
    life <= 0 ||
    period <= 0 ||
    factor <= 0
  ) {
    return NaN;
  }
  let bookValue = cost;
  let current = 0.0;
  let depreciation = 0.0;
  while (current < period && bookValue > salvage) {
    const segment = Math.min(1.0, period - current);
    const full = Math.min(
      Math.max((bookValue * factor) / life, 0.0),
      Math.max(0.0, bookValue - salvage),
    );
    depreciation = Math.min(full * segment, Math.max(0.0, bookValue - salvage));
    bookValue -= depreciation;
    current += segment;
  }
  return depreciation;
}

export function vdbDepreciation(
  cost: f64,
  salvage: f64,
  life: f64,
  startPeriod: f64,
  endPeriod: f64,
  factor: f64,
  noSwitch: bool,
): f64 {
  if (
    !isFinite(cost) ||
    !isFinite(salvage) ||
    !isFinite(life) ||
    !isFinite(startPeriod) ||
    !isFinite(endPeriod) ||
    !isFinite(factor) ||
    cost <= 0 ||
    salvage < 0 ||
    life <= 0 ||
    startPeriod < 0 ||
    endPeriod < startPeriod ||
    factor <= 0
  ) {
    return NaN;
  }

  let bookValue = cost;
  let total = 0.0;
  for (let current = 0.0; current < endPeriod && bookValue > salvage; current += 1.0) {
    const overlap = Math.max(
      0.0,
      Math.min(endPeriod, current + 1.0) - Math.max(startPeriod, current),
    );
    if (overlap <= 0) {
      const full = ddbPeriodDepreciation(
        bookValue,
        salvage,
        life,
        factor,
        life - current,
        noSwitch,
      );
      bookValue -= full;
      continue;
    }
    const full = ddbPeriodDepreciation(bookValue, salvage, life, factor, life - current, noSwitch);
    total += full * overlap;
    bookValue -= full;
  }
  return total;
}
