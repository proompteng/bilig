import type { CellValue } from "@bilig/protocol";
import {
  cumulativePeriodicPayment,
  dbDepreciation,
  ddbDepreciation,
  futureValue,
  interestPayment,
  periodicPayment,
  presentValue,
  principalPayment,
  solveRate,
  totalPeriods,
  vdbDepreciation,
} from "./financial.js";
import type { EvaluationResult } from "../runtime-values.js";

type Builtin = (...args: CellValue[]) => EvaluationResult;

interface FinancialBuiltinDeps {
  toNumber: (value: CellValue) => number | undefined;
  coerceBoolean: (value: CellValue | undefined, fallback: boolean) => boolean | undefined;
  coerceNumber: (value: CellValue | undefined, fallback: number) => number | undefined;
  coercePaymentType: (value: CellValue | undefined, fallback: number) => number | undefined;
  integerValue: (value: CellValue | undefined, fallback?: number) => number | undefined;
  positiveIntegerValue: (value: CellValue | undefined, fallback?: number) => number | undefined;
  numberResult: (value: number) => EvaluationResult;
  numericResultOrError: (value: number) => EvaluationResult;
  valueError: () => EvaluationResult;
}

export function createFinancialBuiltins({
  toNumber,
  coerceBoolean,
  coerceNumber,
  coercePaymentType,
  integerValue,
  positiveIntegerValue,
  numberResult,
  numericResultOrError,
  valueError,
}: FinancialBuiltinDeps): Record<string, Builtin> {
  return {
    EFFECT: (nominalRateArg, periodsArg) => {
      const nominalRate = toNumber(nominalRateArg);
      const periods = positiveIntegerValue(periodsArg);
      if (nominalRate === undefined || periods === undefined) {
        return valueError();
      }
      return numberResult((1 + nominalRate / periods) ** periods - 1);
    },
    NOMINAL: (effectiveRateArg, periodsArg) => {
      const effectiveRate = toNumber(effectiveRateArg);
      const periods = positiveIntegerValue(periodsArg);
      if (effectiveRate === undefined || periods === undefined || effectiveRate <= -1) {
        return valueError();
      }
      return numberResult(periods * ((1 + effectiveRate) ** (1 / periods) - 1));
    },
    PDURATION: (rateArg, presentArg, futureArg) => {
      const rate = toNumber(rateArg);
      const present = toNumber(presentArg);
      const future = toNumber(futureArg);
      if (
        rate === undefined ||
        present === undefined ||
        future === undefined ||
        rate <= 0 ||
        present <= 0 ||
        future <= 0
      ) {
        return valueError();
      }
      return numberResult(Math.log(future / present) / Math.log(1 + rate));
    },
    RRI: (periodsArg, presentArg, futureArg) => {
      const periods = toNumber(periodsArg);
      const present = toNumber(presentArg);
      const future = toNumber(futureArg);
      if (
        periods === undefined ||
        present === undefined ||
        future === undefined ||
        periods <= 0 ||
        present === 0
      ) {
        return valueError();
      }
      return numericResultOrError((future / present) ** (1 / periods) - 1);
    },
    FV: (rateArg, periodsArg, paymentArg, presentArg, typeArg) => {
      const rate = toNumber(rateArg);
      const periods = toNumber(periodsArg);
      const payment = toNumber(paymentArg);
      const present = coerceNumber(presentArg, 0);
      const type = coercePaymentType(typeArg, 0);
      if (
        rate === undefined ||
        periods === undefined ||
        payment === undefined ||
        present === undefined ||
        type === undefined
      ) {
        return valueError();
      }
      return numberResult(futureValue(rate, periods, payment, present, type));
    },
    FVSCHEDULE: (principalArg, ...scheduleArgs) => {
      const principal = toNumber(principalArg);
      if (principal === undefined) {
        return valueError();
      }
      let result = principal;
      for (const scheduleArg of scheduleArgs) {
        const rate = toNumber(scheduleArg);
        if (rate === undefined) {
          return valueError();
        }
        result *= 1 + rate;
      }
      return numberResult(result);
    },
    DB: (costArg, salvageArg, lifeArg, periodArg, monthArg) => {
      const cost = toNumber(costArg);
      const salvage = toNumber(salvageArg);
      const life = toNumber(lifeArg);
      const period = toNumber(periodArg);
      const month = coerceNumber(monthArg, 12);
      if (
        cost === undefined ||
        salvage === undefined ||
        life === undefined ||
        period === undefined ||
        month === undefined
      ) {
        return valueError();
      }
      const depreciation = dbDepreciation(cost, salvage, life, period, month);
      return depreciation === undefined ? valueError() : numberResult(depreciation);
    },
    DDB: (costArg, salvageArg, lifeArg, periodArg, factorArg) => {
      const cost = toNumber(costArg);
      const salvage = toNumber(salvageArg);
      const life = toNumber(lifeArg);
      const period = toNumber(periodArg);
      const factor = coerceNumber(factorArg, 2);
      if (
        cost === undefined ||
        salvage === undefined ||
        life === undefined ||
        period === undefined ||
        factor === undefined
      ) {
        return valueError();
      }
      const depreciation = ddbDepreciation(cost, salvage, life, period, factor);
      return depreciation === undefined ? valueError() : numberResult(depreciation);
    },
    VDB: (costArg, salvageArg, lifeArg, startArg, endArg, factorArg, noSwitchArg) => {
      const cost = toNumber(costArg);
      const salvage = toNumber(salvageArg);
      const life = toNumber(lifeArg);
      const start = toNumber(startArg);
      const end = toNumber(endArg);
      const factor = coerceNumber(factorArg, 2);
      const noSwitch = coerceBoolean(noSwitchArg, false);
      if (
        cost === undefined ||
        salvage === undefined ||
        life === undefined ||
        start === undefined ||
        end === undefined ||
        factor === undefined ||
        noSwitch === undefined
      ) {
        return valueError();
      }
      const depreciation = vdbDepreciation(cost, salvage, life, start, end, factor, noSwitch);
      return depreciation === undefined ? valueError() : numberResult(depreciation);
    },
    PV: (rateArg, periodsArg, paymentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg);
      const periods = toNumber(periodsArg);
      const payment = toNumber(paymentArg);
      const future = coerceNumber(futureArg, 0);
      const type = coercePaymentType(typeArg, 0);
      if (
        rate === undefined ||
        periods === undefined ||
        payment === undefined ||
        future === undefined ||
        type === undefined
      ) {
        return valueError();
      }
      return numberResult(presentValue(rate, periods, payment, future, type));
    },
    PMT: (rateArg, periodsArg, presentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg);
      const periods = toNumber(periodsArg);
      const present = toNumber(presentArg);
      const future = coerceNumber(futureArg, 0);
      const type = coercePaymentType(typeArg, 0);
      if (
        rate === undefined ||
        periods === undefined ||
        present === undefined ||
        future === undefined ||
        type === undefined
      ) {
        return valueError();
      }
      const payment = periodicPayment(rate, periods, present, future, type);
      return payment === undefined ? valueError() : numberResult(payment);
    },
    RATE: (periodsArg, paymentArg, presentArg, futureArg, typeArg, guessArg) => {
      const periods = toNumber(periodsArg);
      const payment = toNumber(paymentArg);
      const present = toNumber(presentArg);
      const future = coerceNumber(futureArg, 0);
      const type = coercePaymentType(typeArg, 0);
      const guess = coerceNumber(guessArg, 0.1);
      if (
        periods === undefined ||
        payment === undefined ||
        present === undefined ||
        future === undefined ||
        type === undefined ||
        guess === undefined
      ) {
        return valueError();
      }
      const rate = solveRate(periods, payment, present, future, type, guess);
      return rate === undefined ? valueError() : numberResult(rate);
    },
    SLN: (costArg, salvageArg, lifeArg) => {
      const cost = toNumber(costArg);
      const salvage = toNumber(salvageArg);
      const life = toNumber(lifeArg);
      if (cost === undefined || salvage === undefined || life === undefined || life <= 0) {
        return valueError();
      }
      return numberResult((cost - salvage) / life);
    },
    SYD: (costArg, salvageArg, lifeArg, periodArg) => {
      const cost = toNumber(costArg);
      const salvage = toNumber(salvageArg);
      const life = toNumber(lifeArg);
      const period = toNumber(periodArg);
      if (
        cost === undefined ||
        salvage === undefined ||
        life === undefined ||
        period === undefined ||
        life <= 0 ||
        period <= 0 ||
        period > life
      ) {
        return valueError();
      }
      return numberResult(((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1)));
    },
    NPER: (rateArg, paymentArg, presentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg);
      const payment = toNumber(paymentArg);
      const present = toNumber(presentArg);
      const future = coerceNumber(futureArg, 0);
      const type = coercePaymentType(typeArg, 0);
      if (
        rate === undefined ||
        payment === undefined ||
        present === undefined ||
        future === undefined ||
        type === undefined
      ) {
        return valueError();
      }
      const periods = totalPeriods(rate, payment, present, future, type);
      return periods === undefined ? valueError() : numberResult(periods);
    },
    NPV: (rateArg, ...valueArgs) => {
      const rate = toNumber(rateArg);
      if (rate === undefined || valueArgs.length === 0) {
        return valueError();
      }
      let result = 0;
      for (let index = 0; index < valueArgs.length; index += 1) {
        const value = toNumber(valueArgs[index]!);
        if (value === undefined) {
          return valueError();
        }
        result += value / (1 + rate) ** (index + 1);
      }
      return numberResult(result);
    },
    IPMT: (rateArg, periodArg, periodsArg, presentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg);
      const period = toNumber(periodArg);
      const periods = toNumber(periodsArg);
      const present = toNumber(presentArg);
      const future = coerceNumber(futureArg, 0);
      const type = coercePaymentType(typeArg, 0);
      if (
        rate === undefined ||
        period === undefined ||
        periods === undefined ||
        present === undefined ||
        future === undefined ||
        type === undefined
      ) {
        return valueError();
      }
      const interest = interestPayment(rate, period, periods, present, future, type);
      return interest === undefined ? valueError() : numberResult(interest);
    },
    PPMT: (rateArg, periodArg, periodsArg, presentArg, futureArg, typeArg) => {
      const rate = toNumber(rateArg);
      const period = toNumber(periodArg);
      const periods = toNumber(periodsArg);
      const present = toNumber(presentArg);
      const future = coerceNumber(futureArg, 0);
      const type = coercePaymentType(typeArg, 0);
      if (
        rate === undefined ||
        period === undefined ||
        periods === undefined ||
        present === undefined ||
        future === undefined ||
        type === undefined
      ) {
        return valueError();
      }
      const principal = principalPayment(rate, period, periods, present, future, type);
      return principal === undefined ? valueError() : numberResult(principal);
    },
    ISPMT: (rateArg, periodArg, periodsArg, presentArg) => {
      const rate = toNumber(rateArg);
      const period = toNumber(periodArg);
      const periods = toNumber(periodsArg);
      const present = toNumber(presentArg);
      if (
        rate === undefined ||
        period === undefined ||
        periods === undefined ||
        present === undefined ||
        periods <= 0 ||
        period < 1 ||
        period > periods
      ) {
        return valueError();
      }
      return numberResult(present * rate * (period / periods - 1));
    },
    CUMIPMT: (rateArg, periodsArg, presentArg, startPeriodArg, endPeriodArg, typeArg) => {
      const rate = toNumber(rateArg);
      const periods = toNumber(periodsArg);
      const present = toNumber(presentArg);
      const startPeriod = integerValue(startPeriodArg);
      const endPeriod = integerValue(endPeriodArg);
      const type = coercePaymentType(typeArg, 0);
      if (
        rate === undefined ||
        periods === undefined ||
        present === undefined ||
        startPeriod === undefined ||
        endPeriod === undefined ||
        type === undefined
      ) {
        return valueError();
      }
      const total = cumulativePeriodicPayment(
        rate,
        periods,
        present,
        startPeriod,
        endPeriod,
        type,
        false,
      );
      return total === undefined ? valueError() : numberResult(total);
    },
    CUMPRINC: (rateArg, periodsArg, presentArg, startPeriodArg, endPeriodArg, typeArg) => {
      const rate = toNumber(rateArg);
      const periods = toNumber(periodsArg);
      const present = toNumber(presentArg);
      const startPeriod = integerValue(startPeriodArg);
      const endPeriod = integerValue(endPeriodArg);
      const type = coercePaymentType(typeArg, 0);
      if (
        rate === undefined ||
        periods === undefined ||
        present === undefined ||
        startPeriod === undefined ||
        endPeriod === undefined ||
        type === undefined
      ) {
        return valueError();
      }
      const total = cumulativePeriodicPayment(
        rate,
        periods,
        present,
        startPeriod,
        endPeriod,
        type,
        true,
      );
      return total === undefined ? valueError() : numberResult(total);
    },
  };
}
