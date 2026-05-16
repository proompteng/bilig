import { z } from 'zod'

export const safeNonNegativeIntegerSchema = z
  .number()
  .int()
  .nonnegative()
  .refine((value) => Number.isSafeInteger(value), { message: 'Expected safe integer' })

export const safePositiveIntegerSchema = z
  .number()
  .int()
  .positive()
  .refine((value) => Number.isSafeInteger(value), { message: 'Expected safe integer' })
