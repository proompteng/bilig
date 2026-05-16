import { pathToFileURL } from 'node:url'

import { readRevenueSummaryAction, updateRevenueRecordsFromFormDataAction } from './next-server-action-formdata.ts'

type Summary = {
  largestDeal: number
  totalRevenue: number
  westCustomers: number
}

const formFields = ['region', 'customers', 'arpa'] as const

export async function validateRevenueRecordsFromFormDataAction(formData: FormData) {
  'use server'

  const before = await readRevenueSummaryAction()
  let validationError: string | undefined

  try {
    await updateRevenueRecordsFromFormDataAction(formData)
  } catch (error) {
    validationError = error instanceof Error ? error.message : String(error)
  }

  const after = await readRevenueSummaryAction()
  const output = {
    action: 'Next.js Server Action FormData validation',
    rejectedInput: {
      fields: formFields,
      records: readFormRecordCount(formData),
      shape: readFormShape(formData),
    },
    validationError,
    summaryBefore: before.summary,
    summaryAfter: after.summary,
    unchanged: summariesMatch(before.summary, after.summary),
    verified: true,
  }

  assertOutput(output)
  return output
}

export async function createNextServerActionValidationDemoOutput() {
  return validateRevenueRecordsFromFormDataAction(createInvalidRevenueFormData())
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  console.log(JSON.stringify(await createNextServerActionValidationDemoOutput(), null, 2))
}

function createInvalidRevenueFormData(): FormData {
  const formData = new FormData()
  formData.append('region', 'West')
  formData.append('customers', '-1')
  formData.append('arpa', '1200')
  return formData
}

function readFormRecordCount(formData: FormData): number {
  return Math.max(0, ...formFields.map((field) => formData.getAll(field).length))
}

function readFormShape(formData: FormData): Record<(typeof formFields)[number], string[]> {
  return {
    region: readStringEntries(formData, 'region'),
    customers: readStringEntries(formData, 'customers'),
    arpa: readStringEntries(formData, 'arpa'),
  }
}

function readStringEntries(formData: FormData, name: (typeof formFields)[number]): string[] {
  return formData.getAll(name).map((value) => (typeof value === 'string' ? value : '[file]'))
}

function summariesMatch(left: Summary, right: Summary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertOutput(actual: Awaited<ReturnType<typeof validateRevenueRecordsFromFormDataAction>>): void {
  const expectedSummary = {
    largestDeal: 24000,
    totalRevenue: 36900,
    westCustomers: 20,
  }

  if (
    actual.action !== 'Next.js Server Action FormData validation' ||
    JSON.stringify(actual.rejectedInput.fields) !== JSON.stringify(formFields) ||
    actual.rejectedInput.records !== 1 ||
    JSON.stringify(actual.rejectedInput.shape) !==
      JSON.stringify({
        region: ['West'],
        customers: ['-1'],
        arpa: ['1200'],
      }) ||
    actual.validationError !== 'record 1 customers must be a non-negative number' ||
    JSON.stringify(actual.summaryBefore) !== JSON.stringify(expectedSummary) ||
    JSON.stringify(actual.summaryAfter) !== JSON.stringify(expectedSummary) ||
    !actual.unchanged ||
    !actual.verified
  ) {
    throw new Error(`unexpected Next.js Server Action validation output: ${JSON.stringify(actual)}`)
  }
}
