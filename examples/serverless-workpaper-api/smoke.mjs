import { handleWorkPaperRequest } from './route.mjs'

const updateRecords = [
  { region: 'West', customers: 20, arpa: 1200 },
  { region: 'East', customers: 30, arpa: 250 },
  { region: 'Central', customers: 18, arpa: 300 },
  { region: 'North', customers: 65, arpa: 180 },
]

const before = await requestJson('/api/workpaper/summary')
const edit = await requestJson('/api/workpaper/revenue', {
  method: 'POST',
  headers: {
    'content-type': 'application/json',
  },
  body: JSON.stringify({ records: updateRecords }),
})
const after = await requestJson('/api/workpaper/summary')

const output = {
  before: before.summary,
  edit: {
    records: edit.records,
    after: edit.after,
    checks: edit.checks,
  },
  after: after.summary,
  verified: true,
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

async function requestJson(path, init) {
  const response = await handleWorkPaperRequest(new Request(`http://localhost:8787${path}`, init))
  const body = await response.json()
  if (!response.ok) {
    throw new Error(`request failed: ${response.status} ${JSON.stringify(body)}`)
  }
  return body
}

function assertOutput(actual) {
  const expectedBefore = {
    totalRevenue: 36900,
    westCustomers: 20,
    largestDeal: 24000,
  }
  const expectedAfter = {
    totalRevenue: 48600,
    westCustomers: 20,
    largestDeal: 24000,
  }

  if (
    JSON.stringify(actual.before) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(actual.edit.after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(actual.after) !== JSON.stringify(expectedAfter) ||
    actual.edit.records !== 4 ||
    actual.edit.checks.totalRevenueChanged !== true ||
    actual.edit.checks.formulasPersisted !== true ||
    actual.edit.checks.serializedBytes <= 0
  ) {
    throw new Error(`unexpected WorkPaper API result: ${JSON.stringify(actual)}`)
  }
}
