import { axios } from '@pipedream/platform'

export function cleanBaseUrl(baseUrl) {
  if (typeof baseUrl !== 'string' || baseUrl.trim() === '') {
    throw new Error('Bilig Base URL is required.')
  }

  return baseUrl.trim().replace(/\/+$/, '')
}

export async function verifyForecastReadback($, { baseUrl, sheetName, address, value }) {
  const response = await axios($, {
    method: 'POST',
    url: `${cleanBaseUrl(baseUrl)}/api/workpaper/n8n/forecast`,
    headers: {
      accept: 'application/json',
      'content-type': 'application/json',
    },
    data: {
      sheetName,
      address,
      value,
    },
  })

  if (!response?.verified) {
    throw new Error('Bilig did not return verified formula readback proof.')
  }

  const checks = response.checks ?? {}
  const requiredChecks = ['formulasPersisted', 'restoredMatchesAfter', 'computedOutputChanged']
  const missingChecks = requiredChecks.filter((key) => checks[key] !== true)

  if (missingChecks.length > 0) {
    throw new Error(`Bilig formula proof failed checks: ${missingChecks.join(', ')}.`)
  }

  return response
}
