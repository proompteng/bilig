import bilig from '../../bilig.app.mjs'
import { verifyForecastReadback } from '../../common/workpaper.mjs'

export default {
  key: 'bilig-verify-formula-readback',
  name: 'Verify Formula Readback',
  description: 'Write a Bilig WorkPaper forecast input cell and return verified recalculated formula output.',
  type: 'action',
  version: '0.0.1',
  annotations: {
    destructiveHint: false,
    openWorldHint: true,
    readOnlyHint: false,
  },
  props: {
    bilig,
    baseUrl: {
      propDefinition: [bilig, 'baseUrl'],
    },
    sheetName: {
      propDefinition: [bilig, 'forecastSheetName'],
    },
    address: {
      propDefinition: [bilig, 'forecastCell'],
    },
    value: {
      propDefinition: [bilig, 'forecastValue'],
    },
    valueDivisor: {
      propDefinition: [bilig, 'forecastValueDivisor'],
    },
  },
  async run({ $ }) {
    const numericValue = Number(this.value) / Number(this.valueDivisor)

    if (!Number.isFinite(numericValue)) {
      throw new Error('Value divided by Value Divisor must produce a finite number.')
    }

    const result = await verifyForecastReadback($, {
      baseUrl: this.baseUrl,
      sheetName: this.sheetName,
      address: this.address,
      value: numericValue,
    })

    $.export('$summary', `Verified ${result.editedCell}: expected ARR ${result.before?.expectedArr} -> ${result.after?.expectedArr}`)

    return result
  },
}
