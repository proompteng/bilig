import { cleanBaseUrl, verifyForecastReadback } from './common/workpaper.mjs'

export default {
  type: 'app',
  app: 'bilig',
  propDefinitions: {
    baseUrl: {
      type: 'string',
      label: 'Bilig Base URL',
      description: 'Base URL for the Bilig app or hosted demo endpoint.',
      default: 'https://bilig.proompteng.ai',
    },
    forecastSheetName: {
      type: 'string',
      label: 'Sheet Name',
      description: 'Forecast input sheet to edit.',
      default: 'Inputs',
    },
    forecastCell: {
      type: 'string',
      label: 'Cell',
      description: 'Editable forecast input cell.',
      options: [
        { label: 'B2 Qualified Opportunities', value: 'B2' },
        { label: 'B3 Win Rate', value: 'B3' },
        { label: 'B4 Average ARR', value: 'B4' },
        { label: 'B5 Expansion Multiplier', value: 'B5' },
      ],
      default: 'B3',
    },
    forecastValue: {
      type: 'integer',
      label: 'Value',
      description: 'Numeric value to write before formula readback.',
      default: 40,
    },
    forecastValueDivisor: {
      type: 'integer',
      label: 'Value Divisor',
      description: 'Divide Value by this number before sending. Use 100 for percentages like 40 -> 0.4.',
      default: 100,
      min: 1,
    },
  },
  methods: {
    cleanBaseUrl,
    verifyForecastReadback,
  },
}
