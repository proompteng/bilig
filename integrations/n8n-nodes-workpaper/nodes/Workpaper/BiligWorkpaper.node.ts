import { NodeConnectionTypes, type INodeType, type INodeTypeDescription } from 'n8n-workflow'

const forecastResource = {
  resource: ['forecast'],
}

const verifyForecastReadback = {
  resource: ['forecast'],
  operation: ['verifyReadback'],
}

export class BiligWorkpaper implements INodeType {
  description: INodeTypeDescription = {
    displayName: 'Bilig WorkPaper',
    name: 'biligWorkpaper',
    icon: { light: 'file:workpaper.svg', dark: 'file:workpaper.dark.svg' },
    group: ['transform'],
    version: 1,
    subtitle: '={{$parameter["operation"]}}',
    description: 'Verify spreadsheet formula readback with Bilig WorkPaper',
    defaults: {
      name: 'Bilig WorkPaper',
    },
    usableAsTool: true,
    inputs: [NodeConnectionTypes.Main],
    outputs: [NodeConnectionTypes.Main],
    credentials: [],
    requestDefaults: {
      baseURL: '={{$parameter["baseUrl"].replace(/\\/$/, "")}}',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
    },
    properties: [
      {
        displayName: 'Resource',
        name: 'resource',
        type: 'options',
        noDataExpression: true,
        options: [
          {
            name: 'Forecast',
            value: 'forecast',
          },
        ],
        default: 'forecast',
      },
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: forecastResource,
        },
        options: [
          {
            name: 'Verify Formula Readback',
            value: 'verifyReadback',
            action: 'Verify formula readback',
            description: 'Edit one forecast input cell and return recalculated formula proof',
            routing: {
              request: {
                method: 'POST',
                url: '/api/workpaper/n8n/forecast',
              },
            },
          },
        ],
        default: 'verifyReadback',
      },
      {
        displayName: 'Bilig Base URL',
        name: 'baseUrl',
        type: 'string',
        default: 'https://bilig.proompteng.ai',
        required: true,
        description: 'Base URL for the Bilig app or hosted demo endpoint',
      },
      {
        displayName: 'Sheet Name',
        name: 'sheetName',
        type: 'string',
        default: 'Inputs',
        required: true,
        displayOptions: {
          show: verifyForecastReadback,
        },
        description: 'Forecast input sheet to edit',
        routing: {
          send: {
            type: 'body',
            property: 'sheetName',
          },
        },
      },
      {
        displayName: 'Cell',
        name: 'address',
        type: 'options',
        default: 'B3',
        required: true,
        displayOptions: {
          show: verifyForecastReadback,
        },
        options: [
          {
            name: 'B2 Qualified Opportunities',
            value: 'B2',
          },
          {
            name: 'B3 Win Rate',
            value: 'B3',
          },
          {
            name: 'B4 Average ARR',
            value: 'B4',
          },
          {
            name: 'B5 Expansion Multiplier',
            value: 'B5',
          },
        ],
        description: 'Editable forecast input cell',
        routing: {
          send: {
            type: 'body',
            property: 'address',
          },
        },
      },
      {
        displayName: 'Value',
        name: 'value',
        type: 'number',
        default: 0.4,
        required: true,
        displayOptions: {
          show: verifyForecastReadback,
        },
        description: 'Numeric value to write before formula readback',
        routing: {
          send: {
            type: 'body',
            property: 'value',
          },
        },
      },
    ],
  }
}
