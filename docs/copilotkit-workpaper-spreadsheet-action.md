---
title: CopilotKit WorkPaper spreadsheet action
published: true
description: Use @bilig/headless behind CopilotKit actions so a UI agent can edit one workbook input and show formula readback.
tags: copilotkit, useCopilotAction, spreadsheet, workpaper, typescript
canonical_url: https://proompteng.github.io/bilig/copilotkit-workpaper-spreadsheet-action.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# CopilotKit WorkPaper Spreadsheet Action

CopilotKit actions are a practical boundary for user-facing workbook changes.
The user asks for a forecast or pricing edit, the action changes one WorkPaper
input, and the UI can show exactly which formula-backed values moved.

Keep the workbook behavior in `@bilig/headless`. The CopilotKit layer should
name the action, describe the parameters, and call the checked WorkPaper
function.

## Run the checked adapter

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:framework-adapters
```

The CopilotKit lane exposes action names and the same verified write result:

```json
{
  "actionNames": ["readWorkPaperSummary", "setWorkPaperInputCell"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

## CopilotKit action shape

```ts
useCopilotAction({
  name: 'setWorkPaperInputCell',
  description: 'Set one WorkPaper input and return formula readback.',
  parameters: [
    { name: 'sheetName', type: 'string', required: true },
    { name: 'address', type: 'string', required: true },
    { name: 'value', type: 'number', required: true },
  ],
  handler: setWorkPaperInputCell,
})
```

For production, keep the handler narrow. If the action changes business logic,
return the exact edited cell, previous summary values, new summary values, and
whether the restored WorkPaper matches the post-write state.

## What to copy

- Make one action for reading and one action for writing.
- Use visible parameter descriptions so the agent does not guess the address
  format.
- Show formula readback in the UI after the action completes.
- Log the serialized WorkPaper only where your app already stores workbook
  documents.

Official CopilotKit hook reference:
<https://docs.copilotkit.ai/reference/hooks/useCopilotAction>.

Runnable source:
[`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts).
