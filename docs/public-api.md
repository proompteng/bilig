# Public API

## React workbook DSL

```tsx
<Workbook>
  <Sheet name="Sheet1">
    <Cell addr="A1" value={10} />
    <Cell addr="B1" formula="A1*2" />
  </Sheet>
</Workbook>
```

## Imperative engine

- `createSheet(name)`
- `deleteSheet(name)`
- `setCellValue(sheet, address, value)`
- `setCellFormula(sheet, address, formula)`
- `clearCell(sheet, address)`
- `getCell(sheet, address)`
- `getDependencies(sheet, address)`
- `exportSnapshot()`
- `importSnapshot(snapshot)`
- `subscribe(listener)`
