# Formula Oracle Capture

## Source of truth

Oracle captures come from Excel for the web, not desktop Excel.

## Capture payload

Every captured fixture must include:

- formula text
- input cells
- expected outputs
- visible error string where relevant
- capture timestamp
- Excel build identifier
- notes for coercion, spill, or metadata-sensitive behavior

## Capture rules

- captures are checked into `@bilig/excel-fixtures`
- CI never depends on live Microsoft services
- volatile captures must record the sampled value and capture timestamp explicitly
- metadata-sensitive captures must declare required workbook context such as names or tables

## Exit gate

- every Top 100 entry links to a checked-in oracle case or an explicit blocker
