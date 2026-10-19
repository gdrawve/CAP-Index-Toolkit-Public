# Timestamp Fields and Hour from Time

Ribbon: CAP Index > Date & Time. Both tools insert generated fields beside the selected source column.

## Select and run

Click one cell within the source timestamp column, then click Timestamp Fields. A plain worksheet uses row 1 as the header and processes through the last populated source cell. An Excel Table uses that table column's header and all data rows, excluding its totals row. For a plain worksheet with a header below row 1, select the header and data as a one-column range before clicking the tool.

The tool inserts six date fields when no valid source value includes time, or eight fields when at least one valid time exists. All-blank/all-invalid data creates no columns. On a plain worksheet, existing columns shift right; Excel Tables receive generated ListColumns. Original source values/formats and adjacent values/formulas are retained. All source rows in the selected range, including hidden/filtered rows, are processed. Outputs are values from the run, not automatically refreshing formulas.

| Order | Header | Example for 10/13/2026 14:35:00 |
|---|---|---|
| 1 | Month Number | 10 |
| 2 | Month Text | October |
| 3 | Day Number | 13 |
| 4 | Day of Week | Tuesday |
| 5 | Year | 2026 |
| 6 | Date Only | 10/13/2026 |
| 7 | Time Only | 14:35:00 |
| 8 | Hour (0-23) | 14 |

Month and weekday names are full English names. Day Number is the day of month, 1–31. Date Only remains a numeric Excel date displayed as `mm/dd/yyyy`. Time Only remains a numeric time fraction displayed as `hh:mm:ss`. Other numeric fields are stored as numbers.

## Accepted inputs

- Native Excel date serials, with or without a fractional time; source workbook's 1900/1904 system is respected.
- ISO dates such as `2026-10-13`, with optional time (`2026-10-13 14:35:00` or `2026-10-13T14:35:00`).
- US month/day/year text such as `10/13/2026`, `3/4/2026 7:05 PM` and `10/13/2026 12:00 AM`. Month/day/year parsing means March 4 for `03/04/2026`.
- A mixture of those input forms in the same column.

Dates without an explicit time leave Time Only and Hour blank. Explicit midnight text produces `00:00:00` and hour 0. A whole-day numeric Excel serial cannot distinguish date-only from actual midnight; this version treats those numeric values as date-only. Numeric serials stored as text, timezone offsets/Z suffixes, fractional-second text and invalid dates are not inferred.

Blank source rows produce blank outputs. Invalid nonblank rows produce blank outputs with light red highlighting. The completion message reports valid, blank and invalid counts. If no valid dates exist, the tool stops before inserting columns. The first supported date is March 1, 1900, avoiding Excel's fictitious 1900 leap day.

## Rerun and constraints

A hidden worksheet-scoped name tracks the source header and generated fields. Rerunning refreshes the marked block and safely adjusts between six and eight fields as source time content changes. Keep the generated headers unchanged; edited headers cause a clear stop. If another field in an Excel Table already uses one of the requested names, the tool stops before changing the table to avoid silently renaming fields or replacing existing data. Multiple timestamp sources in one table will need a future agreed header-prefix convention.

The source must be one unmerged column with a header and 1–200,000 data rows. The workbook must be writable and the worksheet unprotected. New columns inserted during a failed run are removed by error cleanup. The add-in restores Excel's prior screen-updating and event settings.

## Validation

Native Excel tests cover column order, every output value/type, native and text dates in one column, US interpretation, AM/PM, explicit midnight, missing times, invalid/blank inputs, neighboring values/formulas, table rows/totals, rerun refresh, source preservation and the 1904 date system. Ribbon images are resolved natively in VBA. Actual ribbon appearance and user interaction still need visual acceptance in the user's Excel session.

## Hour from Time

For separate date and time columns, use Timestamp Fields on Date and Hour from Time on Time. Select one Time column including its header and data, or a cell in that table column. Hour from Time inserts or refreshes one Hour (0-23) field beside Time. If another Hour field already exists, the generated header distinguishes Hour from Time (0-23).

Numeric Excel time zero means explicit midnight in this time-column context. Numeric times and 24-hour/AM/PM text are accepted; blank, invalid or date-only values remain blank. This tool requires a writable workbook and unprotected source sheet.
