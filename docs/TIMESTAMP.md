> Version 0.8.0: see WORKFLOWS-0.8.md for six/eight timestamp fields and complete working-dataset tables.

# Timestamp Fields: agreed behavior

Updated October 13, 2026. Ribbon: CAP Index > Date & Time > Timestamp Fields. All custom tools now have native Office icons.

## Select and run

Click one cell within the source timestamp column, then click Timestamp Fields. A plain worksheet uses row 1 as the header and processes through the last populated source cell. An Excel Table uses that table column's header and all data rows, excluding its totals row. For a plain worksheet with a header below row 1, select the header and data as a one-column range before clicking the tool.

The tool inserts exactly eight columns immediately to the right of the source. On a plain worksheet, existing worksheet columns shift right. In an Excel Table, eight ListColumns are inserted within the table. Original source values/formats and adjacent values/formulas are retained. All source rows in the selected range, including hidden/filtered rows, are processed. Outputs are values from the run, not automatically refreshing formulas.

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
- US month/day/year text such as `10/13/2026`, `3/4/2026 7:05 PM` and `10/13/2026 12:00 AM`. The user explicitly chose month/day/year; March 4 is the meaning of `03/04/2026`.
- A mixture of those input forms in the same column.

Dates without an explicit time leave Time Only and Hour blank. Explicit midnight text produces `00:00:00` and hour 0. A whole-day numeric Excel serial cannot distinguish date-only from actual midnight; this version treats those numeric values as date-only. Numeric serials stored as text, timezone offsets/Z suffixes, fractional-second text and invalid dates are not inferred.

Blank source rows produce eight blank outputs. Invalid nonblank rows produce eight blank outputs with light red highlighting. The completion message reports valid, blank and invalid counts. If no valid dates exist, the tool stops before inserting columns. The first supported date is March 1, 1900, avoiding Excel's fictitious 1900 leap day.

## Rerun and constraints

A hidden worksheet-scoped name tracks the source header and its eight generated headers. Rerunning refreshes those generated fields rather than adding another eight columns. Keep the generated headers unchanged; edited headers cause a clear stop. If another field in an Excel Table already uses one of the eight requested names, the tool stops before changing the table to avoid silently renaming fields or replacing existing data. Multiple timestamp sources in one table will need a future agreed header-prefix convention.

The source must be one unmerged column with a header and 1–200,000 data rows. The workbook must be writable and the worksheet unprotected. New columns inserted during a failed run are removed by error cleanup. The add-in restores Excel's prior screen-updating and event settings.

## Validation

Native Excel tests cover column order, every output value/type, native and text dates in one column, US interpretation, AM/PM, explicit midnight, missing times, invalid/blank inputs, neighboring values/formulas, table rows/totals, rerun refresh, source preservation and the 1904 date system. Ribbon images are resolved natively in VBA. Actual ribbon appearance and user interaction still need visual acceptance in the user's Excel session.
