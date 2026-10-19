# Crime Data Preparation

Select a rectangular source range including headers and the fields and rows required for the tool. Sixteen preparation and analysis commands use a shared setup window: review the selected range, map headers, choose options, and Run. Tool Guide explains each tool's required inputs. Timestamp Fields and Hour from Time retain their column-selection workflow.

## Dataset Check and Coverage

The tool creates a new review sheet with dataset quality and coverage summaries. Review missing or invalid values, identifiers, geography and date coverage before interpreting analyses. Source date span describes observed records; it does not establish complete reporting by an agency. The tool does not repair or delete source records. Apply the workbook's Analysis Scope when selecting analytical rows.

## Duplicate / Case Review

Map a stable IncidentID and, optionally, the crime-type/offense-description field. The review sheet retains all selected offense rows. CAP_IncidentCrimeTypes lists distinct nonblank crime descriptions for each incident alphabetically. Case and repeated whitespace normalize for uniqueness; the first recorded spelling is displayed. Multiple offenses within an incident are not automatically erroneous duplicates.

Blank/error IDs are not combined into a false incident. Missing/error crime descriptions are omitted. If no usable description exists, the output says No recorded crime type; an unmapped crime field is identified explicitly. Long lists are visibly shortened to fit Excel. Qualify incident IDs across agencies/years when necessary.

Exact-row review compares selected values and types, including formula results. Flags support human review; no rows are automatically deleted. Source-record counts and unique-incident counts are distinct.

## Geography working datasets

[Hundred blocks and randomization](HUNDRED-BLOCKS.md) and [Census Geocoder](CENSUS-GEOCODER.md) create new uniquely named sheets with native Excel Tables. The selected source fields and rows, including headers, are copied as current values; only relevant generated fields are added. Preparation tools include hidden/filtered rows within the selected range and do not expand the selection to unrelated source fields or rows.

Titles, source range, toolkit version and run notes precede the table, which starts at row 6. Reused generated headers avoid duplicate result columns when chaining tools. Source sheets are unchanged. A protected source can be read, but the workbook must be writable and its structure must allow new sheets.

[Timestamp Fields and Hour from Time](TIMESTAMP.md) insert results beside their source column and require an unprotected source sheet.

Hidden source snapshots, generalized addresses and synthetic locations do not anonymize a workbook. Share reviewed summaries or exports when the underlying records are sensitive.
