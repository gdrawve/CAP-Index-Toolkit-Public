# Working datasets and incident review — 0.8.0

## Duplicate / Case Review

Map the stable IncidentID and crime-type/offense-description field. CAP_IncidentCrimeTypes lists distinct nonblank crime descriptions for that incident in alphabetical order on every associated source offense row. Case and repeated whitespace are normalized for uniqueness; the first recorded spelling is displayed. Multiple offense records remain separate; an incident with several crime types is not automatically an erroneous duplicate. Exact-row comparisons and incident-row counts retain their existing definitions.

Blank/error IDs are not pooled into a false incident. If no usable crime description exists for an incident, the list says No recorded crime type. If the optional crime field is not mapped, it says so. Missing/error descriptions are omitted, not invented. Extremely long lists are visibly shortened to fit Excel; all original offense rows remain in the report. Incident IDs must already be unique across agency/year when needed.

## Timestamp Fields

The parser first checks all selected source values. When no valid source value contains time, only Month Number, Month Text, Day Number, Day of Week, Year and Date Only are generated. Time Only and Hour are omitted. If any valid time exists, both time columns are generated; date-only rows within that mixed source remain blank in those two columns. Explicit text midnight is a known time and produces Hour 0. Integer numeric dates still lack time provenance. All-blank/all-invalid source data generates no new columns.

A rerun adjusts the owned timestamp block from six to eight fields, or eight to six, as source content changes. Only marked fields with expected headers are resized. Unrelated adjacent fields and table totals remain. For ordinary worksheets, removal is refused when the time columns contain data above/below the owned output range. Separate Time columns continue to use Hour from Time.

## Geography sheets are working datasets

Hundred_Blocks, Randomized_Addresses and Geocoded_Addresses now contain native Excel Tables starting at row 6, retaining all original dataset fields and rows. Selecting some columns within an existing table expands to that entire table, excluding totals. In a plain worksheet, a selection beginning at the contiguous dataset header expands to that whole contiguous region. For a protected plain worksheet, all adjoining named fields are copied across the selected rows; select all desired record rows. Excel does not permit CurrentRegion on a protected sheet, so unrelated rows are not guessed. A dataset split by blank rows/columns may need the full rectangle explicitly selected; the tool does not guess unrelated regions.

The original sheet is unchanged. Generated fields retain the compact 0.7 definitions and existing matching generated headers are reused on new outputs. Hundred Blocks adds CAP_HundredBlock and Review Status; Randomization adds its synthetic address and clear estimate status; Geocoding adds its seven address/coordinate/match results. IncidentID, Crime Type, Date, Time and all other original fields remain available to subsequent analysis. Source column formats are retained; mixed-format columns use the first data row's format. Source formulas are copied as current values.

The output tables allow a Date field to be selected by clicking a cell when using Timestamp Fields. For other tools, select the table including its header/data. All working-dataset records, including hidden/filtered rows, are copied. Geography tools do not use a partial selection as an analytical time/crime filter. Select an explicitly separated working dataset if limiting preparation scope.

Optional ID/city/state and seed semantics, Census batching/review warnings, analytical definitions and privacy limitations remain as documented in WORKFLOWS-0.7.md. Hidden snapshots and synthetic addresses do not provide anonymization or recover true incident locations.
