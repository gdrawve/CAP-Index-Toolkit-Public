# Data contract and analytical limits

## Record grain

- A source row is a record. An offense export can have multiple rows per incident. Do not label row counts as crimes or unique offenses until the export's grain is verified.
- IncidentID must uniquely identify an incident across the selected agencies and years. Construct an agency/year/case key in the source when case numbers are reused. Leading zeros belong in text identifiers; already lost zeros cannot be recovered.
- IDs are trimmed and case-sensitive. Case variants are distinct until the source owner approves normalization.
- Exact duplicate review compares every selected column's value and type. It compares formula results rather than formulas. Its flags are candidates for human review, not instructions to delete rows.
- No automatic offense deduplication is implemented. A stable offense-record key and explicit correction/supplement rules are needed for that phase.

## Date and time

Timestamp Fields processes one source column containing mixed native Excel dates/timestamps, ISO text, and US month/day/year text. Text dates use US ordering: `03/04/2026` means March 4. Text can include a 24-hour time or an AM/PM time with optional seconds. Timezone suffixes, invalid calendar dates, malformed strings, and numeric serials stored as text are rejected. For separate time fields, use Hour from Time or a temporal analysis setup mode that accepts separate Date and Time. See TIMESTAMP.md for output placement and refresh behavior.

Numeric date serials use the source workbook's 1900/1904 date system. The first supported 1900-system numeric date is March 1, 1900; earlier serials are rejected to avoid Excel's fictitious leap day. Integer serials are treated as date-only because they cannot distinguish missing time from true midnight. Explicit ISO midnight or a separate time value of zero establishes midnight.

Up to eight derived fields are Month Number, Month Text, Day Number (day of month), Day of Week (full English name), Year, Date Only, Time Only and Hour (0-23). No timezone conversion, daylight-saving inference, occurrence/report-date substitution or interval allocation is performed. Source timezone and whether timestamps mean occurrence or report time must be defined before temporal comparisons. Source values and formats remain unchanged; the inserted fields contain typed numbers/dates/times with appropriate formats. Invalid nonblank rows produce blank generated fields with a light red fill; blank source rows produce unhighlighted blanks. When no valid time exists, Time Only and Hour are omitted; all-blank/all-invalid data produces no new columns. No ninth status column is added.

## Hundred blocks

`123 N MAIN ST` maps to `100 BLOCK N MAIN ST`; `99 MAIN ST` maps to `0 BLOCK MAIN ST`. Blocks use floor(house number/100)*100, with documented unit suffixes removed. The key combines state, city and block using length-prefixed text parts to avoid separator collisions. Original addresses are retained locally.

This is an address-number aggregation, not a Census block, geometrically validated street segment or geocode. Street abbreviation/direction aliases are not merged: MAIN ST and MAIN STREET remain separate until an approved normalization table is added. Municipal spelling differences also remain separate. Map optional City/State for jurisdiction qualification, or provide an already qualified location identifier; otherwise the source must represent one jurisdiction. Intersections, address ranges, fractional or hyphenated house numbers and PO boxes are excluded for review. The block parser recognizes 00/XX endings and explicit BLOCK/BLK labels as hundred-block input. In Randomize Block Addresses, 1200 represents the 1200–1299 range. Ordinary ordinal/numeric street names are supported. Only the selected geographic/date coverage is analyzed.

The Randomize Block Addresses action generates separately labeled synthetic house numbers from block base through base + 99 inclusive. Original inputs and the block key remain intact. A repeatable seed and optional incident ID keep offense rows of the same incident/block together. Without an incident ID, source-row identity is used; sorting can change regenerated values. Generated values are simulated and do not establish a real building, the true incident location or geometric street position. Census review flags retain synthetic/masked precision when the generated fields are mapped. See HUNDRED-BLOCKS.md.

Hundred-block labels reduce precision but do not anonymize the workbook. Source addresses, case IDs and rare-event combinations remain sensitive. Shared reports should omit original rows and follow company aggregation/suppression policy.

## Concentration

The Crime Concentration ribbon action creates the dashboard described in CRIME-CONCENTRATION.md. Its blue crime-type and count-measure selectors update totals, rankings, summaries and three charts. The denominator includes only populated eligible hundred blocks for the selected crime. Missing crime categories and source-field errors have explicit selectable groups; All includes them. Top 1%, 5%, 10% and 20% shares are measured rather than assumed.

RECORDS counts all selected rows with a valid block key, including suspected duplicates and rows without incident IDs. This is an offense-record count only when the input grain establishes it. INCIDENTS counts each valid ID once. An ID with multiple block keys, or with any missing/error block key, is excluded from incident concentration everywhere; no first-record assignment is made. Valid IDs with no usable block are included in excluded-ID coverage counts. Missing IDs cannot contribute to unique-incident counts.

The denominator is activity on observed, populated, eligible blocks, not all blocks in a jurisdiction. Zero-activity blocks require a separate street/block universe that is not supplied. Output reports source rows, missing/error block rows, missing/error ID rows, valid IDs, excluded IDs, eligible incidents, eligible activity, and populated block count.

Ranking is descending count with binary block-key order for ties. Top 20% uses ceiling(0.20 * populated blocks), so small datasets can include more than exactly 20%. The output reports the actual share, the smallest ranked prefix reaching 80% and the number including all ties at that threshold. The 80/20 pattern is tested, never assumed. Results are snapshots; duplicates, missing locations, changing agency/reporting coverage and alias fragmentation can change measured concentration. No crime-rate, causal or statistical significance inference is provided.

## Census geocoding

Census Geocoder creates a new working dataset with selected source fields plus Full Address, Matched Address, Latitude, Longitude, Match Status, Match Type and Review Flag. Benchmark and run time are recorded in notes. Source sheets are unchanged. See [Census Geocoder](CENSUS-GEOCODER.md).

Distinct eligible component combinations are submitted once per run and mapped back to all corresponding records. CSV fields are quoted and escaped, ZIP strings are preserved, and batches split at 10,000 addresses or 4,900,000 UTF-8 CSV bytes. Incident IDs, offense descriptions, victim information and the workbook are not transmitted. Use company-approved address data. A hundred-block mask changes matchability and accuracy and does not replace authorization for external sharing.

The Census service uses US/territory address-range coverage with documented Puerto Rico variations. This version supports the ordinary five-field format, not Puerto Rico urbanization fields. Unknown/masked addresses, PO boxes, intersections and new construction may not match. Returned coordinates are address-range estimates, not guaranteed rooftop or incident locations. Exact is an address-match category, not a numeric confidence or positional-accuracy measure. Match/no-match/tie/invalid-input/failed-batch outcomes remain distinct. Do not infer jurisdiction from a returned point without boundary validation.

Responses join by generated request ID within each in-memory run, not response row order. Duplicate, unknown or missing IDs reject a batch. Latitude and longitude are separated in the correct order and validated against coordinate bounds. Completed batch results survive later failures; unfinished rows are flagged and have blank coordinates. The completion summary reports matched, non-exact, tie, unmatched, invalid and unfinished source-row counts. A new run creates another working dataset and submits eligible addresses again. Persistent cross-run caching/resume, an archived run ledger, resolved benchmark capture and optional Census geographies remain future work. `Current` changes over time; the current output records its alias and local geocoding timestamp.

## Maps

Maps and 3D Maps retain Excel's built-in behavior. Filled maps generally summarize geographic areas, not individual hundred-block street locations. Use validated coordinates in 3D Maps for point-based work. Connected mapping can transmit location data to map services; use approved data and company connected-experience settings. A map does not establish geocoding accuracy or complete data coverage.

## Temporal and location reports

See TIME-LOCATION-ANALYSIS.md for counting units, calendar bins, incident consistency, missing-time treatment, partial-bin handling and descriptive threshold definitions. No-record intervals represent observed source counts, not verified agency coverage.

## Reporting and export

See REPORTING-EXPORT.md for selected ranges, incident consistency, category overlap, observed-date-span/partial-period flags, zero baselines and static-export behavior. Report counts use occurrence dates; source offense records are not assumed unique offenses.
