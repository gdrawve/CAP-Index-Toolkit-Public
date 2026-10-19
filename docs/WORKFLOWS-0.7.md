# Revised workflows — version 0.7.0

## Geography outputs

Hundred_Blocks, Randomized_Addresses and Geocoded_Addresses use unique numbered sheet names. Titles, source range, version, run notes and seed/benchmark precede the table. The table starts at row 6. Existing matching generated fields are reused in the new sheet when chaining/rerunning a tool; duplicated generated headers are not added. Original source fields are retained for row alignment, incident identifiers and follow-on analysis. Only new generated fields are minimized; pre-existing columns in the selected source are retained.

Hundred Blocks adds CAP_HundredBlock and Review Status. No separate block number, key, precision, seed or synthetic fields are generated. Randomized Addresses adds CAP_HundredBlock, CAP_SyntheticAddress and Review Status. Masked 00/XX/explicit-block inputs generate inclusive base…base+99 values; exact addresses are retained. Review Status identifies synthetic estimates. The seed is recorded once in run notes. Optional incident ID/city/state supports repeatability. Apartment suffixes are removed for hundred-block aggregation but retained for unchanged exact-address output.

Census output adds Full Address, Matched Address, Latitude, Longitude, Match Status, Match Type and Review Flag. Benchmark and run time are recorded in the sheet notes instead of repeated columns. Completed batches remain on that sheet if a later batch fails or is cancelled. Original source data is unchanged. Synthetic-address warnings persist. Census supplies match type and matched address, not a numeric confidence score. Existing 10,000-distinct-address/byte-size splitting and minimal-data network submission remain in force. A protected source sheet is acceptable; writable workbook and unprotected workbook structure are required to add outputs.

## Separate date and time

Timestamp Fields works on a Date or combined timestamp column and retains the requested eight outputs. Date-only inputs leave its Time/Hour fields blank. Hour from Time is a separate command: select exactly one Time column including header and data. It inserts or refreshes one Hour (0–23) column beside Time. If another Hour (0–23) field already exists (for example from Timestamp Fields), the separate-time output is named Hour from Time (0–23) to avoid duplicate headers. The heat-map day/hour prompt prefers that field when present. Numeric Excel time 0 represents explicit midnight in this time-column context. Blank, invalid time, or date-only values stay blank. Text times accept 1:25 PM, 12:00 AM and 23:59:59. These in-sheet tools require an unprotected worksheet.

## Heat-map selection

An initial input-mode prompt precedes range selection:
1. Date/timestamp plus optional explicit Hour.
2. Separate Date and Time.
3. Existing Day of Week text and Hour 0–23.
4. Existing Day of Week text and Time.

Include Crime Type and optional stable IncidentID. The report is Monday–Sunday by 0–23, regardless of mode. Valid zero is midnight; empty hours never become zero. Missing/invalid date or day, missing/invalid time/hour and unsuitable IDs are excluded and counted. Separate time that contradicts a timestamp's hour is excluded. Day names are English full or three-letter names. Numeric weekday conventions are not assumed. In day/hour mode no actual dates exist in the source mapping; ID consistency can verify selected day/hour, not actual calendar dates.

## Repeat Locations

The ribbon tool now ranks addresses without requiring dates or times. Include Address or qualified Full Address, Crime Type and optional IncidentID/City/State. Apartment/unit strings remain part of the address. Case and repeated whitespace normalize; address aliases are not resolved. The output has only Address, Event Count and Additional Events, with an adjustable minimum count (default two), crime/measure selectors and chart. Unique incident IDs with conflicting locations are excluded. Dates do not determine identity for this tool. Earlier consecutive-gap calculations remain internal regression code, outside the ribbon workflow.

## Hot Streets and concentration

Hot Streets maps Hundred Block (or a stable street ID), Crime Type and Date. Date-only values work; time is unnecessary. The existing weeks/months, interval width, full source span, zero-record bins and descriptive classification controls remain. Dates are essential to analyze change. Crime Concentration needs Hundred Block and Crime Type without dates. Both construct qualified geography internally from optional City/State; no visible Block Key column is required. When combining jurisdictions, provide geographic qualification. Missing city/state components in explicitly mapped qualifier fields exclude those location records rather than merging them.

The classification thresholds and coverage limitations in TIME-LOCATION-ANALYSIS.md still apply to Hot Streets. Reports store hidden snapshots containing identifiers/geography; hidden sheets do not anonymize data. Use an approved summary when sharing.
