> Version 0.8.0: see WORKFLOWS-0.8.md for six/eight timestamp fields and complete working-dataset tables.

> Version 0.7.0 changes the user workflow and visible fields. See WORKFLOWS-0.7.md and README.md; older in-sheet examples below describe the retained calculation engine.

# Census Geocoder

Version 0.3.0. CAP Index > Geography > Census Geocoder inserts results into the selected worksheet.

## Workflow

1. Select one rectangular range with column headers and all relevant data rows. For an Excel Table, include its header and all data rows but exclude its totals row.
2. Map Address, City, State and optional ZIP by their exact headers. City/State may be omitted when ZIP is provided. ZIP may be omitted when City/State are provided.
3. Run the tool. It sends only street, city, state, ZIP and generated address request IDs to the US Census Bureau over HTTPS. Incident IDs, offenses and other selected source fields remain local. Use company-approved address data.
4. Review Match Status, Match Type and Review Flag. Save the workbook to retain results.

The following fields are inserted immediately after the rightmost mapped address component, on the same sheet: Full Address; Matched Address; Latitude; Longitude; Match Status; Match Type; Review Flag; Census Benchmark; Geocoded At. Existing fields shift right. Tables receive new ListColumns; source rows and total rows remain intact. Original source values are retained. Dates/timestamps and coordinates remain typed numbers.

Full Address joins populated source components for reading and comparison. The actual Census batch CSV retains the required separate component fields and blank ZIP when absent. ZIP must be text to preserve leading zeros. Numeric ZIPs, source cell errors, missing street/locality, PO boxes, masked hundred-block addresses and common intersection forms are flagged locally and not submitted. Ordinary address strings, including numeric street names, are sent to Census for its match decision; this does not depend on the hundred-block parser.

## Batch handling

Only distinct eligible component combinations are submitted. The deduplication key trims and normalizes case/repeated spaces in street, city and state; ZIP remains part of the key. Street spelling aliases and missing-versus-provided ZIPs are not inferred equivalent. Every original incident/offense record remains in the worksheet and receives its address result.

Batches contain at most 10,000 addresses and 4,900,000 UTF-8 CSV bytes, conservatively below the service's 5 MB file limit. CSV fields are quoted and embedded quotes escaped; no source header or workbook is sent. Multipart requests specify the Census public current address-range benchmark. Returned rows join by generated request ID, never response order. Unknown/duplicate/missing IDs, malformed CSV, invalid coordinates or unknown status/type reject the batch instead of silently applying partial or shifted responses.

Progress appears in Excel's status bar. Worksheet interaction is temporarily suspended during processing to prevent source rows changing while results are being joined. Esc/Ctrl+Break can interrupt VBA processing. Excel's previous event, screen, interaction, cancellation and status-bar settings are restored on completion or error.

The transport uses Windows WinHTTP and ADODB.Stream, both late bound. HTTPS certificate verification remains enabled. Network/timeouts, HTTP 429 and server errors receive up to three attempts with short waits; non-retryable HTTP errors stop immediately. A batch has a 240-second client wait limit; Census processing time is not guaranteed.

Each fully validated completed batch is written to the worksheet before submitting the next batch. A later failure/cancellation retains completed results and flags unfinished rows as Batch Failed or Cancelled. It does not convert service/transport failures into No Match or zero coordinates. A rerun refreshes the same generated columns and performs a fresh submission of all eligible distinct addresses; it does not implement persistent cross-run caching or resume. The prior version remains available for rollback.

## Result interpretation

- Match + Exact: the address matched the Census locator's address range exactly. Review Flag is blank; this is not a numeric confidence score or a guarantee of rooftop accuracy.
- Match + Non-Exact: coordinates and Matched Address are returned with a review flag.
- Tie: unresolved ambiguity, blank coordinates and a review flag.
- No Match: no locator match, blank coordinates and a review flag.
- Invalid Input: not submitted; the reason is recorded in Review Flag.
- Batch Failed / Cancelled: no valid result applied to that unfinished row; coordinates remain blank.

The Census service interpolates positions along address ranges. Exact matching does not establish a building's existence or a surveyed incident location. The public current benchmark changes as Census updates its underlying data. The output records its alias and local run timestamp; it does not archive the underlying locator. Puerto Rico urbanization-specific fields and external geographic boundary validation are not implemented. Counties/tracts/Census blocks are a future option, separate from street hundred blocks.

## Verification and deployment

Native tests cover CSV escaping/UTF-8 bytes, no-match spelling, coordinate order, request-ID joins with shuffled responses, local invalid-address handling, repeated addresses mapped to multiple source rows, source preservation, worksheet/table insertion and refresh, and completed-batch retention after a forced later failure. A separate 10,001-address synthetic test verifies the 10,000 + 1 split. The optional live smoke test uses two public demonstration addresses only; no crime data is submitted during building.

Version 0.3.0 remains an unsigned pilot. Disable the prior add-in before enabling this version. Company signing, production network policy and visual acceptance remain pending.

Sources checked October 13, 2026: [Census API](https://geocoding.geo.census.gov/geocoder/Geocoding_Services_API.html), [Census upload limits](https://geocoding.geo.census.gov/geocoder/locations/address), [Census user guide](https://www2.census.gov/geo/pdfs/maps-data/data/Census_Geocoder_User_Guide.pdf), [accuracy methodology](https://www.census.gov/programs-surveys/geography/technical-documentation/complete-technical-documentation/census-geocoder.html).
