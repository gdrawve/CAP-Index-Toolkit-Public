# Build plan

## Foundation implemented

1. Separate RibbonX layout and VBA command routing. Six CAP Index groups; native Maps and 3D Maps controls.
2. Shared range selection, header validation, source-preserving output sheets and run provenance.
3. Timestamp Fields v0.2: eight adjacent fields, mixed native/ISO/US date and timestamp values, date systems, explicit-midnight/missing distinction, Excel Table support and safe rerun refresh.
4. Hundred-block tools v0.4: in-sheet columns, normal/00/XX/BLOCK/BLK parsing, city/state-qualified keys, seeded synthetic-address generation and incident/offense consistency. Synthetic precision warnings propagate into Census results.
5. Exact-row and repeated-case review without deletion.
6. Crime Concentration v0.5: new dashboard with crime-field mapping, live dropdown filtering, distinct incident/record modes, selected-crime populated-block universe, 1/5/10/20% shares, 50/80% thresholds, coverage, ranked detail and three native charts. Hidden local snapshots and application events support filtering across saved workbooks.
7. Census Geocoder v0.3: HTTPS submission of distinct addresses, row/UTF-8 byte batch limits, request-ID response validation, in-sheet results, progress, retries and completed-batch retention. The older local staging procedure remains available only in VBA for compatibility.
8. Native Windows Excel build, executable VBA assertions, package validation and hash manifest. Every custom ribbon icon is checked natively in VBA.

## Pilot acceptance

- Confirm tab and all six groups display on the pilot Excel versions.
- Run Timestamp Fields on the selected timestamp column. Verify all eight columns appear immediately to the right in the requested order, neighboring values/formulas shift correctly, and a rerun refreshes the same eight fields. Test an Excel Table and a plain range. Other tools still create new output sheets. Save/reopen results.
- Select the newly generated hundred-block table including its row-6 header and run both concentration modes; reconcile source record and unique-incident counts.
- Test cancellation of range/field prompts, duplicate headers, errors, missing IDs, no eligible blocks, read-only/protected workbook rejection and reruns.
- Test 1904-system workbook, mixed date-only/datetime rows, US month/day/year text with AM/PM, explicit midnight, invalid ISO dates, filtered/hidden source rows and leading-zero IDs/ZIPs.
- Validate Maps and 3D Maps launch using approved sample geography; do not send crime data during the pilot.
- Test representative 50k/200k rows and 32/64-bit Windows Excel before promising company performance.

## Next modules

1. Replace repeated field prompts with a mapping dialog and saved workbook-level profiles. Define agency/year/offense grain, source timezone and occurrence/report field semantics.
2. Add explicit analysis period and jurisdiction scope, approved street aliases, stable offense keys and reviewed corrections. Preserve unreviewed flags and source lineage.
3. Extend Census Geocoder with persistent cache/resume, optional Census geographies, explicit benchmark selection/metadata capture and a saved run ledger. The current version already performs minimal-field network submission, 10k/byte splitting, CSV escaping, retries and review status handling. Reruns refresh all eligible addresses; they do not resume across sessions.
4. Expand validation to record and incident reconciliation, missingness and geocoding coverage; add map-ready aggregated output without raw identifying columns.
5. Add optional block universe, time comparisons, offense classifications, repeat-location summaries, day/hour distributions and rates only with compatible denominators and coverage.
6. Sign a release candidate, pilot company deployment, then publish a versioned package and rollback instructions. Keep development VBA/source out of general user installation directories.

## Boundaries

Build validation uses synthetic fixtures. Reconcile representative approved datasets during pilot acceptance. User acceptance of the ribbon, prompts and built-in mapping workflow, real-data reconciliation, performance, and IT signing/rollout remain separate tasks.

## Version 0.6 additions

Implemented Monday–Sunday/hour heat map, repeat locations with adjustable consecutive-gap windows, and weekly/monthly descriptive street profiles across the entire source date span. Native Excel validation and synthetic chart previews accompany the build. Optional local Python modeling remains a future module; current labels do not claim fitted trajectory groups.

## Version 0.7 additions

Compact named geography sheets; Hour from Time; four heat-map input modes including separate date/time and existing day/hour; address-only repeats; internal city/state geography qualification; explicit selection prompts. Regression checks exercise retained engines and revised public workflows.

## Version 0.8 additions

Distinct incident crime-type lists; date-only Timestamp outputs with safe six/eight-field refresh; complete native working Tables for all geography outputs. Validate source field/row preservation and chained block/randomizer/geocoder datasets.

## Version 0.9 additions

Monthly/annual trends, two-period comparison and crime profiles; scoped PNG, static PowerPoint and PDF exports. Mapping automation postponed at the user's request. Validate native rendering, no workbook/OLE embeddings, state restoration, existing-file protection and report counting units.


## Version 0.15.0
New modules: modAnalysisScope (workbook preference, scoped values and original-row mapping), modNIBRS (reviewed agency crosswalk and FBI reference), modCalendarAnalysis (YoY/Seasonality), modReportBuilder (section/content/order selectors and combined exports), modRelease015Tests (native regression fixtures). CAP15 ribbon callbacks isolate this release. Native builds verify controls/icons, analytical counts, partial/zero-baseline/leap-year/1904 dates, hidden/filtered rows, classification restrictions, saved preferences/selectors, combined PDF/PPT and nonblank PNG outputs. Mapping automation remains deferred.
