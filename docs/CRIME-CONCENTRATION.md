> Version 0.7.0 changes the user workflow and visible fields. See WORKFLOWS-0.7.md and README.md; older in-sheet examples below describe the retained calculation engine.

# Crime Concentration report

Version 0.5.0. CAP Index > Crime Analysis > Crime Concentration creates a new report sheet with a crime-type dropdown, count-measure dropdown, concentration summaries, a ranked block table and three native Excel charts.

## Create the report

Select a rectangle containing headers and source data, then map:

- Hundred-block key: usually CAP_BlockKey from Hundred-Block Builder. Include city/state/jurisdiction in the key; a street name alone combines different hundred blocks.
- Optional readable hundred-block label: usually CAP_HundredBlock, for chart labels.
- Crime-type field: the user selects the actual source header, such as CrimeType, OffenseDescription, or CrimeCategory. No classifications are invented.
- Optional IncidentID: provide a key unique across agencies/years to use distinct-incident counts. Without it, the initial measure is Source records.

Original data remains intact. The report owns a local, very hidden snapshot sheet containing only the mapped incident ID, block key/label, crime category, incident-validity tag and source row. Both sheets are saved in the source workbook. Rebuild the report from source after data changes; filter changes operate on the captured snapshot, not on live source cells or an external database. Hidden data is not anonymization: share an approved summary/export if the workbook's incident IDs are sensitive.

## Filter behavior

Blue cells B4 and B5 select Crime type and Count measure. The crime dropdown offers All crime types plus observed categories. Nonblank source categories are trimmed and normalized for case/repeated spaces, displayed as Type: CATEGORY to distinguish them from the All/missing/error options. Missing crime types and source cell errors have separate selectable categories; All includes them rather than silently losing their activity.

Changing either selector refreshes the report automatically while the CAP Index add-in is enabled and Excel events are active. Initialization is performed on add-in open and Ribbon load. Refresh Report on the ribbon recalculates the active report when an automatic event was disabled. Displayed crime and Displayed measure identify the actual saved chart/metric selection. A worksheet formula flags changed filters when displayed values have not been refreshed. With macros/add-in disabled, filters cannot rebuild the charts; enable the add-in and refresh. The underlying source snapshot still needs rebuilding after source edits.

The user explicitly selected the denominator **only hundred blocks with the selected crime**. Accordingly, the block universe contains positive eligible activity for the selected crime and selected measure. It changes when the crime filter changes. No zero-crime streets outside the source are inferred. Results do not claim coverage of all streets in a city.

## Measures and exclusions

Unique incidents counts a valid incident ID once within the selected filter. An incident with different block keys, or any missing/error block key, anywhere in the full source snapshot is excluded from every incident view. Crime filtering does not resolve a globally inconsistent location. Rows lacking an incident ID are also excluded from incident counts. Multiple offense categories within one incident can count that incident once in each filtered view, while All counts it once; filtered incident totals must not be added together.

Source records counts every selected row with a valid block key, including legitimate repeated offense rows, missing incident IDs and suspected duplicates. It is an offense count only when the source's record grain establishes that meaning. No automatic deletion/deduplication of offense records is performed.

Coverage shows selected source rows, rows missing a block key, rows missing IncidentID, and rows with missing/conflicting incident locations. Coverage counts may overlap and are not additive exclusion totals. The headline eligible-activity total and populated-block denominator reconcile to the ranked table.

## Summaries

- Actual activity shares held by the top 1%, 5%, 10% and 20% of populated blocks.
- Number of blocks used at each target, the actual rounded block share, and the last rank of ties at the cutoff.
- Minimum ranked block count/share required to reach 50% and 80% of eligible activity.
- Plain-language 5% headline, for example: Top 5% target uses 1 of 20 blocks (actual 5.0%), accounting for 65.0% of eligible incidents.

Block counts round upward using ceiling(target percentage × populated blocks). With four blocks, the 5% target uses one block and explicitly reports the actual 25% share; it does not label that one block as exactly 5%. Rankings use count descending and binary block-key order for ties. The selected prefix's activity share is invariant when tied blocks have the same count; tied membership is not a substantive distinction.

No 65%, 50% or 80/20 concentration pattern is assumed. No population-normalized crime rate, causal effect or significance test is inferred. Empty/invalid selections clear generated metrics/charts and show an unavailable/error status rather than retaining a healthy old result.

## Charts and ranked detail

1. Cumulative concentration curve: blocks ranked highest crime first; X is cumulative share of populated blocks and Y is cumulative share of selected eligible activity. A dashed equal-distribution line is a reference, not an expected crime distribution. For large universes, the plotted curve is sampled to approximately 1,000 ranks; all summary/table calculations still use every eligible block.
2. Highest-crime blocks: counts for up to the first 20 ranked hundred blocks, with numbered readable labels that map to the ranked detail.
3. Crime-count distribution: number of populated blocks containing 1, 2–5, 6–10, 11–25, 26–50, 51–100 or 101+ incidents/records. Bin counts sum to the selected block denominator.

The table contains key, count, label, rank, activity share, cumulative crime share and cumulative block share. Its ordinary row filter is for viewing detail; it does not recalculate the report. Use the blue crime/measure selectors to change the analytical scope. Each refresh replaces the three generated charts while preserving separately named user-added charts.

## Research interpretation

This is a descriptive analysis inspired by crime-concentration research. Address-number hundred-block bins are not necessarily street segments between intersections or validated street geometry. The selected-crime populated-block denominator also differs from studies using a complete street universe that includes zero-crime streets. The output identifies those limits and does not claim a direct replication of a published law or benchmark. [Weisburd (2015)](https://doi.org/10.1111/1745-9125.12070), [NIJ research summary](https://nij.ojp.gov/library/publications/law-crime-concentration-application-and-recommendations-future-research).

## Verification

Native tests verify a synthetic 100-incident/20-block dataset where the top 5% accounts for exactly 65%, an extra offense row that changes record counts but not unique incidents, automatic robbery/theft/All filters, a changing selected-crime denominator, small-sample rounding, three chart series and histogram reconciliation, source preservation, exclusion coverage, stale-filter warnings, and invalid/empty states. All three chart images were exported in native Excel and visually inspected. No real crime dataset was supplied or sent externally. Version 0.5.0 remains an unsigned pilot.
