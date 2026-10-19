> Version 0.7.0 changes the user workflow and visible fields. See WORKFLOWS-0.7.md and README.md; older in-sheet examples below describe the retained calculation engine.

# Time and location analysis — 0.6.0

All three commands create a new sheet and a VeryHidden source snapshot. Select headers and data, then map occurrence date, crime type and optional stable incident ID. Source records and unique incidents are separate measures. Source data is unchanged. Crime/measure selectors update the saved snapshot automatically while the add-in and Excel events are enabled; Refresh Time Report also recalculates it. Rebuild to incorporate source edits. Hidden/filtered rows inside the selection are included.

## Temporal Heat Map

Monday–Sunday rows, hours 0–23 columns. Native conditional formatting runs from white at zero through yellow to red at the selected peak. Counts are not rates adjusted for activity or exposure. Mixed date/time and date-only rows are accepted using the Timestamp Fields parser (US month/day/year text and ISO/native Excel dates). Unknown time remains excluded, never assigned midnight. Explicit text midnight is eligible; integer numeric Excel dates lack time provenance. An optional explicit hour column accepts integers 0–23; blank/invalid hours or contradictions with timestamps are excluded. Coverage counts show excluded date/time and ID rows.

## Repeat Locations

Map a qualified full-address key, hundred-block key or stable street ID; optional readable label. The chosen granularity defines a location. Default minimum events is two. The report ranks counts and provides first/last date, mean/median/minimum consecutive calendar-day gap, number of gaps within 7/30/90 days and additional events beyond the first. Minimum count and ordered positive day windows are adjustable. Windows count consecutive gaps, not all pairs; same-day distinct incidents have zero-day gaps. Source-record mode can inflate repeats through multiple offense records. Exact address aliases are not resolved automatically.

## Hot Streets

Map a stable street ID to analyze whole streets, or CAP_BlockKey to analyze hundred blocks. Choose Weeks or Months and a positive number of units per interval (1–120). Weeks start Monday; months start on the first. The axis uses the earliest/latest valid date anywhere in the selected source, unchanged by crime filters. Empty intervening bins remain zero record counts. Agency reporting completeness is unverified: empty bins do not establish zero crime.

The matrix includes partial first/last intervals. Descriptive classifications require at least three complete calendar intervals wholly inside the observed source date span. Early/late means use the first/last third of complete bins, converted to events/day. Default rise/fall threshold is 25%, adjustable from 1% to 1000%. Rising/falling labels follow that threshold; zero early activity and positive late activity is Emerging from zero baseline. Otherwise Persistent higher-volume means top-quartile total activity and activity in at least 75% of complete bins; Intermittent means activity in fewer than half; remaining locations are Stable/mixed. Ties can expand the top quartile. These thresholds are descriptive and are not fitted group-based trajectory models.

The chart averages per-location daily counts within each descriptive class using the observed days in each bin, including edge bins. The matrix contains raw counts. Labels depend on the selected crime, measure, interval and observation span. Limits: 2,000 intervals and 400,000 output matrix cells; choose wider intervals if exceeded.

## Incident consistency and coverage

Unique incidents require nonblank IDs and consistent date/time (and for location reports location key) across all rows bearing that ID in the full snapshot. Conflicting or invalid records conservatively exclude that ID. Within a selected crime, an eligible ID contributes once. Source-record mode retains eligible rows independently. Coverage counts are row counts before deduplication; they are not all mutually interchangeable with incident totals. Hours are unnecessary for repeat and street reports.

Hidden snapshots contain original IDs and location keys. Share approved aggregated exports rather than treating hidden sheets as anonymization. Hundred-block keys remain generalized geography; randomized synthetic addresses do not recover actual locations.

## Python roadmap

Current descriptive tools run locally in VBA without Python installation. A future optional managed local Python engine could support larger panels and formal count-based group trajectory models with group selection, fit diagnostics and assignment uncertainty. It is not part of this build. Microsoft Python in Excel runs code in the Microsoft Cloud and is a separate deployment option: https://support.microsoft.com/en-us/excel/python/data-security-and-python-in-excel
