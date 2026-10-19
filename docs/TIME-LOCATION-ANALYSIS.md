# Time and Location Analysis

Select headers and data, open the tool's setup window, map its required fields, choose options, and Run. Reports are created on new sheets; source data is unchanged. New analyses follow [Analysis Scope](ANALYSIS-WORKFLOWS.md). Report filters and Refresh Report use a saved hidden snapshot; rebuild to reload source changes. Hidden snapshots are not anonymization.

## Temporal Heat Map

Monday–Sunday rows and hours 0–23 columns use conditional colors from white at zero through yellow to red at the selected peak. Choose Date/timestamp, separate Date + Time, Day of Week + Hour, or Day of Week + Time. Map Crime Type and optionally a stable IncidentID. English full or abbreviated weekday names are accepted; numeric weekday conventions are not assumed.

Unknown times are excluded, never assigned midnight. Explicit text midnight or zero in a dedicated Time column is eligible; an integer numeric date lacks time provenance. Explicit Hour must be an integer 0–23. Missing/invalid values and contradictions are excluded and counted. A Day of Week/Hour analysis has no calendar-date coverage unless dates were mapped. Counts are not exposure-adjusted rates.

## Repeat Locations

Map Address and Crime Type, optionally IncidentID/City/State. Dates and times are not required. The report ranks Address, Event Count and Additional Events, with a crime/measure filter, chart and adjustable minimum count (default two). Apartment/unit text remains part of the address. Case and repeated whitespace normalize; spelling aliases are not resolved. Qualify locations across jurisdictions. Conflicting incident locations are excluded in unique-incident mode. Source-record mode can include multiple offense rows.

## Hot Streets

Map Hundred Block (such as CAP_HundredBlock) or a stable street ID, Crime Type and Date; optionally map IncidentID/City/State. Date-only inputs work; time is not required. Choose Weeks or Months and a positive number of units per interval (1–120). Weeks start Monday; months start on the first. The axis uses the earliest/latest valid date anywhere in the selected source, unchanged by crime filters. Empty intervening bins remain zero record counts. Agency reporting completeness is unverified: empty bins do not establish zero crime.

The matrix includes partial first/last intervals. Descriptive classifications require at least three complete calendar intervals wholly inside the observed source date span. Early/late means use the first/last third of complete bins, converted to events/day. Minimum comparison events is user controlled at B9 (default 2), counting the early and late comparison windows combined. Locations below that minimum receive no trend label. Default rise/fall threshold is 25%, adjustable from 1% to 1000%. Rising/falling labels follow that threshold; zero early activity and positive late activity is Emerging from zero baseline. Otherwise Persistent higher-volume means top-quartile total activity and activity in at least 75% of complete bins; Intermittent means activity in fewer than half; remaining locations are Stable/mixed. Ties can expand the top quartile. These thresholds are descriptive and are not fitted group-based trajectory models.

The chart averages per-location daily counts within each descriptive class using the observed days in each bin, including edge bins. The matrix contains raw counts. Labels depend on the selected crime, measure, interval and observation span. Limits: 2,000 intervals and 400,000 output matrix cells; choose wider intervals if exceeded.

## Incident consistency and coverage

Unique incidents require nonblank IDs and consistent date/time (and for location reports location key) across all rows bearing that ID in the full snapshot. Conflicting or invalid records conservatively exclude that ID. Within a selected crime, an eligible ID contributes once. Source-record mode retains eligible rows independently. Coverage counts are row counts before deduplication; they are not all mutually interchangeable with incident totals. Hours are unnecessary for repeat and street reports. Repeat Locations does not require date consistency; temporal tools apply consistency checks for the mapped temporal fields.

Hidden snapshots contain original IDs and location keys. Share approved aggregated exports rather than treating hidden sheets as anonymization. Hundred-block keys remain generalized geography; randomized synthetic addresses do not recover actual locations.

## Runtime and mapping

These descriptive tools run locally in VBA without a Python installation. Hot Streets labels are descriptive rules, not fitted group-based trajectory models. Maps and 3D Maps use Excel's existing commands; automated basemap or density-layer configuration is not included.
