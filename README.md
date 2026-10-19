# CAP Index Crime Analysis Excel Toolkit

Version 1.2.2 is the first public release of the Windows desktop Excel VBA/RibbonX add-in. Geography tools create complete working Excel Tables on new named sheets, copying only selected fields and rows, including headers plus relevant generated results. Hour from Time supports datasets whose date and time are separate. Analysis prompts state the required inputs before selection. Reporting & Export now adds trends, period comparisons, crime profiles, PNG images, static PowerPoint briefings and PDFs. Mapping changes are deferred; existing Maps/3D Maps commands remain. This package is unsigned.

Copyright (c) 2026 CAP Index. Original toolkit code and documentation are licensed under the [MIT License](LICENSE). CAP names and logos are covered separately by [Branding](BRANDING.md). See [Ownership Notice](OWNERSHIP-NOTICE.txt). Source agency data and third-party material retain their respective ownership.

Download the installable package from [GitHub Releases](https://github.com/gdrawve/CAP-Index-Toolkit-Public/releases/latest). Choose the release ZIP containing the .xlam add-in and license files; the automatic source-code archives are for development.

## Start

1. Work from a copy of the dataset. Disable the previous CAP Index add-in, then use File > Options > Add-ins > Manage Excel Add-ins > Go > Browse to enable this version's `.xlam`.

2. Each tool explains its required selection. Geography and analysis tools use only the selected rows and fields. Include headers and every field needed by the tool; selections are not expanded. For generated geography sheets, select the header/data table starting at row 6, excluding rows 1-5 of title/run notes. Include all fields needed for the next analysis. Analyses use the workbook Analysis Scope (All Rows by default, or Visible Rows Only). Preparation tools process all selected rows.

3. Geography outputs preserve source fields so date, crime type, incident ID, city and state remain available for follow-on tools. The original source sheet is not modified. Source sheets may be protected; workbook structure must permit a new sheet and the workbook must be writable.

4. Timestamp Fields and Hour from Time insert columns beside their selected source column. These two tools require an unprotected source sheet. Timestamp Fields produces six date fields when no valid time exists, or all eight date/time fields when at least one valid time exists; Hour from Time produces only Hour (0-23).

5. Rebuild analyses after source changes; their crime/measure controls recalculate a saved hidden snapshot. Save the data workbook to retain results. No automatic source-workbook save occurs.

## Tool inputs and results

| Tool | Required inputs | Results |
|---|---|---|
| Hundred-Block Builder | Address | New Hundred_Blocks sheet: selected source fields, CAP_HundredBlock, Review Status |
| Randomize Block Addresses | Agency block Address; optional IncidentID/City/State and seed | New Randomized_Addresses sheet: selected source fields, CAP_HundredBlock, CAP_SyntheticAddress, Review Status; seed recorded in run notes |
| Census Geocoder | Address + City/State or ZIP | New Geocoded_Addresses sheet: selected source fields, Full Address, Matched Address, Latitude, Longitude, Match Status, Match Type, Review Flag; benchmark/run time recorded once |
| Timestamp Fields | One Date/timestamp column including header | Six date fields beside source; Time Only and Hour (0-23) are added only when valid times exist |
| Hour from Time | One Time column including header | Only Hour (0-23), beside Time; accepts numeric Excel times and 24-hour/AM/PM text; blank/invalid times stay blank |
| Temporal Heat Map | Crime Type plus Date/timestamp, separate Date + Time, Day of Week + Hour, or Day of Week + Time | New sheet: Monday-Sunday x hours 0-23 with native conditional colors and exclusion counts |
| Repeat Locations | Address + Crime Type; optional IncidentID/City/State | New address ranking: Address, Event Count, Additional Events; no date/time required |
| Hot Streets | Hundred Block or street ID + Crime Type + Date; optional IncidentID/City/State | New weekly/monthly timeline and descriptive profiles; time is not required |
| Crime Concentration | Hundred Block + Crime Type; optional IncidentID/City/State | Concentration metrics and three charts; percentages are not directly comparable with published estimates that include zero-crime places |
| Duplicate / Case Review | Selected header/data table; optional incident key | Review sheet including distinct crime types associated with each incident; all offense rows retained |
| Monthly / Annual Trends | Date + Crime Type; optional IncidentID | Continuous periods, counts, daily rates, comparable-period percentage change and chart |
| Period Comparison | Date + Crime Type; optional IncidentID and qualified location | Two disjoint periods: counts, differences, daily rates and A/B chart |
| Crime Profile | Date + Crime Type; optional IncidentID and qualified location | Incident/offense totals, crime mix, coverage and top 20 locations |
| Save Image | Selected report chart or formatted range | Captioned chart PNG or report-range PNG |
| PowerPoint Briefing | Selected CAP report sheets | Widescreen summary/chart slides as static pictures, no embedded source workbook |
| Save Report PDF | Active CAP report sheet | Landscape report PDF with charts and page numbering |
| Maps / 3D Maps | Appropriate map-ready data | Excel's existing commands |

Use stable incident IDs to count unique incidents. Source-record mode counts offense rows separately. City/state fields prevent identical address text in different jurisdictions being combined. If omitted, the selected source must already represent one jurisdiction or supply a qualified location key. Street aliases and address spelling variations require review. Synthetic addresses estimate positions within blocks and do not recover actual incident locations.

The heat map accepts existing Day of Week and Hour fields without requiring dates. Day names may be Monday-Sunday or Mon-Sun; numeric weekdays are rejected because numbering conventions differ. Hot Streets needs dates to place events in intervals; hundred block and crime type alone support concentration, not temporal change. Empty periods are observed source counts, not verified zero-crime periods.

See [Reporting & Export](docs/REPORTING-EXPORT.md), [current dataset changes](docs/WORKFLOWS-0.8.md) and [revised workflows](docs/WORKFLOWS-0.7.md) and the detailed behavior documents for counting definitions and limits.

## Files

```text

CAP-Index-Toolkit/

  README.md

  src/

    ribbon/customUI14.xml

    vba/modCore.bas

        modRibbon.bas

        modDateTime.bas

        modGeography.bas

        modHundredBlocks.bas

        modBlockTests.bas

        modReview.bas

        modPareto.bas

        modConcentration.bas

        modConcentrationTests.bas

        clsCAPEvents.cls

        modCensus.bas

        modCensusGeocode.bas

        modCensusTransport.bas

        modCensusTests.bas

        clsCensusFixture.cls

        modTests.bas

  docs/

    BUILD-PLAN.md

    DATA-CONTRACT.md

    DEPLOYMENT.md

    SOURCES.md

  tests/fixtures/crime-synthetic.csv

  tools/Build-Addin.ps1

  build/<timestamp>/

    CAP-Index-Toolkit-1.2.2.xlam

    validation.txt

    SHA256.txt

```

Build from source using `tools/Build-Addin.ps1` on a Windows Excel build workstation. It imports VBA into a new workbook, saves format 55, runs assertions, adds RibbonX package parts, reopens the add-in and reruns assertions. It does not change macro security, install the add-in, or touch existing workbooks. Build outputs receive new timestamp folders. Failed builds contain `FAILED.txt`; do not deploy those folders.

When upgrading, disable the old add-in entry and browse to the new version so only one CAP Index tab is loaded. Prior builds remain available for rollback. Timestamp Fields, Census Geocoder and the two block tools follow the user's in-sheet workflows. PrepareCensus and BuildHundredBlocksLegacy remain in VBA for compatibility and regression checks; they are not ribbon actions.

Crime Concentration stores a hidden local source snapshot for filtering. It contains incident IDs, so share an approved summary/export when those identifiers are sensitive. Rebuild after source edits. Live selector updates require the enabled Windows Excel add-in; the report identifies its displayed filter/measure and flags changed filters that have not been refreshed.

See [data definitions](docs/DATA-CONTRACT.md), [build plan](docs/BUILD-PLAN.md) and [deployment](docs/DEPLOYMENT.md).

## Version 1.2.2 (first public release)

The first public release includes crime data preparation, hundred-block tools, Census geocoding, temporal and location analyses, reviewed NIBRS classification, and reporting/export tools. Sixteen tools share a CAP-branded setup window with field dropdowns, options, validation and tool-specific help.

Original toolkit code and documentation are MIT licensed. CAP ownership attribution remains; names and logos have separate branding terms. See [1.2.2 release notes](docs/RELEASE-1.2.2.md) and the [deployment guide](docs/DEPLOYMENT.md).

[View the tool setup window](docs/setup-window-1.2.0.png).
