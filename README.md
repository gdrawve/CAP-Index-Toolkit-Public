# CAP Index Crime Analysis Excel Toolkit

The CAP Index Crime Analysis Excel Toolkit helps crime analysts prepare incident and offense data, examine patterns across places and time, and produce consistent reports within Excel. It brings common analytical tasks into one ribbon to reduce repetitive work and make workflows easier to repeat. Results support descriptive analysis and require review of data quality, coverage, and counting definitions.

Version 1.2.3 is the public baseline for the Windows desktop Excel VBA/RibbonX add-in. The CAP Index ribbon brings together data preparation, geography, date and time fields, reviewed NIBRS classification, crime analysis, Excel essentials, and reporting/export. Sixteen tools share a branded setup window with field dropdowns, options, validation, and tool-specific help. This package is unsigned.

Copyright (c) 2026 CAP Index. Original toolkit code and documentation are licensed under the [MIT License](LICENSE). CAP names and logos are covered separately by [Branding](BRANDING.md). See [Ownership Notice](OWNERSHIP-NOTICE.txt). Source agency data and third-party material retain their respective ownership.

Download the installable package from [GitHub Releases](https://github.com/gdrawve/CAP-Index-Toolkit-Public/releases/latest). Choose the release ZIP containing the .xlam add-in and license files; the automatic source-code archives are for development.

## First-time installation

1. Download `CAP-Index-Toolkit-release-1.2.3.zip` from [GitHub Releases](https://github.com/gdrawve/CAP-Index-Toolkit-Public/releases/latest). Use the installable release ZIP, not the automatic source-code archive.
2. Extract the ZIP to a permanent local folder. Keep the add-in, `LICENSE`, `BRANDING.md`, `OWNERSHIP-NOTICE.txt`, `SHA256.txt` and `validation.txt` together.
3. Before opening the add-in, open PowerShell in that folder and run:

   ```powershell
   Get-FileHash -LiteralPath '.\CAP-Index-Toolkit-1.2.3.xlam' -Algorithm SHA256
   ```

   Compare the returned Hash with the value beside the same filename in `SHA256.txt` and the digest displayed for that add-in on GitHub Releases. Letter case does not matter. If they differ, do not install that copy. The checksum verifies file integrity, not publisher identity or safety. [Microsoft checksum guidance](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-filehash).
4. For the reviewed download, right-click the extracted `.xlam` in File Explorer, select **Properties**, then **Unblock** and **Apply** if the Unblock option appears. This addresses Windows' internet-file marking; it does not override an organization's security policy. [Microsoft guidance for Excel add-ins](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked).
5. Disable older CAP Index add-ins, close every Excel window, and reopen Excel. Select **File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse**, select the extracted `.xlam`, and enable its checkbox. The CAP Index ribbon tab should appear.
6. Work from a copy of your dataset. If company policy blocks the add-in, contact IT rather than disabling macro protections globally. This release is unsigned; internal deployment follows CAP Index IT policy, including signing where required.

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

    CAP-Index-Toolkit-1.2.3.xlam

    validation.txt

    SHA256.txt

```

Build from source using `tools/Build-Addin.ps1` on a Windows Excel build workstation. It imports VBA into a new workbook, saves format 55, runs assertions, adds RibbonX package parts, reopens the add-in and reruns assertions. It does not change macro security, install the add-in, or touch existing workbooks. Build outputs receive new timestamp folders. Failed builds contain `FAILED.txt`; do not deploy those folders.

When upgrading, disable the old add-in entry and browse to the new version so only one CAP Index tab is loaded. Prior builds remain available for rollback. Timestamp Fields, Census Geocoder and the two block tools follow the user's in-sheet workflows. PrepareCensus and BuildHundredBlocksLegacy remain in VBA for compatibility and regression checks; they are not ribbon actions.

Crime Concentration stores a hidden local source snapshot for filtering. It contains incident IDs, so share an approved summary/export when those identifiers are sensitive. Rebuild after source edits. Live selector updates require the enabled Windows Excel add-in; the report identifies its displayed filter/measure and flags changed filters that have not been refreshed.

See [data definitions](docs/DATA-CONTRACT.md), [build plan](docs/BUILD-PLAN.md) and [deployment](docs/DEPLOYMENT.md).

## A note from Grant

I developed this toolkit to make everyday crime analysis more accessible and reduce the repetitive preparation that takes time away from understanding the data.

It reflects the spirit of [CAP Index's Data for Good initiative](https://capindex.com/data-for-good/): expanding access to analytical resources and supporting responsible research and informed decisions. My goal is to give analysts and researchers practical tools they can use, examine, and improve while keeping data quality, context, and interpretation central to their work.

cheers\
Grant Drawve

## Version 1.2.3 — public baseline

This is the starting point for public installation, use, and future development. The release includes:

- Dataset Check and Coverage, duplicate/case review, and analysis scope controls.
- Timestamp and hour fields, hundred-block datasets, randomized block addresses, and Census batch geocoding.
- Reviewed NIBRS crosswalks, temporal heat maps, repeat locations, Hot Streets, and hundred-block crime concentration.
- Trends, period comparisons, year-over-year and seasonality views, and single-site distance analysis.
- Report selection, PNG images, static PowerPoint briefings, and PDF exports.
- Consistent setup windows, ribbon icons, tool help, and Excel essentials.

The package uses CAP's approved Public sensitivity classification, excludes personal build-folder hints and personal Forms cache references, and is reopened and validated in Excel after packaging. SHA256 verification and installation instructions appear above.

Use the tool-input table and data definitions to review counting, coverage, geocoding, and synthetic-location limitations. Live Census submission and manual pilot acceptance are separate from automated validation. The add-in remains unsigned; internal deployment follows CAP Index IT policy.

Original code and documentation are MIT licensed, with separate CAP branding terms. See [baseline release notes](docs/RELEASE-1.2.3.md) and the [deployment guide](docs/DEPLOYMENT.md).

[View the tool setup window](docs/setup-window-1.2.0.png).
