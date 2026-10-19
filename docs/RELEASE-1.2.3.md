# CAP Index Toolkit 1.2.3 — Public Baseline

October 19, 2026.

Version 1.2.3 is the current public baseline for installation, use, and future development. Earlier builds remain historical references.

## Included capabilities

- Crime Data Prep: Dataset Check and Coverage, duplicate/case review, and source-quality guidance.
- Analysis Scope and Classification: selected-data workflows, incident versus offense counting, and reviewed NIBRS crosswalks.
- Geography: complete working datasets for hundred blocks, randomized block addresses, and Census batch geocoding.
- Date and Time: mixed date/timestamp field creation and a separate Hour from Time tool.
- Crime Analysis: temporal heat maps, repeat locations, Hot Streets, crime concentration, trends, period comparisons, year-over-year, seasonality, and single-site distance analysis.
- Reporting and Export: selected reports, PNG images, static PowerPoint briefings, and PDFs.
- Usability: CAP ribbon icons, sixteen consistent setup windows, tool-specific help, one Refresh command, and Excel essentials.

## Installation and distribution

Download the installable release ZIP and follow the README's SHA256 verification and Windows Unblock instructions. The package uses the approved Public sensitivity classification, removes saved personal build-folder hints, and uses a neutral process-local Microsoft Forms cache. The finished add-in is reopened and validated in Excel after metadata cleanup.

Code and documentation are MIT licensed, with separate CAP branding terms. The README explains the toolkit's purpose and includes Grant Drawve's note connecting it with CAP Index's Data for Good initiative.

The release is unsigned. Internal deployment is subject to CAP Index IT policy, including signing where required. This statement does not place the public release on hold.

## Interpretation and validation

Preserve the distinction between unique incidents and offense records. Review source coverage, address matching, geocoding limitations, and synthetic-location flags before interpreting outputs. The tool guide and data definitions describe these boundaries.

Native Excel regression suites passed, including PNG, PDF, and PowerPoint checks. Post-cleanup About, setup-form, 37 core and 150 setup assertions passed. Saved Public classification and absence of personal profile paths were verified. See validation.txt for details. Live Census submission, ribbon display/Maps launch, and manual pilot acceptance are separate from automated validation.