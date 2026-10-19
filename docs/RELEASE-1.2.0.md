# CAP Index Toolkit 1.2.0

October 15, 2026

## Setup experience

One modal, branded setup window replaces sequential prompts for Hundred Blocks, Randomize Block Addresses, Census Geocoder, Crime Concentration, Temporal Heat Map, Repeat Locations, Hot Streets, Trends, Period Comparison, Crime Profile, Prepare NIBRS Crosswalk, Single Site, Year over Year, Seasonality, Dataset Check and Duplicate / Case Review.

Select headers and data, map fields using dropdowns, review options and click Run. The data summary names the workbook, sheet and selected range, row/field counts and applicable Analysis Scope. Geography tools use only selected rows and fields; no range expansion or extra confirmation is introduced.

Required fields have an asterisk. Heat-map fields change with the time-input choice. Census requires Address plus City/State or ZIP and displays its external-service notice. Count measure is fixed to Source records when Incident ID is absent. Reporting crime filters display plain source values. Dates default to the observed selected/scoped span; Period Comparison has both A and B date ranges. Single Site permits an undated source; Dataset Check permits all optional mappings to remain blank.

Help opens inside the setup window without creating a worksheet. Cancel and the window close button create no output. Run errors preserve the choices for correction and retry; completed output from a failed run may remain, particularly geocoder batches. Review it before retrying.

New CAP Role custom properties store mappings under explicit role keys. Safe older mappings are read as fallback; contaminated legacy Address, City and Hundred-Block values are ignored. Existing public tool signatures and direct source/header calls are retained. Calculations, output layouts, report-sheet controls, Refresh, Reset, exports and ribbon callback names are unchanged. Timestamp and Hour retain their single-column entry flow; Apply NIBRS retains its crosswalk action.

Disable 1.1.0 before enabling 1.2.0. The package remains unsigned. See validation.txt and acceptance-1.2.0.txt in the validated build for check results. Automated regression checks alone do not establish completion of the interactive acceptance pass.

Build safety: intermediate add-ins use an UNVALIDATED filename. The normal release filename is created only after native checks and ribbon/image packaging succeed. Never install an UNVALIDATED file or an add-in from a folder containing FAILED.txt.
