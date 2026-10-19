# CAP Index Toolkit 0.15.0

## Excel Essentials

Native Excel commands: Filter, Custom Sort, Freeze Panes, PivotTable, Table, Clear Filters, Text to Columns, Find, and Conditional Formatting. These commands also remain on their original Excel tabs and operate on the current selection. Clear Filters changes the visible source rows; it does not change a report's saved snapshot. Refresh Report remains the single CAP analytical refresh button.

## NIBRS classification

1. Select the source table, including headers. Click Prepare NIBRS Crosswalk under Classification. Map the crime label or supplied code, and optionally Agency. Specify federal/tribal applicability when asked.
2. The toolkit creates NIBRS_Reference and NIBRS_Crosswalk sheets. The reference follows the [FBI NIBRS User Manual 2025.0, section 2.3](https://le.fbi.gov/file-repository/nibrs-user-manual-2025-0-062625.pdf). Source labels are paired with Agency when mapped.
3. Exact official labels and recognized codes receive proposals. Broad descriptions such as Theft remain unresolved. Verify the code against the agency's definitions, edit it where necessary, and explicitly mark Approved = Yes. Proposed offense names update when the code changes. No proposal is approved automatically.
4. With the crosswalk sheet active, click Apply NIBRS Crosswalk. A new NIBRS_Classified working table retains the scoped source fields and records, plus code, offense name, Group A/B, crime-against category and classification status.

Unapproved mappings, duplicate crosswalk keys, obsolete codes and federal/tribal-only codes outside that applicability remain unclassified. Legacy 11B/11C proposals point to 11A and require review. Group B refers to arrest reporting and remains explicitly labelled. A reviewed analytical crosswalk does not validate all NIBRS reporting rules, incident structures, victims, arrests or submission requirements. It does not convert offense records into unique incidents. Qualify incident IDs across agencies.

The linked source range retains its selected boundaries. If new rows are added outside that range, prepare a new crosswalk using the expanded selection. Applying a crosswalk rereads the source in the current workbook Analysis Scope; new/unmapped labels remain flagged. Source sheets are unchanged.

## Analysis Scope

All Rows is the default. Visible Rows Only excludes filtered and manually hidden source rows, while preserving the header and original worksheet row references. Hidden columns remain included. Scope is saved with the data workbook and identified in new outputs with included/excluded counts.

Scope applies to new analytical reports, Dataset Check, Duplicate/Case Review and NIBRS classification. Hundred blocks, randomized addresses, geocoding, Timestamp Fields and Hour from Time process all selected preparation rows. An empty visible selection stops with an explanatory error.

Changing scope does not alter existing saved report snapshots. Rebuild with the original tool to change the scope or reload edited source data. Refresh Report and crime dropdowns recalculate the saved data only.

## Year over Year

Map Date and Crime Type, optionally IncidentID. Choose source/offense records or unique incidents, then the requested dates and monthly or annual intervals. The report starts with All crime types and uses the same blue crime dropdown and Refresh Report controls as other reports.

Each row compares the requested calendar period with the corresponding prior-year period in the selected source. Current and prior counts, counts per day, count change and daily-rate change are shown. Prior-year data must be present in the selected range even if it lies before the requested report start. Percentage changes are suppressed for partial or observed-edge comparisons; zero baselines are labelled. Leap-year differences are identified. The observed date span establishes calendar-bin eligibility, not verified reporting completeness.

## Seasonality

Map Date, Crime Type and optional IncidentID, choose a count measure and date range. The new sheet contains January-December rows and year columns with warmer colors for larger counts. Partial or observed-edge bins are labelled and excluded from averages. Select at most 20 years.

Comparable years and average events per full calendar month are shown. The daily rate pools event counts and calendar days across eligible years, including leap days. This accounts for month lengths; it does not adjust for population, reporting participation or exposure. Zero-record months within the date span are not proof of complete reporting. A chart shows the pooled daily rate. The crime dropdown refreshes the saved snapshot.

## Selectable reports

1. Create the analyses needed in the data workbook.
2. Click Select Report Sections. The new Report_Builder sheet lists available CAP reports.
3. Set Include = Yes for desired sections. Enter unique positive Order values, optionally edit Section Title, and choose Tables, Charts or Both. Unselected sections are omitted.
4. With the selector sheet active, click Export Selected Report. Choose PDF or PowerPoint and a new filename.

PDF contains full selected report tables and separate captioned chart pages. PowerPoint contains summary table pictures and chart slides, without embedding the source workbook. Long summary tables can be truncated according to the report's summary note; use PDF for full tables. Section selections, titles and order persist with the saved workbook. Refresh the reports before export; deleted reports or stale selectors produce an explanatory error. Recreate the selector to list newly created reports.

Sections retain their own crime filters, periods, count measures and scope. The builder does not combine totals from different sections. Review outputs before distribution. Mapping automation remains deferred. This is an unsigned development/pilot add-in; company deployment still requires the deployment checks in DEPLOYMENT.md.
