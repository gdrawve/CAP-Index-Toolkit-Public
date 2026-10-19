# Reporting & Export — 0.9.0

Mapping integration is deferred at the user's request. ArcGIS and 3D Maps retain their existing manual workflows; this release does not create automated basemap/density layers.

## Shared report inputs

Select headers/data with occurrence Date (or combined timestamp), Crime Type and optional stable IncidentID. Time is unnecessary. A qualified Address/Hundred Block is optional for location summaries; supply a full address containing city/state or a key that is already unique across jurisdictions. All selected rows, including hidden/filter-hidden rows, are processed. Choose a crime category, counting measure and inclusive start/end dates. Outputs are saved values on new named sheets; rebuild after source/filter changes. No source workbook is automatically saved.

Unique incident mode requires a mapped ID. IDs are trimmed and case-sensitive, and must be unique across agency/year when needed. All source records sharing an ID must agree on occurrence date and mapped location (when supplied); invalid/conflicting IDs are excluded conservatively. Different offense types within one incident are retained as category membership, not an ID conflict. Source-record mode retains independently eligible date rows and can include duplicates/multiple offenses. No offense-key deduplication is inferred. Date parsing follows the existing US month/day/year, ISO and native Excel/1904 behavior.

## Monthly / Annual Trends

Choose Months or Years. All periods in the requested span are present, including zero-record bins. Columns show calendar period, event count, requested days, events/day, period coverage note and change from the preceding period. Percentage change is only calculated when both calendar bins lie wholly inside the requested period and observed source date span; partial/edge comparisons show N/A. Zero baseline also shows N/A, not infinity. Daily rates describe counts per requested calendar day, not population/exposure rates. Source first/last incident dates do not verify agency reporting completeness. Maximum 1,000 periods.

## Period Comparison

The initially selected dates define period B. Choose an earlier, nonoverlapping period A. Rank by period B count, grouped by qualified location when mapped or by crime type otherwise. The table includes A/B counts, count difference, A/B events/day, raw count percentage change and daily-rate percentage change. Raw percentage change is suppressed when period lengths differ. Changes are suppressed when boundaries extend beyond the observed source date span; daily rates remain descriptive. Zero baselines show N/A. The chart compares A and B for up to 20 groups.

In incident mode an incident containing several offenses contributes once to each applicable crime category. Crime-group counts can therefore exceed the overall incident total; they are not mutually exclusive. Location grouping uses consistent incident location and does not replicate incident counts across offense rows.

## Crime Profile

When IncidentID is not mapped, incident-specific metrics say so rather than showing zero incidents. Metrics separate selected-measure total, dated source/offense records, valid unique incidents, missing/conflicting ID rows, crime categories (including missing/error labels) and events missing a mapped location. The crime mix lists counts and incidence percentages for each category, with a visible note that multi-offense incident membership can overlap. Percentages need not sum to 100% in incident mode. A separate ranking shows up to 20 locations with counts and events beyond the first. The crime mix chart shows up to 20 categories. Existing temporal heat map, repeat-location and concentration sheets can be selected alongside this profile for briefings.

## Save Image

Select an embedded report chart and click Save Image to export PNG with crime/measure/period notes. Otherwise export the active report summary or select a rectangular range including its title and notes (maximum 100 rows × 80 columns). A named summary may show only the first table rows; long new reports disclose this and retain full results in the worksheet. Chart and range exports leave original values intact and remove their temporary graphics. Existing export files are not overwritten; choose another name. Visible addresses or identifiers in a selected image remain visible; static output is not anonymization.

## PowerPoint Briefing

Requires Windows desktop PowerPoint. Choose 1–20 CAP report sheets in slide order. The builder creates a widescreen presentation with one summary slide and separate chart slides per report, CAP Index title bars, and crime/measure/date/coverage captions. All graphics are embedded static pictures; no Excel workbook/OLE object is embedded or linked. Existing open PowerPoint presentations are left open. The created briefing is saved to the chosen new `.pptx` filename; temporary image files are cleaned up.

Existing day/hour heat maps disclose when calendar dates were not mapped. Existing concentration reports disclose that their selected-source scope does not contain mapped dates. Changed report controls or failed refreshes require refresh before export. Source data edits are not automatically detectable in saved report snapshots; rebuild first. Long report summaries retain full tables in Excel rather than reducing text to unreadable slide sizes. Review image readability and slide scope before sharing.

## Save Report PDF

Select a CAP report sheet. Export its report body/summary in landscape with repeating title/scope/header rows and page numbers. Each chart has a separate captioned page so a wide side-by-side worksheet layout is not reduced to unreadable text. New long report tables continue over pages. For older dashboards, summary plus charts is used rather than the full hidden snapshot. Source data/hidden sheets are not included. A temporary workbook is used and closed; original page setup remains unchanged. Existing files are not overwritten.

## Validation and limits

Native Excel and PowerPoint tests use synthetic multi-offense records, conflicting IDs, invalid dates, monthly/yearly periods, partial bins, unequal comparisons and zero baselines. Exports are reopened and examined as PNG/PDF/PPTX. No real incident records are sent to a network service by these reporting/export commands. The add-in remains an unsigned pilot pending company signing and user acceptance.

Version 0.11.0: Monthly/Annual Trends and Period Comparison start with All crime types. Use the blue Crime type dropdown at B3 to refresh counts and charts on the same sheet. Reset to All crime types to restore the full view. Date ranges and count measure remain fixed. Filtering uses a hidden saved snapshot; rebuild after source edits. Refresh Time Report also refreshes these reports. Older reports need rebuilding to acquire the dropdown.

Export save dialogs use a local workbook folder when available. Web addresses are excluded from default filenames; select an accessible local or OneDrive-synced folder. Destination errors include the error number and description. Existing files are not overwritten. This handling applies to PNG, PowerPoint and PDF.
Version 0.14.0 adds Single Site under Crime Analysis. Select Latitude, Longitude, Crime Type and optional Date/IncidentID, then enter site coordinates and count measure. The Single_Site report contains nested radius totals within 0.25, 0.5 and 1 mile, separate distance bands, crime-type counts and exclusions. Site_Nearby retains selected nearby source/offense rows plus distance in miles, distance band and location precision. Its row count may exceed the unique-incident summary because multiple offenses are retained. Missing/invalid coordinates, 0,0 placeholders and flagged synthetic locations are excluded. Incident mode also excludes missing IDs and conflicting dates or coordinates across the whole selected source. IDs assume a single agency/key namespace. Distances use a mean-radius sphere (3958.7613 miles); they are approximate surface distances, not road distance. Decisions use unrounded values; Earth-model and coordinate uncertainty matter near boundaries. No coordinates are transmitted externally by this tool.

Single Site can filter crime and reporting dates in its blue controls. Dates are optional: an undated dataset explicitly displays Not mapped and counts all selected rows. Changing site coordinates, crime or mapped dates refreshes the report and nearby sheet from their private saved snapshot. Geocoding status is labelled when supplied; it does not establish positional accuracy.

One Refresh Report button now handles supported concentration, temporal/location, Trends, Period Comparison, Crime Profile and Single Site analyses. Reset to All restores the crime filter and retains other settings. Crime selectors are consistently at B3 in new reports; older internal B4 references remain available for compatible report calculations. Reports show version and last refresh status, with rebuild guidance for older versions. Refresh uses saved data; source changes require rebuilding with the original tool. Dataset Check and geography outputs remain static working outputs and give explicit guidance when Refresh is requested.

Field mappings are remembered by semantic role in the data workbook's custom properties and offered only when appropriate for the current headers. Save the workbook to retain them; the add-in does not automatically save source workbooks. Selection confirmation identifies the workbook, sheet, selected rows and field count; all selected rows, including filtered/hidden rows, are included. Output messages name the sheet and refer to exclusions/coverage. Long operations display status-bar progress.

Excel Essentials adds native Filter, Custom Sort, Freeze Panes and PivotTable commands while retaining their original Excel locations. The tool guide includes all new features. PNG capture retries a very small screen capture using Excel's print renderer; an available default printer is needed for that fallback. Native regression builds also inspect PNG pixels for blank images. The release remains an unsigned pilot; mapping automation remains deferred.