# Sources and verification

Checked October 13, 2026.

- Census batch API, five-field CSV format, quoting, 10,000-record limit and benchmark/vintage behavior: https://geocoding.geo.census.gov/geocoder/Geocoding_Services_API.html
- Census documentation and address-range coordinate methodology: https://www.census.gov/programs-surveys/geography/technical-documentation/complete-technical-documentation/census-geocoder.html
- Official Office control identifiers, Microsoft 365 Current Channel: https://github.com/OfficeDev/office-fluent-ui-command-identifiers/blob/main/Microsoft%20365/Current%20Channel/excelcontrols.xlsx
  - `ChartTypeMapsInsertGallery` (gallery; TabInsert / GroupInsertChartsExcel).
  - `PowerMapLaunch` (button; TabInsert / GroupPowerMap).
- Native add-in file format: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat (`xlOpenXMLAddIn = 55`).
- Workbook save method: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
- RibbonX callback/control concepts: https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-a-managed-com-add-in
- 3D Maps: https://support.microsoft.com/en-us/excel/get-started-with-3d-maps
