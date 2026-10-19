# Build and Validation

The current public baseline is v1.2.3. Exported VBA, form builders, RibbonX and embedded artwork are the source for the Windows desktop Excel .xlam add-in.

## Build

Use tools/Build-Addin.ps1 on an approved Windows Excel build workstation with authorized VBA project access. The script imports source into an isolated workbook, creates forms, saves add-in format 55, runs native regression checks, embeds RibbonX/icons, reopens the package and validates it again. It does not alter macro security or install the add-in.

Public builds require approved sensitivity classification supplied through local parameters. Do not commit organization policy identifiers. The build uses a neutral process-local Forms cache, removes personal save-folder metadata, verifies the saved classification, and reopens the cleaned package. See [Deployment](DEPLOYMENT.md).

Each build has a timestamp directory. FAILED or UNVALIDATED builds must not be distributed. A successful output includes the add-in, validation.txt, SHA256.txt and license/branding notices. Build folders also contain development workbooks and exports; do not publish the whole directory.

## Automated verification

Native fixtures exercise setup validation, field mappings, source preservation, incident/offense distinctions, mixed dates/times, geocoding request/response handling and batch splitting, seeded block randomization, analytical reports, scope, NIBRS review, calendar comparisons, single-site distances, and PNG/PDF/PowerPoint exports. Packaging checks cover icon relationships, personal paths, approved classification and post-cleanup reopening. Check validation.txt from the exact distributed build.

Live Census submission is optional and separate from synthetic transport tests. Automated tests do not establish live service availability, agency-data completeness or interactive visual acceptance.

## Manual acceptance checklist

- Enable only the intended add-in version and verify ribbon visibility in light and dark modes.
- Check setup windows on the intended screen size; exercise Run, Help, Cancel and correction/retry.
- Use an approved dataset copy and verify selected fields, rows, scope and counting measure.
- Review invalid/excluded records and confirm source sheets remain unchanged where promised.
- Check chart readability and PNG, PDF and PowerPoint outputs before sharing.
- Confirm connected-service permissions, local export folders and internal deployment policy.
- Retain the prior approved artifact for rollback; never overwrite user data.
