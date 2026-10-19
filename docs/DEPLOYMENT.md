# Company deployment

Copyright (c) 2026 CAP Index. Licensed under the MIT License.

Distribute `LICENSE`, `BRANDING.md` and `OWNERSHIP-NOTICE.txt` with each add-in release. The build copies them into the versioned build folder. Original toolkit software and documentation are MIT licensed; retain the copyright and license notices. Source data and third-party materials retain their respective ownership. CAP names and logos have separate branding terms.

Digital signing verifies publisher identity and changes to signed code. This build remains unsigned; deployment signing and pilot acceptance are separate from software licensing.

Target: Windows desktop Excel with VBA and RibbonX. Excel for the web cannot run this add-in. The code has no Windows API declarations or third-party references; 32/64-bit compatibility still requires pilot checks. Excel's standard Office object library supplies Ribbon callback types. Scripting.Dictionary and VBScript.RegExp are late-bound Windows components.

## Build and release

1. Build on an approved Windows Excel workstation. VBA project object-model access is needed only for automated building, not normal end-user use. If access is blocked, manually import the standard `.bas` modules and `.cls` class module using Windows CRLF/ANSI import files, compile the VBA project, save a new workbook as `.xlam`, then inject the supplied RibbonX with a suitable packaging editor. The build script produces correctly formatted import files in its build folder. Do not relax company security settings to build.
2. Review `validation.txt` and SHA256. Complete the pilot checklist in BUILD-PLAN.md. A successful native macro run establishes compilation of exercised paths, not visual acceptance of the ribbon.
3. Follow CAP Index IT policy for internal deployment. If that policy requires signing, sign the finalized VBA project with an approved certificate using IT's process. Complete packaging before signing, verify the signature, and calculate the final hash afterward. Code changes require new validation and, when signed, a new signature. Public GitHub availability does not establish an internal signing requirement.
4. Publish an immutable versioned release through the company software distribution mechanism. Retain the prior approved version and its hash for rollback.

Install only the final versioned `.xlam` from a successful build folder containing `validation.txt` and `SHA256.txt` and no `FAILED.txt`. Files marked `UNVALIDATED` are intermediate build products. Importing VBA alone does not install the ribbon, embedded icons or setup forms.

## Install

Install a local per-user or IT-managed copy of the approved `.xlam`, then enable it through Excel Add-ins. Use the company-approved trusted publisher/location policy and connected-experience policy. Do not disable macro protections globally. Network/OneDrive locations may have availability, concurrent edit, synchronization or trust issues; the selected OneDrive folder is a development workspace, not automatically the final execution location.

Close Excel before replacing the installed add-in. Ensure only one toolkit version is enabled. Update the release manifest/version and pilot package; do not overwrite user data. Roll back by disabling the new add-in and restoring/enabling the prior approved artifact. The toolkit adds output sheets but has no automated undo or source-workbook save.

## Operating controls

- Data owner defines record grain, key scope, reporting period/jurisdiction and approved data sharing.
- Analysts work on copies, review statuses and preserve run provenance. Summary outputs do not prove completeness of agency data.
- IT owns macro policy, signing, deployment/rollback and mapping connected-experience permissions.
- Product owner approves classification thresholds, address normalization and future transformations. Document and validate changes to analytical definitions before release.

No production install, signing or company-wide deployment is performed by the build script.

## Unsigned public release and internal policy

The public release is unsigned. Internal deployment is subject to CAP Index IT policy, including signing where required. The build report does not establish CAP policy or imply that public GitHub distribution is awaiting signing. Adoption feedback can inform future signing and support priorities; it does not override an internal requirement. See the README for first-time download verification and installation.
