# Source control and publication

CAP Index maintains the working source for the Crime Analysis Excel Toolkit. The MIT License applies to original toolkit software and documentation in this release; CAP branding and third-party material have separate terms. See [Ownership Notice](../OWNERSHIP-NOTICE.txt).

## Repository contents

- Exported VBA, UserForm code, RibbonX and reviewed icon/brand assets.
- Build scripts, documentation and entirely synthetic fixtures.
- Reviewed changes with an explanation and relevant validation results.

Real incident datasets, agency workbooks, credentials, local backups and generated exports must not be committed. The synthetic CSV under tests/fixtures is the only approved CSV exception. Review staged changes before every commit; ignore rules do not remove already tracked files.

## Development

Use a short branch, validate changes in Windows desktop Excel, and document resulting behavior and checks in a commit or pull request. Exported source is authoritative; compiled Excel add-ins do not provide useful line-by-line diffs.

Run tools/Build-Addin.ps1 on an approved Windows Excel build workstation. VBA project access is needed for building, not normal add-in use. Do not change recipients' security settings to build or run the toolkit. Included PNG assets let ordinary builds run without Node; npm install and npm run icons are optional developer steps for regenerating icons.

## Public publication

Preserve the private development repository and its history. Prepare a separate reviewed source snapshot for a new public repository with a clean initial commit. Copy only approved tracked files; exclude .git, old builds, backups, workbooks and generated reports. Do not mirror private branches or version tags into the new public repository.

Removing content from the latest files does not remove it from earlier commits or tags. A public source repository exposes the implementation and included development documents. Review that scope before changing visibility. Retain the designated professional contact and CAP ownership notices.

## Releases

After native checks pass, tag the matching source commit with its version. Distribute only the validated .xlam, LICENSE, BRANDING.md, OWNERSHIP-NOTICE.txt, SHA256.txt and validation.txt with release notes and upgrade steps. Review every release asset for source data, local paths and metadata. Never upload entire build folders: they include temporary workbooks, reports and other development outputs. Failed or UNVALIDATED packages must never be released.

Maintain signing, pilot acceptance and rollback procedures described in [Deployment](DEPLOYMENT.md). Keep package.json private=true to prevent accidental npm publication; this setting is independent of GitHub repository visibility.
