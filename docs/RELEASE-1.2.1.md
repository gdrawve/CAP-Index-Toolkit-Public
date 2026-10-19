# CAP Index Toolkit 1.2.1

October 19, 2026.

This maintenance release removes retired score recategorization code and its tests/fixture field. It also removes internal workbook references and updates source-publication guidance. Current ribbon tools, setup windows, NIBRS classification, analytical calculations and report controls retain their behavior. Grant Drawve (gdrawve@capindex.com) remains the designated contact.

Disable the previous add-in and restart Excel before enabling the validated 1.2.1 package. Existing 1.2.0 and earlier files still contain their original code. Use the new package for reviewed distribution. The package remains unsigned; deployment signing and pilot acceptance remain separate steps.

See validation.txt in the successful build folder for native checks. Old history and tags are retained in the private development repository; they are not included in a clean public source snapshot.

## Validation

Native Windows Excel build passed: 150 setup checks, 68 command integration checks, 36 reporting/export checks and 35 scope/NIBRS/calendar/selected-export checks, plus the remaining timestamp, geography, Census, concentration, quality and single-site suites. Packaged reopen checks passed; RibbonX and embedded icons were verified. Live Census submission was not requested. New manual ribbon/Maps acceptance was not performed for this maintenance cleanup.
