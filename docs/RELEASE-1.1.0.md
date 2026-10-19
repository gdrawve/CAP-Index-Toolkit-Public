# CAP Index Toolkit 1.1.0

October 15, 2026

- Crime Concentration now begins with the number of blocks with eligible recorded crime. Its research note distinguishes studies including zero-crime places. Calculations are unchanged; crime filtering still uses blocks with the selected crime only. Older reports acquire the headline on refresh; rebuild for the research note.
- Hot Streets adds Minimum comparison events at B9, default 2, whole numbers from 1 to 200,000. It counts eligible events in the first and last thirds of complete intervals combined, using the chosen crime and incident/record measure. Middle intervals do not contribute to the minimum. Below the minimum, rates remain visible but the label becomes Low volume; no trend label. Fewer than three complete intervals still takes priority. Controls refresh through the existing Refresh button and event handling; older Hot Streets reports acquire the new control on refresh. These are descriptive labels, not significance tests.
- Hundred Blocks, Randomize and Census Geocoder copy exactly the selected rows and fields, including headers, on protected and unprotected sources and Excel Tables. They do not expand selections. Include all required fields and exclude table totals. Original source cells remain unchanged.
- Census Geocoder pads 3-5 digit ZIPs to five digits; 6-9 digit values are inferred ZIP+4 values padded to nine digits and hyphenated. Review Flag retains a repair warning, with an explicit inference warning for ZIP+4. Repaired ZIPs are used for full addresses, validation, deduplication and submission; original source ZIP fields remain unchanged. Valid hyphenated ZIPs are preserved. Malformed values still follow existing validation. Legacy PrepareCensus is unchanged.

Disable 1.0.0 before enabling 1.1.0. Ownership notices and data/coverage limitations remain. The package is unsigned; automated mapping and code signing remain outside this change.

During validation an Excel CopyPicture 1004 failure was reported. Range-image export now retries handled capture failures using screen vector/bitmap formats and a final printer fallback, cleans temporary charts and failed files, and reports a recoverable export error if all attempts fail. This shared capture routine is used by image and PowerPoint exports; PDF paths that render table summaries also benefit.
