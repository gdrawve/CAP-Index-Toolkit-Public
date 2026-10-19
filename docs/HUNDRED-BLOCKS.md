> Version 0.8.0: see WORKFLOWS-0.8.md for six/eight timestamp fields and complete working-dataset tables.

> Version 0.7.0 changes the user workflow and visible fields. See WORKFLOWS-0.7.md and README.md; older in-sheet examples below describe the retained calculation engine.

# Hundred-Block Builder and Randomize Block Addresses

Version 0.4.0. Both buttons appear under CAP Index > Geography with native Office icons. They insert eight fields beside the mapped address column in the current worksheet; existing fields shift right. Original address values and records remain intact.

## Select and run

Select a rectangular range including its headers and data. For an Excel Table, include the header and all data rows but exclude totals. Map the Address field and optionally City/State. City and State are used to make a geographic block key, so the same street in different cities is not merged.

- **Hundred-Block Builder:** converts addresses to hundred-block numbers and labels, without generating fictional addresses.
- **Randomize Block Addresses:** also fills synthetic-address fields for agency hundred-block inputs. Choose an optional stable IncidentID and a repeatable integer seed. The default seed is 20261013; choosing the same seed and ID/row inputs reproduces values. Changing the seed reshuffles the generated numbers. A pseudorandom generator is used; it is not a security or anonymization mechanism.

The user explicitly chose the convention that an address ending in 00 means hundred-block input when using this tool. A numeric 1200 is therefore treated as the 1200–1299 range, not proof of a building at house number 1200. Apply the randomizer to agency-masked records. The source string alone cannot distinguish a genuine house number ending in 00 from a masked block.

## Inserted fields

| Field | Meaning |
|---|---|
| CAP_HundredBlockNumber | Numeric block base: floor(house number/100)*100 |
| CAP_HundredBlock | Street-qualified label, such as 1200 BLOCK MAIN ST |
| CAP_BlockKey | Length-prefixed state/city/block key; blank when city/state are missing |
| CAP_AddressPrecision | Street number, Hundred block (00), Hundred block (XX), Hundred block (label), or Unknown |
| CAP_SyntheticAddress | Generated street address for masked records only |
| CAP_SyntheticFlag | TRUE for generated addresses; FALSE otherwise |
| CAP_RandomSeed | Seed used for that generated address |
| CAP_BlockStatus | Parsing outcome, geographic-key gaps and synthetic-location warning |

Examples:

- 1234 MAIN ST → block 1200; source precision Street number; synthetic fields blank.
- 1299 MAIN ST → block 1200.
- 1200 MAIN ST → block 1200; randomizer chooses any integer from 1200 through 1299, including both endpoints.
- 12XX MAIN ST → block 1200; same randomizer range.
- 1200 BLOCK OF MAIN ST / 1200 BLK MAIN ST → block 1200.
- XX MAIN ST → zero hundred block; synthetic number can be 0 through 99.

Unit suffixes are removed from the block street label, not from the original address. Ordinal and common numeric street names are supported. Intersections, PO boxes, fractional/range house numbers, malformed strings and explicit block labels that do not lie on a hundred boundary are flagged; the tool does not invent an address for those cases. Street abbreviation/direction aliases are not standardized.

## Incident and offense consistency

The synthetic number is a deterministic function of seed, incident identity and geographic block key. All offense rows with the same valid IncidentID and block receive the same synthetic address. Use an incident key unique across agency/year as appropriate. Different incidents may receive the same generated number by chance, which is expected in a 100-number range and does not make them duplicates.

Without a valid IncidentID, the generator uses the source worksheet row number. Sorting or moving source records can then change the generated value on a subsequent run. Missing-ID fallback counts appear in the completion summary. The tool retains all rows and makes no changes to incident/offense classifications or duplicate flags.

Reruns refresh the same eight marked columns. Keep their headers unchanged. The ordinary builder clears generated values because it is the non-randomizing action; use the randomizer to retain/regenerate synthetic values. Tables reject conflicting existing field names rather than silently renaming or replacing unrelated fields. Prior tool versions remain available for rollback.

## Analytical accuracy

The synthetic address is a simulated within-block house number, not a recovered incident address, a verified building, a street centroid or a uniform spatial sample. Parity/street side and true distance along the street are not known from a masked hundred block. A seed supports repeatability, not accuracy. Original masked precision remains visible and the geographic block key is unchanged, so hundred-block concentration analysis does not depend on the random draw.

If CAP_SyntheticAddress is selected in Census Geocoder, its Review Flag explicitly identifies a synthetic hundred-block estimate even when Census returns Exact. When the original address is mapped and its adjacent generated CAP_AddressPrecision indicates a masked block, Census results carry a masked-block approximate-location warning. Preserve those generated headers and flags in shared analysis. The matched address/coordinates do not establish the incident location. Census itself uses interpolated address-range coordinates; ranges can include numbers for structures that do not exist. [Census methodology](https://www.census.gov/programs-surveys/geography/technical-documentation/complete-technical-documentation/census-geocoder.html)

No live addresses are sent during the new-tool tests. Validation covers in-sheet/table insertion, 00/XX and block labels, number/label/key outputs, preservation of originals, seeded repeatability, repeated offense records, refresh behavior, synthetic warnings through Census, and 1,000 generated suffix checks including 00 and 99. This remains an unsigned pilot; visual ribbon acceptance and company deployment remain pending.
