# Hundred-Block Builder and Randomize Block Addresses

Use CAP Index > Geography. Select headers and data, open the tool, map Address and optional City/State, review the selected range, and Run. Each run creates a complete working dataset on a new named sheet; the original sheet is unchanged. Only the selected fields and rows are copied.

## Outputs

| Tool | New sheet | Generated fields |
|---|---|---|
| Hundred-Block Builder | Hundred_Blocks | CAP_HundredBlock; Review Status |
| Randomize Block Addresses | Randomized_Addresses | CAP_HundredBlock; CAP_SyntheticAddress; Review Status |

Outputs are native Excel Tables starting at row 6. Existing matching generated headers are reused in the new dataset. Seed and other run information appear once in the notes rather than on every row. City/State qualify locations internally; no extra block-key or precision columns are required.

## Address rules

- 1234 MAIN ST and 1299 MAIN ST become 1200 BLOCK MAIN ST.
- In Randomize Block Addresses, 1200 MAIN ST, 12XX MAIN ST, and explicit 1200 BLOCK/BLK labels represent the inclusive 1200–1299 range.
- A genuine address ending in 00 cannot be distinguished from an agency mask by text alone. Use the randomizer for known agency block data.
- Unit suffixes are removed from the block label while original source text is retained. Exact nonmasked addresses are retained by the randomizer.
- Intersections, PO boxes, fractional/range house numbers, malformed strings and invalid block boundaries are flagged for review. Street abbreviation and directional aliases are not standardized.

## Repeatability and record grain

Choose an optional stable IncidentID and integer seed. The same seed, incident identity and qualified block reproduce the synthetic address, keeping offense rows for one incident/block consistent. Qualify IDs across agencies/years. Different incidents can receive the same number by chance.

Without a usable ID, source worksheet row identity is used; sorting or moving rows can change a later draw. Review Status identifies synthetic estimates, and run notes record the seed. Randomization is not a security or anonymization mechanism. No offense rows are deduplicated or reclassified.

## Analytical accuracy

The synthetic address is a simulated within-block house number, not a recovered incident address, a verified building, a street centroid or a uniform spatial sample. Parity/street side and true distance along the street are not known from a masked hundred block. A seed supports repeatability, not accuracy. Original masked precision remains visible and the geographic block key is unchanged, so hundred-block concentration analysis does not depend on the random draw.

If CAP_SyntheticAddress is selected in Census Geocoder, its Review Flag explicitly identifies a synthetic hundred-block estimate even when Census returns Exact. Preserve Review Status and Review Flag when sharing synthetic or masked-location results. The matched address/coordinates do not establish the incident location. Census itself uses interpolated address-range coordinates; ranges can include numbers for structures that do not exist. [Census methodology](https://www.census.gov/programs-surveys/geography/technical-documentation/complete-technical-documentation/census-geocoder.html)
