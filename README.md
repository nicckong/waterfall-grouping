# Waterfall Grouping

A single-file browser tool for NPS driver tables. Upload an Excel file, group driver rows into
"Others", combine groups, sort within groups, and export the result back to Excel with formulas,
formatting and colours intact.

Everything runs inside your browser. The Excel file is never uploaded anywhere.

## Using it

1. Download the repository as a zip (green **Code** button, then **Download ZIP**) and unzip it.
2. Open `waterfall_grouping.html` in a current version of Chrome or Edge.
3. Choose your `.xlsx` file, or drag it onto the page.

That one HTML file is the whole tool. You can copy it anywhere; it does not need the other files in
this repository and does not need an internet connection.

### What it does

- **Group into Others:** tick drivers within a group (e.g. Product Fit and Product Range under
  Product) and they become one row, "Product - Others". Pro NPS and Det NPS are summed, Pro-Det NPS
  and chg are recalculated, and Rank is redone across all brands. Tied scores share a rank (2, 2, 4).
- **Combine groups:** tick two or more groups (e.g. Journey and Product) to merge them into one
  group, "Journey & Product", with all their drivers underneath.
- **Sort:** click a column heading to sort drivers within each group (best first, then reversed,
  then back to the original order). Groups stay in place, and an Others row stays last. The export
  uses the sorted order.
- **Export:** a new `.xlsx` with formulas kept live, column A merges adjusted, and chg columns
  coloured flat red/green (reversed for Det NPS, where a fall is good). Every sheet in the file is
  processed the same way.

### Expected file layout

- Three header rows: brand names, then periods (two periods plus a `chg` column group), then
  `Rank`, `Pro NPS`, `Det NPS`, `Pro-Det NPS` under each period.
- Row labels in the column to the left of the first data column.
- The hierarchy is read, in order of preference, from: merged group names in the column left of the
  labels; `=SUM(...)` formulas in parent rows; the numbers themselves (a parent equals the sum of the
  rows under it); bold parent labels.
- Any number of brands, any group and driver names. Files with or without formulas both work.

## If the page shows an error after copying

The Excel reader is stored compressed at the bottom of `waterfall_grouping.html`, in four numbered
parts with checksums. If a copy was cut short or changed (for example by pasting its text somewhere
with a size limit), the page says which part is missing or damaged. Copy the file itself (download,
email attachment or shared drive) rather than its text.

## Development

Needs Node.js 18+ and, for the tests, Python 3 with `openpyxl` (`pip install openpyxl`).

```bash
npm install
npm run build   # rebuilds waterfall_grouping.html from src/
npm test        # generates test workbooks and checks every scenario
```

| Path | What it is |
|---|---|
| `src/core.js` | Grouping logic: reads the layout, plans the new row order, rewrites formulas, merges and colours |
| `src/page.html` | The page: upload, tick boxes, table, sorting, export |
| `build/build.js` | Combines the page, the logic and ExcelJS into `waterfall_grouping.html` |
| `tests/make_fixtures.js` | Generates made-up test workbooks with the real layout |
| `tests/run_export.js` | Runs the grouping logic on a file from the command line |
| `tests/verify.py` | Independent checker: recalculates every formula and re-derives every number from the source |
| `tests/run_all.sh` | Runs all scenarios |

`waterfall_grouping.html` is a build output but is committed, so it can be downloaded directly.
After changing anything in `src/`, run `npm run build` and `npm test` and commit the rebuilt file.

## Licence

ExcelJS is bundled inside `waterfall_grouping.html` under the MIT licence; see
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
