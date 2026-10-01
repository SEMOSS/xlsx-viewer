# xlsx-viewer-core

The framework-agnostic engine behind `xlsx-react-viewer`. It reads, edits, calculates, prints and writes
Excel workbooks in browsers, Node.js, Bun and workers. Its only dependency is `jszip`.

## Read and write

```ts
import { readWorkbook, writeWorkbook, XlsxPasswordError } from 'xlsx-viewer-core';

try {
	const workbook = await readWorkbook(bytes, { password: 'optional' }); // .xlsx, .xlsm, .xltx, .xls
	const out = await writeWorkbook(workbook); // Uint8Array; pass { password } to encrypt
} catch (e) {
	if (e instanceof XlsxPasswordError) {
		// e.reason: 'required' | 'incorrect' | 'unsupported'
	}
}
```

- Saving starts from the original package: sheets with `dirty === false` and parts the model does not
  understand are copied byte-for-byte; edited sheets are regenerated using the original element order.
- `workbook.warnings` lists anything that could not be read (never fatal).
- Legacy `.xls` files are converted to the same model (`readXlsWorkbook` is used automatically).

## The model

- Coordinates are 0-based; ranges are inclusive `{ top, left, bottom, right }`.
- `sheet.cells` is a `Map` keyed by `cellKey(row, col)`; cells are immutable objects
  (`{ value, formula?, style?, ... }`); a value is `number | string | boolean | CellError | null`.
- Formulas are stored without the leading `=` in display form (`XLOOKUP(...)`, not `_xlfn.XLOOKUP(...)`).
- Styles live in `workbook.styles` and are append-only; derive new formats with `StylePool`.

## Editing

Edits run in an `EditTransaction`, which records what changes so it can be undone:

```ts
import { applyStylePatch, EditTransaction, insertRows, parseRange, setCellInput, sortRange } from 'xlsx-viewer-core';

const tx = new EditTransaction(workbook, 'Update');
setCellInput(tx, workbook, 0, 0, 0, 'Revenue'); // typed input: numbers, dates, booleans, "=formulas"
setCellInput(tx, workbook, 0, 1, 0, '=SUM(B2:B10)');
applyStylePatch(tx, workbook, 0, [parseRange('A1:D1')!], { font: { bold: true }, fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFDDEBF7' } } });
insertRows(tx, workbook, 0, 5, 2); // formulas, merges, tables, charts and names follow
sortRange(tx, workbook, 0, parseRange('A2:D10')!, [{ index: 1, descending: true }], { hasHeader: false });
const entry = tx.commit();
entry.undo();
entry.redo();
```

Other operations include fill series, merge, clear, column widths / row heights, freeze panes, sheets (add,
rename, move, duplicate, hide), tables (create, style, totals row, rename, convert to range), filters, cell
styles, conditional formats, data validation, names, hyperlinks, notes, outlines (group / ungroup),
protection, text to columns, remove duplicates, cell shifting and page setup.

## Calculation

```ts
import { CalcEngine } from 'xlsx-viewer-core';

const calc = new CalcEngine(workbook);
calc.rebuild(); // dependency graph; trusts saved values
calc.onCellsChanged([{ sheet: 0, row: 0, col: 1 }]); // after edits
calc.recalculateAll();
calc.evaluateFormula('SUMIFS(C:C,B:B,"Paris")', 0, 0, 0);
calc.getPrecedents({ sheet: 0, row: 4, col: 2 });
```

About 500 functions are implemented, including dynamic arrays (FILTER, SORT, UNIQUE, SEQUENCE, XLOOKUP,
GROUPBY, ...), LAMBDA and its helpers, and the statistical, financial and engineering libraries. A help catalog
(`getFunctionInfo`, `searchFunctions`, `listFunctions`) describes every function and its arguments.

## Formatting and display

```ts
import { formatValue, StyleResolver } from 'xlsx-viewer-core';

formatValue(45383.5, 'dddd, mmmm d, yyyy h:mm AM/PM'); // { text: 'Monday, April 1, 2024 12:00 PM', ... }
const style = new StyleResolver(workbook.styles, workbook.theme).resolve(cell.style);
```

## Printing

`paginateSheet(workbook, sheetIndex)` lays a sheet out on pages (print area, titles, scaling, fit to pages,
breaks, margins, headers and footers) and returns page blocks in points that a renderer can draw.

## Converting

`sheetToCsv`, `workbookToMarkdown`, `workbookToJson` and `sheetToHtml` turn workbooks into text formats
(values as displayed, formulas optional).

## License

MIT
