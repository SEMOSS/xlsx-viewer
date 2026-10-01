# xlsx-viewer

**Open, render, edit, calculate and save Excel workbooks entirely in the browser (or Node.js).**
One TypeScript engine plus a React viewer/editor. No servers, no Excel install, files never leave the machine.

```
packages/
  core/    xlsx-viewer-core    Parse, edit, calculate, serialize, print and convert workbooks
  react/   xlsx-react-viewer   <ExcelViewer> React viewer / editor with an Office-style ribbon
demos/
  demo-react/                  Vite + React demo app (http://localhost:4183)
```

## Feature tour

|                   |                                                                                                                                                                                                                                                                                                                                                       |
| ----------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Files**         | `.xlsx`, `.xlsm` (macros preserved), `.xltx`/`.xltm`, OOXML Strict, legacy `.xls` (Excel 97-2003, imported and saved as `.xlsx`), password-encrypted workbooks (open and save, ECMA-376 Agile / Standard encryption).                                                                                                                                  |
| **Rendering**     | Fonts, fills (solid, pattern, gradient), borders, alignment, rotation, wrap, shrink to fit, merged cells, text overflow, number formats (dates, fractions, accounting, conditional sections, colors), frozen panes, hidden rows and columns, zoom, gridlines.                                                                                           |
| **Rich content**  | Tables with all 60 built-in styles, conditional formatting (cell rules, data bars, color scales, icon sets), sparklines, charts (bar, column, line, area, pie, doughnut, scatter, radar, bubble, stock, combo), pictures, shapes, notes, threaded-comment markers, hyperlinks, outline groups.                                                         |
| **Formulas**      | A dependency-tracking calculation engine with about 500 worksheet functions (math, statistics, lookup, text, dates, financial, engineering, database, logical, LAMBDA/LET), dynamic arrays with spilling, structured references, defined names, circular-reference detection and manual calculation.                                                  |
| **Editing**       | In-cell and formula-bar editing with function autocomplete, argument hints, F4 reference toggling and point mode; AutoComplete for text; fill handle and series; copy / paste (Excel-compatible HTML and text); insert / delete rows, columns and cells; sort and filter; find and replace; undo / redo.                                                 |
| **Formatting**    | Home ribbon (font, alignment, number formats, borders, fills), Format Cells dialog, cell styles, Format as Table and Table Design, conditional formatting galleries and rules manager, format painter, column widths and row heights with AutoFit.                                                                                                     |
| **Data**          | Data validation (lists, number, date, length and custom rules with input messages and error alerts), remove duplicates, text to columns, grouping and outlines, named ranges.                                                                                                                                                                          |
| **Review**        | Notes (new, edit, show / hide, navigate), sheet and workbook protection enforced like Excel (locked cells, allowed actions, Excel-compatible password hashes), Trace Precedents / Dependents.                                                                                                                                                         |
| **Printing**      | Page Layout tab and Page Setup dialog (orientation, paper size, margins, scaling, fit to pages, print area, print titles, page breaks, headers and footers), File > Print with a page preview, real printing via the browser (and so "Save as PDF").                                                                                                    |
| **Exporting**     | `.xlsx`, CSV, Markdown, JSON and HTML.                                                                                                                                                                                                                                                                                                                 |
| **Round trip**    | Unedited sheets and parts are written back byte-for-byte; markup the engine does not model is preserved; styles and shared strings stay stable, so saved files diff cleanly against the originals.                                                                                                                                                    |

## Quick start

### React

```tsx
import { ExcelViewer } from 'xlsx-react-viewer';
import 'xlsx-react-viewer/styles';

export function App({ bytes }: { bytes: ArrayBuffer }) {
	return (
		<ExcelViewer
			content={bytes}
			fileName='Budget.xlsx'
			onSave={(saved, name) => upload(saved, name)} // the edited .xlsx
			style={{ height: '100vh' }}
		/>
	);
}
```

Useful props: `canEdit`, `theme` (`'light'`, `'dark'` or a theme object), `locale` (`'en'`, `'de'`, `'fr'`,
`'es'`, `'zh-CN'`), `password`, `onContentChange`, `onDirtyChange`, `onSelectionChange`, `translations`,
`authorName`, and `showRibbon` / `showFormulaBar` /
`showSheetTabs` / `showStatusBar` / `showTitleBar`. A `ref` exposes `getContent()`, `save()`, `undo()`,
`redo()`, `recalculate()`, `select('Sheet2!B4')`, `setActiveSheet(i)` and `focus()`. See
[packages/react/README.md](packages/react/README.md).

### Headless (Node.js, workers)

```ts
import { CalcEngine, readWorkbook, sheetToCsv, writeWorkbook } from 'xlsx-viewer-core';

const workbook = await readWorkbook(bytes); // .xlsx, .xlsm or .xls; pass { password } for encrypted files
const calc = new CalcEngine(workbook);
calc.rebuild(); // trusts saved values, computes formulas that have none
console.log(sheetToCsv(workbook, workbook.sheets[0]));
const saved = await writeWorkbook(workbook); // Uint8Array of a valid .xlsx
```

See [packages/core/README.md](packages/core/README.md) for editing, calculation, printing and conversion.

## Development

The package manager is **pnpm** (workspaces). From the repository root:

```bash
pnpm install     # install workspace dependencies
pnpm test        # vitest for core and react
pnpm e2e         # Playwright end-to-end tests against the demo
pnpm typecheck   # tsc --noEmit for every package and the e2e specs
pnpm build       # core -> react -> demo
pnpm demo        # demo dev server on http://localhost:4183
```

The demo aliases both packages to their sources, so edits under `packages/*/src` show up on reload without
a build. [CLAUDE.md](CLAUDE.md) describes the architecture, data-model rules and coding conventions.

The end-to-end tests in `e2e/` start the demo themselves (or reuse one already running on port 4183) and
drive it in the installed Google Chrome, so no browser download is needed. To use Playwright's own
Chromium instead, run `pnpm exec playwright install chromium` once and set `XLSX_E2E_CHANNEL=chromium`.
`pnpm e2e:headed` shows the browser and `pnpm e2e:ui` opens Playwright's UI mode.

## Status and known limitations

- Excel itself was not available while this was built: round-trip, encryption and `.xls` behaviour is
  verified against the specifications, real files where available, and extensive tests, but files written
  here have not yet been opened in desktop Excel by an automated check.
- Not yet supported: PivotTables (preserved, not rendered), the newer chart family (treemap, sunburst,
  waterfall, histogram, box and whisker, funnel, map; preserved and shown as placeholders), error bars and
  data tables in charts, `.xlsb` and `.ods` files, encrypted `.xls` files and real-time collaboration.
  The UI is translated into German, French, Spanish and Simplified Chinese; messages that come from the
  core engine are still English.
- Fonts are rendered with the browser's fonts; when a workbook font is not installed a metrically similar
  fallback is used, so text widths can differ slightly from Excel.

## License

MIT
