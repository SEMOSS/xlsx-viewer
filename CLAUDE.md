# CLAUDE.md

Guidance for Claude Code (and other agents) working in this repository.

## What this is

`xlsx-viewer`: open, render, edit, calculate and save Excel workbooks entirely in
the browser (or Node.js).

```
packages/
  core/    xlsx-viewer-core    Parse, edit, calculate, serialize, convert .xlsx (framework-agnostic)
  react/   xlsx-react-viewer   <ExcelViewer> React viewer/editor component
demos/
  demo-react/                  Vite + React demo app (port 4183)
```

## Commands

The package manager is **pnpm** (workspaces). Run from the repo root unless noted.

```bash
pnpm install          # install workspace dependencies
pnpm test             # vitest for core and react
pnpm e2e              # Playwright e2e specs (e2e/) against the demo, in installed Chrome
pnpm typecheck        # tsc --noEmit for every package and the e2e specs
pnpm build            # core -> react -> demo
pnpm demo             # React demo dev server on http://localhost:4183

# inside a package (packages/core or packages/react)
npx vitest run src/<area>         # focused tests
npx tsc --noEmit -p tsconfig.json # typecheck
```

The demo aliases `xlsx-viewer-core` and `xlsx-react-viewer` to their **source**, so
edits under `packages/*/src` are live on reload without a build.

## React architecture (`packages/react/src`)

- `engine/`: pure (React-free) logic: geometry and hit testing, viewport layout (fills, texts, borders,
  overlays), selection and navigation, keymap, clipboard, sparklines, print rendering models.
- `viewer/state/`: `ViewerStore` (workbook, calc engine, undo history, UI state) consumed through
  `useSyncExternalStore`. Every document change goes through `store.transact(label, fn)`, which installs
  the sheet-protection guard, rolls back on errors, re-fits automatic row heights and recalculates.
- `viewer/commands/`: user actions on top of core edit operations (one file per area, with tests).
- `viewer/components/`: grid, chrome (title bar, formula bar, sheet tabs, status bar, backstage), ribbon
  tabs and groups, dialogs (registered in `dialogs/registry.ts`), drawings, charts, print preview.
- Keyboard focus (`viewer/focus.ts`): between edits the grid's hidden focus sink owns the keyboard, as in
  Excel. Grid presses are cancelled so the browser cannot move focus off it; clicks on toolbar buttons and
  empty chrome keep (or bring back) focus to it; text fields, sliders and drag sources take focus as usual
  and return it with `useReturnToGrid()`. Dialogs, popovers and the backstage take focus while open and
  give it back when they close.

## Core architecture (`packages/core/src`)

| Directory      | Responsibility                                                                    |
| -------------- | --------------------------------------------------------------------------------- |
| `types/`       | The data model (`Workbook`, `Worksheet`, `Cell`, styles, drawings, charts, ...)   |
| `address/`     | A1 notation, ranges, `cellKey`, sheet-name quoting                                |
| `values/`      | Error singletons, date serials, numeric-text parsing                              |
| `xml/`         | Order-preserving XML parser / serializer and DOM helpers                          |
| `package/`     | OPC packages, relationships, content types; `zip/` native ZIP (JSZip fallback); `cfb/` compound files; `crypto/` encryption |
| `color/`       | Theme / indexed / tint / DrawingML color resolution                               |
| `styles/`      | Built-in number formats, style resolution, append-only `StylePool`, column widths |
| `sheet/`       | Worksheet / workbook factories, cell accessors, outlines, hyperlink resolution    |
| `numfmt/`      | Number format engine (format codes -> display text)                               |
| `formula/`     | Tokenizer, parser, AST, evaluator, function library                               |
| `calc/`        | Dependency graph and recalculation (`CalcEngine`)                                 |
| `conditional/` | Conditional formatting evaluation                                                 |
| `validation/`  | Data validation                                                                   |
| `charts/`      | Chart part parsing / writing and data resolution                                  |
| `read/`        | Load pipeline: bytes -> `Workbook` (`xls/` reads legacy BIFF8 `.xls`)             |
| `write/`       | Save pipeline: `Workbook` -> bytes                                                |
| `edit/`        | Editing operations in `EditTransaction`s (insert/delete, fill, sort, tables, protection, ...) |
| `print/`       | Pagination, page setup, headers and footers                                       |
| `convert/`     | Markdown / CSV / JSON / HTML converters                                           |

### Data model rules

- **Coordinates are 0-based** everywhere in the model (row 0 = Excel row 1, col 0 = A).
  Ranges are inclusive `{ top, left, bottom, right }`, normalised.
- Cells live in `sheet.cells: Map<number, Cell>` keyed by `cellKey(row, col)`.
  **Cell objects are immutable**: replace them, never mutate in place.
- A cell value is `number | string | boolean | CellError | null`. Dates are serial
  numbers. Errors are frozen singletons from `values/errors.ts` (`ERRORS.DIV0`).
- Formulas are stored **without** the leading `=`, in display form: file prefixes
  (`_xlfn.`, `_xlws.`, `_xlpm.`) are stripped on load and restored on save; shared
  formulas are expanded so every formula cell has its own text.
- Styles follow `styles.xml` and are **append-only**: never reorder or mutate existing
  fonts / fills / borders / xfs / dxfs; derive new ones with `StylePool`.

### Round-trip fidelity (load -> save)

- `workbook.sourcePackage` is the `OpcPackage` the workbook was read from. The writer
  starts from a copy of it, so parts the engine does not understand are preserved.
- Change tracking: the reader sets `dirty = false` on the workbook, each sheet,
  `drawingsDirty` / `commentsDirty` on sheets and `dirty` on tables. A sheet with
  `dirty === false` is written back byte-for-byte; anything else is regenerated from
  the model, using `sheet.source.root` to keep element order and unknown markup.
- Original style entries keep their raw XML (`styles.raw`) and are re-emitted
  verbatim; only new entries are serialized from the model.
- Shared-string indices of the original file stay stable (new strings are appended).

## Conventions

- **TypeScript strict.** Avoid `any`; narrow with the discriminants (`kind`, `type`).
- **Keep source files at or under ~300-400 lines.** Split by responsibility rather than
  growing a file. Every directory has an `index.ts` barrel.
- **Pure logic belongs in core.** The React package maps framework-neutral results
  (style descriptors, layout geometry, formatted text) onto JSX. If logic does not need
  React, put it in `packages/core` or `packages/react/src/engine` (no React imports).
- **Tests are colocated** (`*.test.ts`), run with vitest. Every module ships tests.
- **End-to-end specs** live in `e2e/` (Playwright, `playwright.config.ts`) and drive the demo the way a
  user does: accessible names, the demo's `#file-input`, and the `data-cell="B2"` attribute on rendered
  cell text. Shared helpers are in `e2e/support/viewer.ts`. Beware Excel's AutoComplete when a spec types
  text into a column that already has entries starting with the same letters.
- **No em-dashes** (U+2014) in code, comments, docs or UI copy; use `:`, `,`, `;`,
  parentheses or a spaced hyphen.
- File names: kebab-case for modules, PascalCase for React components and classes.
- **UI text is localized** (en, de, fr, es, zh-CN). Never hard-code user-visible strings in the
  React package: add a key to the area file under `packages/react/src/i18n/locales/en/` and the same
  key to every other locale (the dictionaries are typed against English, so a missing key fails the
  typecheck), then use `useT()` in components or `store.t` in commands. Excel syntax (function names,
  number-format codes, references, shortcuts) stays untranslated.
- Formatting: tabs, single quotes, trailing commas, 100-column lines (`.prettierrc.json`).
- No new dependencies without discussion: the core depends only on `jszip` (used as the fallback where
  the platform lacks `CompressionStream('deflate-raw')`).

## Offline install note

This repo was bootstrapped in a sandbox without registry access, so `package.json`
pins a few transitive versions under `pnpm.overrides` that were available in the
local pnpm store. They are harmless; remove them once you have normal network access
if you want the latest transitive versions.
