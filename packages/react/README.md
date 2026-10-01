# xlsx-react-viewer

`<ExcelViewer>`: a React 18/19 viewer and editor for Excel workbooks with an Office-style ribbon, formula
bar, sheet tabs and status bar. Built on `xlsx-viewer-core`; everything runs in the browser.

```tsx
import { ExcelViewer, type ExcelViewerHandle } from 'xlsx-react-viewer';
import 'xlsx-react-viewer/styles';

function Editor({ bytes }: { bytes: ArrayBuffer }) {
	const ref = useRef<ExcelViewerHandle>(null);
	return (
		<ExcelViewer
			ref={ref}
			content={bytes}
			fileName='Budget.xlsx'
			theme='light'
			onSave={async (saved, name) => upload(saved, name)}
			onDirtyChange={(dirty) => setUnsaved(dirty)}
			style={{ height: '100%' }}
		/>
	);
}
```

## Props

| Prop                                   | Description                                                                                       |
| -------------------------------------- | ------------------------------------------------------------------------------------------------- |
| `content`                              | Workbook bytes (`.xlsx`, `.xlsm`, `.xltx`, `.xls`). Omit for a blank workbook.                    |
| `workbook`                             | An already parsed `Workbook` (takes precedence over `content`).                                   |
| `fileName`                             | Title bar text and download name. Saves always use an Open XML name (`.xls` becomes `.xlsx`).    |
| `password`                             | Password for an encrypted file; without it the viewer prompts. Saves keep the encryption.         |
| `canEdit`                              | `false` for a read-only viewer (default `true`).                                                  |
| `theme`, `onThemeModeChange`           | `'light'`, `'dark'` or a `ViewerTheme` (see `jadeLightTheme` / `jadeDarkTheme`).                  |
| `onSave`                               | Called with the serialized workbook on Save (Ctrl+S). Without it, Save downloads the file.        |
| `onContentChange`, `contentChangeDelayMs` | Fresh bytes after each committed change (debounced).                                           |
| `onDirtyChange`, `onActiveSheetChange`, `onSelectionChange`, `onError` | State callbacks.                                                  |
| `onOpenFile`, `onNewWorkbook`          | Show File > Open / New and route them to the host.                                               |
| `locale`                               | UI language: `'en'` (default), `'de'`, `'fr'`, `'es'` or `'zh-CN'`. Switching it re-renders the UI without reloading the workbook. |
| `translations`                         | Per-key UI string overrides applied on top of `locale` (English is the final fallback).          |
| `authorName`                           | Author recorded on new notes.                                                                     |
| `showRibbon`, `showFormulaBar`, `showSheetTabs`, `showStatusBar`, `showTitleBar` | Toggle chrome.                                                |

The `ref` handle offers `getContent()`, `getWorkbook()`, `save()`, `undo()`, `redo()`, `recalculate()`,
`select(reference)`, `setActiveSheet(index)` and `focus()`.

## Languages

The UI ships in English, German, French, Spanish and Simplified Chinese, worded with Microsoft
Office's own terminology for each language. Pick one with `locale` and override single strings with
`translations` (keys are listed in `englishTranslations`):

```tsx
<ExcelViewer content={bytes} locale='de' translations={{ 'app.name': 'Tabellen' }} />
```

`BUILT_IN_LOCALES` lists the codes with their native names for a language picker, and the
dictionaries are exported as `englishTranslations`, `germanTranslations`, `frenchTranslations`,
`spanishTranslations` and `simplifiedChineseTranslations` (each a complete `TranslationDictionary`;
`BUILT_IN_DICTIONARIES` maps codes to them). Numbers in the chrome, such as the zoom percentage and
the status-bar statistics, follow the locale; cell values keep the workbook's number formats.
Messages produced by `xlsx-viewer-core` (for example some validation errors) are in English.

## What the UI covers

- **Grid:** virtualized DOM rendering with frozen panes, merged cells, text overflow, conditional
  formatting overlays, sparklines, tables, outlines, drawings and charts; keyboard navigation and
  shortcuts as in Excel (Ctrl+arrows, Shift selection, F2, F4, Ctrl+Enter, Alt+Enter, Ctrl+;, ...).
- **Ribbon:** Home, Insert, Page Layout, Formulas, Data, Review and View tabs, plus the contextual Table
  Design tab, with galleries for cell styles, table styles and conditional formatting.
- **Dialogs:** Format Cells, Insert Function, Name Manager, Data Validation, Sort, Hyperlink, Find and
  Replace, Go To, Page Setup, Protect Sheet / Workbook, conditional formatting rules, Text to Columns,
  Remove Duplicates and more.
- **File menu:** save, export (CSV, Markdown, JSON, HTML), print with preview, document properties and
  password encryption.

## Styling

The stylesheet is self-contained (Tailwind utilities without a global reset, scoped to the viewer). Colors
come from CSS variables derived from the theme, so a custom `ViewerTheme` restyles the chrome, the grid
headers and the selection.

## License

MIT
