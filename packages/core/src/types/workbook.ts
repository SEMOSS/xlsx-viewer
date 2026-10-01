import type { Person } from './comment';
import type { MediaItem } from './drawing';
import type { Worksheet } from './sheet';
import type { StyleSheet } from './style';
import type { Theme } from './theme';
import type { XlsxWarning } from './warnings';

export interface DefinedName {
	name: string;
	/** Formula / reference in display form (no leading `=`), e.g. `Sheet1!$A$1:$B$10`. */
	formula: string;
	/** Sheet index (into `workbook.sheets`) for sheet-scoped names; omitted for workbook scope. */
	localSheet?: number;
	hidden?: boolean;
	comment?: string;
	/** Built-in names such as `_xlnm.Print_Area`, `_xlnm._FilterDatabase`. */
	builtin?: boolean;
}

export interface DocumentProperties {
	title?: string;
	subject?: string;
	creator?: string;
	keywords?: string;
	description?: string;
	lastModifiedBy?: string;
	category?: string;
	/** ISO 8601. */
	created?: string;
	/** ISO 8601. */
	modified?: string;
	company?: string;
	manager?: string;
	application?: string;
	appVersion?: string;
	/** Custom document properties (`docProps/custom.xml`). */
	custom?: Array<{ name: string; value: string | number | boolean; type?: string }>;
}

export interface WorkbookView {
	/** Index of the active sheet (into `workbook.sheets`). */
	activeTab: number;
	firstSheet?: number;
	showHorizontalScroll?: boolean;
	showVerticalScroll?: boolean;
	showSheetTabs?: boolean;
	tabRatio?: number;
}

export interface CalcProperties {
	calcId?: number;
	calcMode?: 'auto' | 'autoNoTable' | 'manual';
	fullCalcOnLoad?: boolean;
	iterate?: boolean;
	iterateCount?: number;
	iterateDelta?: number;
	refMode?: 'A1' | 'R1C1';
}

export interface WorkbookProtection {
	lockStructure?: boolean;
	lockWindows?: boolean;
	workbookPassword?: string;
	workbookAlgorithmName?: string;
	workbookHashValue?: string;
	workbookSaltValue?: string;
	workbookSpinCount?: number;
}

/** The complete in-memory workbook. */
export interface Workbook {
	sheets: Worksheet[];
	definedNames: DefinedName[];
	styles: StyleSheet;
	theme: Theme;
	properties: DocumentProperties;
	view: WorkbookView;
	/** Dates use the 1904 date system (Mac Excel legacy). */
	date1904: boolean;
	calc: CalcProperties;
	/** Binary parts (images) keyed by package path. */
	media: Map<string, MediaItem>;
	/** Authors of threaded comments. */
	persons: Person[];
	protection?: WorkbookProtection;
	/** Approximations and unsupported features found while loading. */
	warnings: XlsxWarning[];
	/** The file was saved with macros (`.xlsm`); the VBA project is preserved. */
	hasMacros?: boolean;
	/**
	 * The file was encrypted with a password (File > Info > Protect Workbook >
	 * Encrypt with Password). The password itself is not kept: pass it to
	 * `writeWorkbook` again to save an encrypted file.
	 */
	encrypted?: boolean;
	/**
	 * Opaque handle to the source package, used by the writer to copy
	 * untouched parts verbatim. Absent for workbooks created in memory.
	 */
	sourcePackage?: unknown;
	/** Workbook-level parts changed (workbook.xml must be regenerated). */
	dirty?: boolean;
}
