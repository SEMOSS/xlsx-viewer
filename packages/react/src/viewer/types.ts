import type { CellAddress, RangeAddress, Workbook } from 'xlsx-viewer-core';
import type { CSSProperties } from 'react';

import type { SupportedLocale, Translations } from '../i18n';
import type { ViewerTheme } from '../theme';

export interface ExcelViewerProps {
	/** Workbook bytes (`.xlsx` / `.xlsm`). Omit (and omit `workbook`) to start with a blank workbook. */
	content?: Uint8Array | ArrayBuffer | null;
	/** An already-parsed workbook (takes precedence over `content`). */
	workbook?: Workbook;
	/** Shown in the title bar and used for downloads. */
	fileName?: string;
	/**
	 * Password of an encrypted workbook, for hosts that know it: opening skips the
	 * password prompt. Read when `content` loads. As in Excel, saves are encrypted
	 * with the password the file was opened with until the user changes it.
	 */
	password?: string;
	/** Allow editing (default true). */
	canEdit?: boolean;
	/** Theme object, or `'light'` / `'dark'` for the built-in palettes. */
	theme?: ViewerTheme | 'light' | 'dark';
	/** Called when the user toggles light / dark mode in the viewer. */
	onThemeModeChange?: (mode: 'light' | 'dark') => void;
	/**
	 * Called on Save (Ctrl+S / File > Save) with the serialized workbook.
	 * Without it, Save downloads the file.
	 */
	onSave?: (bytes: Uint8Array, fileName: string) => void | Promise<void>;
	/** Called with fresh bytes after every committed change (debounced). Prefer `onSave` for large files. */
	onContentChange?: (bytes: Uint8Array) => void;
	/** Debounce for `onContentChange` in ms (default 800). */
	contentChangeDelayMs?: number;
	onDirtyChange?: (dirty: boolean) => void;
	onActiveSheetChange?: (index: number) => void;
	onSelectionChange?: (selection: {
		sheet: number;
		active: CellAddress;
		ranges: RangeAddress[];
	}) => void;
	/** Show File > Open and route it to the host. */
	onOpenFile?: () => void;
	/** Show File > New and route it to the host. */
	onNewWorkbook?: () => void;
	/** Called when loading fails. */
	onError?: (error: Error) => void;
	/**
	 * UI language: one of the built-in dictionaries (default `'en'`). Also used
	 * for numbers in the chrome (zoom, status-bar statistics). Changing it
	 * re-renders the UI without reloading the workbook.
	 */
	locale?: SupportedLocale;
	/** UI string overrides applied on top of the `locale` dictionary (English is the final fallback). */
	translations?: Translations;
	/** Author recorded on new notes. */
	authorName?: string;
	showRibbon?: boolean;
	showFormulaBar?: boolean;
	showSheetTabs?: boolean;
	showStatusBar?: boolean;
	showTitleBar?: boolean;
	className?: string;
	style?: CSSProperties;
}

/** Imperative API exposed through `ref`. */
export interface ExcelViewerHandle {
	/** Serialize the current workbook. */
	getContent: () => Promise<Uint8Array>;
	/** The live in-memory workbook (mutations outside the viewer are not tracked). */
	getWorkbook: () => Workbook | null;
	/** Trigger Save (calls `onSave` or downloads). */
	save: () => Promise<void>;
	undo: () => void;
	redo: () => void;
	/** Recalculate every formula. */
	recalculate: () => void;
	/** Select a reference (`B4`, `A1:C9`, `Sheet2!D3`, a defined name). Returns false when invalid. */
	select: (reference: string) => boolean;
	setActiveSheet: (index: number) => void;
	/** Give keyboard focus to the grid. */
	focus: () => void;
}
