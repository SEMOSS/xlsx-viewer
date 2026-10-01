import { createContext, useCallback, useContext } from 'react';

/** Dialogs and host-level actions the viewer shell provides to chrome components. */
export type DialogKind =
	| 'formatCells'
	| 'find'
	| 'replace'
	| 'goTo'
	| 'rowHeight'
	| 'columnWidth'
	| 'renameSheet'
	| 'insertCells'
	| 'deleteCells'
	| 'hyperlink'
	| 'note'
	| 'nameManager'
	| 'defineName'
	| 'dataValidation'
	| 'sort'
	| 'insertFunction'
	| 'zoom'
	| 'conditionalRule'
	| 'manageRules'
	| 'unhideSheet'
	| 'properties'
	| 'removeDuplicates'
	| 'textToColumns'
	| 'pageSetup'
	| 'alert'
	| 'protectSheet'
	| 'protectWorkbook'
	| 'unprotect'
	| 'validationAlert';

export interface ShellActions {
	openDialog: (kind: DialogKind, payload?: unknown) => void;
	/** Save via the host (`onSave` / `onContentChange`) or download. */
	save: () => void;
	/** Download in a format. */
	download: (format: 'xlsx' | 'csv' | 'markdown' | 'json' | 'html') => void;
	/** Ask the host to open another file, when supported. */
	openFile?: () => void;
	newWorkbook?: () => void;
	print: () => void;
	notify: (message: string) => void;
	toggleTheme?: () => void;
	themeMode: 'light' | 'dark';
	/** Give keyboard focus back to the grid. */
	focusGrid?: () => void;
}

export const ShellContext = createContext<ShellActions | null>(null);

export function useShell(): ShellActions {
	const ctx = useContext(ShellContext);
	if (!ctx) {
		throw new Error('useShell must be used inside <ExcelViewer>');
	}
	return ctx;
}

/**
 * Leave a chrome control (Name Box, font size, ...): keyboard focus goes back
 * to the grid, as in Excel. Outside a viewer the control is just blurred.
 */
export function useReturnToGrid(): (from?: HTMLElement | null) => void {
	const shell = useContext(ShellContext);
	return useCallback(
		(from?: HTMLElement | null) => {
			if (shell?.focusGrid) {
				shell.focusGrid();
			} else {
				from?.blur();
			}
		},
		[shell],
	);
}
