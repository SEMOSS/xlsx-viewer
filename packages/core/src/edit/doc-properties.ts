import type { DocumentProperties, Workbook } from '../types/workbook';

/** Text document properties the Properties dialog edits (`docProps/core.xml` / `app.xml`). */
export const DOCUMENT_TEXT_PROPERTIES = [
	'title',
	'subject',
	'creator',
	'manager',
	'company',
	'category',
	'keywords',
	'description',
	'lastModifiedBy',
] as const;

export type DocumentTextProperty = (typeof DOCUMENT_TEXT_PROPERTIES)[number];

function clean(value: unknown): string | undefined {
	if (typeof value !== 'string') {
		return undefined;
	}
	const t = value.trim();
	return t === '' ? undefined : t;
}

/**
 * Update the text document properties (title, subject, author, ...). Only
 * the keys present in `patch` are considered; values are trimmed and empty
 * text removes the property. Dates and other fields are ignored. Not
 * undoable (like Excel). Marks the workbook dirty and returns `true` when
 * something changed.
 */
export function setDocumentProperties(
	workbook: Workbook,
	patch: Partial<DocumentProperties>,
): boolean {
	const current = workbook.properties;
	const next: DocumentProperties = { ...current };
	let changed = false;
	for (const key of DOCUMENT_TEXT_PROPERTIES) {
		if (!Object.prototype.hasOwnProperty.call(patch, key)) {
			continue;
		}
		const value = clean(patch[key]);
		if (clean(current[key]) === value) {
			continue;
		}
		changed = true;
		if (value === undefined) {
			delete next[key];
		} else {
			next[key] = value;
		}
	}
	if (changed) {
		workbook.properties = next;
		workbook.dirty = true;
	}
	return changed;
}

/** Counts shown on the Statistics tab of the Properties dialog. */
export interface WorkbookStatistics {
	/** Sheets of every kind, hidden ones included. */
	sheets: number;
	/** Cells with a value or a formula. */
	cells: number;
	/** Cells with a formula. */
	formulas: number;
}

/** Sheet, non-empty cell and formula counts of a workbook. */
export function workbookStatistics(workbook: Workbook): WorkbookStatistics {
	let cells = 0;
	let formulas = 0;
	for (const sheet of workbook.sheets) {
		for (const cell of sheet.cells.values()) {
			if (cell.formula) {
				formulas++;
				cells++;
			} else if (cell.value !== null && cell.value !== '') {
				cells++;
			}
		}
	}
	return { sheets: workbook.sheets.length, cells, formulas };
}
