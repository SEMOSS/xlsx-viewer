import type { Cell, ResolvedCellStyle } from 'xlsx-viewer-core';
import {
	cellKey,
	clearRanges,
	isCellError,
	isDateFormat,
	numberToText,
	serialToDateParts,
	setCellInput,
	translateFormula,
} from 'xlsx-viewer-core';

import type { Direction } from '../../engine/selection/navigation';
import type { EditState } from '../state/types';
import type { ViewerStore } from '../state/viewer-store';
import { advanceActive } from './selection-commands';
import { blockedByProtection } from './protection-commands';
import { checkEntry } from './validation-commands';
import { autoCompleteCandidate } from '../../engine/selection/autocomplete';
import { autoExpandTable } from 'xlsx-viewer-core';

function pad2(n: number): string {
	return n < 10 ? `0${n}` : String(n);
}

/** Text Excel shows when you start editing a cell (formula bar contents). */
export function editTextForCell(
	cell: Cell | undefined,
	style: ResolvedCellStyle,
	date1904: boolean,
): string {
	if (!cell) {
		return '';
	}
	if (cell.formula !== undefined) {
		return `=${cell.formula}`;
	}
	const v = cell.value;
	if (v === null) {
		return '';
	}
	if (typeof v === 'string') {
		return style.quotePrefix ? `'${v}` : v;
	}
	if (typeof v === 'boolean') {
		return v ? 'TRUE' : 'FALSE';
	}
	if (isCellError(v)) {
		return v.code;
	}
	const fmt = style.numFmt;
	if (fmt.includes('%') && !/^general$/i.test(fmt)) {
		return `${numberToText(v * 100)}%`;
	}
	if (isDateFormat(fmt) && v >= 0) {
		const p = serialToDateParts(v, date1904);
		const hasDate = Math.floor(v) !== 0;
		const hasTime = v % 1 !== 0;
		const date = `${p.month}/${p.day}/${p.year}`;
		const h12 = p.hour % 12 === 0 ? 12 : p.hour % 12;
		const time = `${h12}:${pad2(p.minute)}:${pad2(p.second)} ${p.hour < 12 ? 'AM' : 'PM'}`;
		if (hasDate && hasTime) {
			return `${date} ${time}`;
		}
		return hasDate ? date : time;
	}
	return numberToText(v);
}

/** Start editing the active cell. `initialText` replaces the content (typing over a cell). */
export function beginEdit(
	store: ViewerStore,
	opts: { initialText?: string; source?: EditState['source']; mode?: EditState['mode'] } = {},
): void {
	if (store.readOnly) {
		return;
	}
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const { row, col } = store.selectionOf().active;
	if (blockedByProtection(store, row, col)) {
		return;
	}
	const cell = sheet.cells.get(cellKey(row, col));
	const style = store.styles.resolve(cell?.style);
	const original = editTextForCell(cell, style, store.workbook.date1904);
	const text = opts.initialText ?? original;
	const suggestion =
		opts.initialText !== undefined && (opts.source ?? 'cell') === 'cell'
			? autoCompleteCandidate(sheet, row, col, text)
			: null;
	store.setUi({
		edit: {
			sheet: sheetIndex,
			row,
			col,
			text,
			caret: text.length,
			mode: opts.mode ?? (opts.initialText !== undefined ? 'enter' : 'edit'),
			source: opts.source ?? 'cell',
			original,
			suggestion: suggestion ?? undefined,
		},
	});
}

export function updateEditText(store: ViewerStore, text: string, caret: number): void {
	const edit = store.state.edit;
	if (!edit) {
		return;
	}
	store.setUi({ edit: { ...edit, text, caret, point: undefined, suggestion: undefined } });
}

/**
 * Typing in the cell: update the text and, when characters were added at
 * the end, offer Excel's AutoComplete from the column's entries.
 */
export function typeInCell(store: ViewerStore, text: string, caret: number): void {
	const edit = store.state.edit;
	if (!edit) {
		return;
	}
	const grew = text.length > edit.text.length && caret === text.length;
	const candidate = grew
		? autoCompleteCandidate(store.workbook.sheets[edit.sheet], edit.row, edit.col, text)
		: null;
	store.setUi({
		edit: { ...edit, text, caret, point: undefined, suggestion: candidate ?? undefined },
	});
}

export function cancelEdit(store: ViewerStore): void {
	if (store.state.edit) {
		store.setUi({ edit: null });
	}
}

export interface CommitOptions {
	/** Ctrl+Enter: put the entry in every selected cell. */
	fillSelection?: boolean;
	/** Where the active cell goes afterwards. */
	move?: Direction | 'none';
	/** The user already accepted a data validation warning for this entry. */
	skipValidation?: boolean;
}

/** Commit the editor text into the cell (or selection) and move on. Returns false if rejected. */
export function commitEdit(store: ViewerStore, opts: CommitOptions = {}): boolean {
	const current = store.state.edit;
	if (!current) {
		return true;
	}
	// An AutoComplete suggestion is accepted on commit, in the case of the existing entry.
	const edit = current.suggestion
		? { ...current, text: current.suggestion, suggestion: undefined }
		: current;
	const changed = edit.text !== edit.original;
	if (changed && !opts.skipValidation) {
		const failure = checkEntry(store, edit);
		if (failure?.style) {
			// Excel's alert: Stop retries or cancels, Warning asks to continue, Information informs.
			store.setUi({
				dialog: {
					kind: 'validationAlert',
					payload: {
						style: failure.style,
						title: failure.title ?? '',
						message: failure.message ?? '',
						move: opts.move ?? 'down',
						fillSelection: opts.fillSelection,
					},
				},
			});
			return false;
		}
	}
	if (changed) {
		const wb = store.workbook;
		const sel = store.selectionOf(edit.sheet);
		let rejected = false;
		const label = edit.text.startsWith('=')
			? store.t('undo.formula')
			: store.t('undo.typing', { text: edit.text.slice(0, 20) });
		store.transact(label, (tx) => {
			if (opts.fillSelection) {
				for (const range of sel.ranges) {
					for (let row = range.top; row <= range.bottom; row++) {
						for (let col = range.left; col <= range.right; col++) {
							const text = edit.text.startsWith('=')
								? `=${translateFormula(edit.text.slice(1), row - edit.row, col - edit.col)}`
								: edit.text;
							if (setCellInput(tx, wb, edit.sheet, row, col, text).error) {
								rejected = true;
							}
						}
					}
				}
			} else if (setCellInput(tx, wb, edit.sheet, edit.row, edit.col, edit.text).error) {
				rejected = true;
			} else if (edit.text !== '') {
				// Typing next to a table extends it, as with Excel's AutoCorrect option.
				autoExpandTable(tx, wb, edit.sheet, edit.row, edit.col);
			}
		});
		if (rejected) {
			store.setUi({ message: store.t('message.arrayPart') });
			return false;
		}
	}
	store.setUi({ edit: null });
	const move = opts.move ?? 'down';
	if (move !== 'none' && !opts.fillSelection) {
		advanceActive(store, move);
	}
	return true;
}

/** Delete key: clear contents of the selection. */
export function clearSelectionContents(store: ViewerStore): void {
	const sheetIndex = store.state.activeSheet;
	const sel = store.selectionOf();
	store.transact(store.t('cmd.clear'), (tx) => {
		clearRanges(tx, store.workbook, sheetIndex, sel.ranges, 'contents');
	});
}
