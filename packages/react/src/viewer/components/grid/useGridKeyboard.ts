import { useCallback, useRef, type KeyboardEvent } from 'react';

import { mapGridKey, type GridCommand } from '../../../engine/keymap';
import {
	copyToSystem,
	clearClipboardMark,
	pasteFromSystem,
} from '../../commands/clipboard-commands';
import { beginEdit, clearSelectionContents } from '../../commands/edit-commands';
import { deleteSelectedObject } from '../../commands/object-commands';
import { fillDownOrRight } from '../../commands/fill-commands';
import {
	changeDecimals,
	setNumberFormat,
	toggleFontFlag,
	toggleUnderline,
} from '../../commands/format-commands';
import { hideColumns, hideRows } from '../../commands/layout-commands';
import { autoSum, insertNow, switchSheet, toggleShowFormulas } from '../../commands/misc-commands';
import {
	advanceActive,
	goEnd,
	goHome,
	moveActive,
	pageMove,
	selectAllOrRegion,
	selectEntireColumn,
	selectEntireRow,
} from '../../commands/selection-commands';
import type { ViewerStore } from '../../state/viewer-store';
import { groupSelection, ungroupSelection } from '../../commands/outline-commands';

export interface KeyboardHost {
	store: ViewerStore;
	/** `end`: the moving end of a selection being extended, else the active cell. */
	ensureActiveVisible: (target?: 'active' | 'end') => void;
	/** Rows visible in the body, for Page Up / Down. */
	pageRows: () => number;
	pageCols: () => number;
	/** Host-level actions (save, open, print, dialogs) handled by the viewer shell. */
	onAction?: (action: GridCommand['id']) => void;
	openContextMenu?: () => void;
}

/** Commands that the viewer shell handles (dialogs, file actions). */
const SHELL_ACTIONS = new Set<GridCommand['id']>([
	'formatCells',
	'find',
	'replace',
	'goTo',
	'save',
	'open',
	'print',
	'hyperlink',
	'note',
	'filter',
	'deleteCells',
	'insertCells',
	'newSheet',
	'table',
	'dropdown',
	'repeat',
	'unhideRows',
	'unhideCols',
]);

/** Keyboard handling of the grid while it has focus and no cell is being edited. */
export function useGridKeyboard(host: KeyboardHost) {
	const hostRef = useRef(host);
	hostRef.current = host;

	return useCallback((e: KeyboardEvent<HTMLElement>) => {
		const { store, ensureActiveVisible } = hostRef.current;
		if (store.state.edit || e.nativeEvent.isComposing) {
			return;
		}
		const cmd = mapGridKey(e);
		if (!cmd) {
			return;
		}
		// Let the browser fire copy / cut / paste events (handled on the focus sink).
		if (cmd.id === 'copy' || cmd.id === 'cut' || cmd.id === 'paste') {
			return;
		}
		if (cmd.id === 'type') {
			if (!store.readOnly) {
				// The character itself arrives through the input event of the focus sink.
				return;
			}
			return;
		}
		e.preventDefault();
		let moved = true;
		switch (cmd.id) {
			case 'move':
				moveActive(store, cmd.dir, { extend: cmd.extend, jump: cmd.jump });
				break;
			case 'advance':
				advanceActive(store, cmd.dir);
				break;
			case 'page': {
				const vertical = cmd.dir === 'up' || cmd.dir === 'down';
				pageMove(
					store,
					cmd.dir,
					vertical ? hostRef.current.pageRows() : hostRef.current.pageCols(),
					cmd.extend,
				);
				break;
			}
			case 'home':
				goHome(store, cmd.ctrl, cmd.extend);
				break;
			case 'end':
				goEnd(store, cmd.extend);
				break;
			case 'sheet':
				switchSheet(store, cmd.delta);
				moved = false;
				break;
			case 'edit':
				beginEdit(store, { mode: 'edit' });
				break;
			case 'editClear':
				if (!store.readOnly) {
					beginEdit(store, { initialText: '', mode: 'enter' });
				}
				break;
			case 'clear':
				if (store.readOnly) {
					break;
				}
				if (store.state.selectedObject) {
					deleteSelectedObject(store);
					moved = false;
				} else {
					clearSelectionContents(store);
				}
				break;
			case 'escape':
				clearClipboardMark(store);
				store.setUi({ formatPainter: null, selectedObject: null });
				moved = false;
				break;
			case 'selectAll':
				selectAllOrRegion(store);
				moved = false;
				break;
			case 'selectColumn':
				selectEntireColumn(store);
				moved = false;
				break;
			case 'selectRow':
				selectEntireRow(store);
				moved = false;
				break;
			case 'undo':
				store.undo();
				break;
			case 'redo':
				store.redo();
				break;
			case 'bold':
			case 'italic':
				toggleFontFlag(store, cmd.id);
				moved = false;
				break;
			case 'strike':
				toggleFontFlag(store, 'strike');
				moved = false;
				break;
			case 'underline':
				toggleUnderline(store);
				moved = false;
				break;
			case 'fillDown':
				fillDownOrRight(store, 'down');
				break;
			case 'fillRight':
				fillDownOrRight(store, 'right');
				break;
			case 'insertDate':
				insertNow(store, 'date');
				break;
			case 'insertTime':
				insertNow(store, 'time');
				break;
			case 'numberFormat':
				setNumberFormat(store, cmd.code);
				moved = false;
				break;
			case 'showFormulas':
				toggleShowFormulas(store);
				moved = false;
				break;
			case 'calculate':
				store.recalculate();
				moved = false;
				break;
			case 'autoSum':
				autoSum(store);
				break;
			case 'hideRows':
				hideRows(store, true);
				break;
			case 'hideCols':
				hideColumns(store, true);
				break;
			case 'contextMenu':
				hostRef.current.openContextMenu?.();
				moved = false;
				break;
			case 'group':
				groupSelection(store);
				moved = false;
				break;
			case 'ungroup':
				ungroupSelection(store);
				moved = false;
				break;
			default:
				if (SHELL_ACTIONS.has(cmd.id)) {
					hostRef.current.onAction?.(cmd.id);
				}
				moved = false;
		}
		if (moved) {
			ensureActiveVisible('extend' in cmd && cmd.extend ? 'end' : 'active');
		}
		void changeDecimals;
		void copyToSystem;
		void pasteFromSystem;
	}, []);
}
