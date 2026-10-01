import type { RangeAddress } from 'xlsx-viewer-core';
import { cellKey, fillRange, MAX_ROW_INDEX } from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';

function hasData(store: ViewerStore, row: number, col: number): boolean {
	const c = store.activeSheet.cells.get(cellKey(row, col));
	return c !== undefined && (c.value !== null || c.formula !== undefined);
}

/**
 * Fill from the fill handle. `target` is the dragged range; `'auto'`
 * (double-click on the handle) fills down as far as the data in the
 * adjacent column goes.
 */
export function fillFromHandle(
	store: ViewerStore,
	source: RangeAddress,
	target: RangeAddress | 'auto',
): void {
	let dest: RangeAddress;
	if (target === 'auto') {
		const neighbour =
			source.left > 0 && hasData(store, source.bottom + 1, source.left - 1)
				? source.left - 1
				: source.right + 1;
		let bottom = source.bottom;
		while (bottom < MAX_ROW_INDEX && hasData(store, bottom + 1, neighbour)) {
			bottom++;
		}
		if (bottom === source.bottom) {
			return;
		}
		dest = { ...source, bottom };
	} else {
		dest = target;
	}
	if (
		dest.top === source.top &&
		dest.bottom === source.bottom &&
		dest.left === source.left &&
		dest.right === source.right
	) {
		return;
	}
	const sheetIndex = store.state.activeSheet;
	const sel = store.selectionOf();
	store.transact(
		store.t('undo.autoFill'),
		(tx) => {
			fillRange(tx, store.workbook, sheetIndex, source, dest);
		},
		{ selectionAfter: { active: sel.active, anchor: sel.anchor, ranges: [dest] } },
	);
}

/** Ctrl+D / Ctrl+R: copy the first row (column) of the selection into the rest. */
export function fillDownOrRight(store: ViewerStore, direction: 'down' | 'right'): void {
	const sheetIndex = store.state.activeSheet;
	const sel = store.selectionOf();
	store.transact(store.t(direction === 'down' ? 'undo.fillDown' : 'undo.fillRight'), (tx) => {
		for (const r of sel.ranges) {
			if (direction === 'down') {
				const src =
					r.top === r.bottom
						? { ...r, top: Math.max(0, r.top - 1), bottom: Math.max(0, r.top - 1) }
						: { ...r, bottom: r.top };
				fillRange(tx, store.workbook, sheetIndex, src, { ...r, top: src.top }, { copyOnly: true });
			} else {
				const src =
					r.left === r.right
						? { ...r, left: Math.max(0, r.left - 1), right: Math.max(0, r.left - 1) }
						: { ...r, right: r.left };
				fillRange(
					tx,
					store.workbook,
					sheetIndex,
					src,
					{ ...r, left: src.left },
					{ copyOnly: true },
				);
			}
		}
	});
}
