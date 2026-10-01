import type { OutlineAxis, OutlineGroup } from 'xlsx-viewer-core';
import {
	computeOutline,
	groupLines,
	MAX_COL_INDEX,
	MAX_ROW_INDEX,
	setGroupCollapsed,
	showOutlineLevel,
	ungroupLines,
} from 'xlsx-viewer-core';

import { currentRange } from '../../engine/selection/selection';
import type { ViewerStore } from '../state/viewer-store';

/** The axis a Group / Ungroup applies to: columns for whole-column selections, else rows. */
export function outlineAxisOfSelection(store: ViewerStore): OutlineAxis {
	const r = currentRange(store.selectionOf());
	return r.top === 0 && r.bottom === MAX_ROW_INDEX && r.right < MAX_COL_INDEX ? 'cols' : 'rows';
}

function span(store: ViewerStore, axis: OutlineAxis): [number, number] {
	const r = currentRange(store.selectionOf());
	return axis === 'rows' ? [r.top, r.bottom] : [r.left, r.right];
}

/** Data > Group (Alt+Shift+Right). */
export function groupSelection(store: ViewerStore, axis = outlineAxisOfSelection(store)): void {
	const [a, b] = span(store, axis);
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('cmd.group'), (tx) =>
		groupLines(tx, store.workbook, sheetIndex, axis, a, b),
	);
}

/** Data > Ungroup (Alt+Shift+Left). */
export function ungroupSelection(store: ViewerStore, axis = outlineAxisOfSelection(store)): void {
	const [a, b] = span(store, axis);
	const sheetIndex = store.state.activeSheet;
	store.transact(store.t('cmd.ungroup'), (tx) =>
		ungroupLines(tx, store.workbook, sheetIndex, axis, a, b),
	);
}

/** The +/- outline buttons. */
export function setOutlineCollapsed(
	store: ViewerStore,
	axis: OutlineAxis,
	group: OutlineGroup,
	collapsed: boolean,
): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(
		store.t(collapsed ? 'cmd.hideDetail' : 'cmd.showDetail'),
		(tx) => setGroupCollapsed(tx, store.workbook, sheetIndex, axis, group, collapsed),
		{ autoRowHeight: false },
	);
}

/** The outline level buttons (1, 2, 3...). */
export function showOutlineLevelCmd(store: ViewerStore, axis: OutlineAxis, level: number): void {
	const sheetIndex = store.state.activeSheet;
	store.transact(
		store.t('undo.showLevel', { level }),
		(tx) => showOutlineLevel(tx, store.workbook, sheetIndex, axis, level),
		{ autoRowHeight: false },
	);
}

/**
 * Data > Show Detail / Hide Detail: the innermost group around the active
 * cell (on its detail lines or its summary line).
 */
export function toggleDetailAtSelection(store: ViewerStore, collapse: boolean): void {
	const { row, col } = store.selectionOf().active;
	for (const axis of ['rows', 'cols'] as const) {
		const at = axis === 'rows' ? row : col;
		const groups = computeOutline(store.activeSheet, axis).groups.filter(
			(g) => (at >= g.start && at <= g.end) || at === g.summary,
		);
		const target = groups.sort((a, b) => b.level - a.level)[0];
		if (target && target.collapsed !== collapse) {
			setOutlineCollapsed(store, axis, target, collapse);
			return;
		}
	}
}
