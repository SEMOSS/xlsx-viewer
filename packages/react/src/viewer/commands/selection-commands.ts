import type { CellAddress, RangeAddress, Worksheet } from 'xlsx-viewer-core';
import { expandToMerges, findSheetIndex, getColInfo, parseRange } from 'xlsx-viewer-core';

import {
	currentRegion,
	DataProbe,
	jumpToEdge,
	moveOne,
	type Direction,
} from '../../engine/selection/navigation';
import {
	currentRange,
	cycleActiveWithinSelection,
	rangeFromCorners,
	selectAll,
	singleCellSelection,
	wholeColumns,
	wholeRows,
	type SelectionState,
} from '../../engine/selection/selection';
import type { ViewerStore } from '../state/viewer-store';

export function probeFor(sheet: Worksheet): DataProbe {
	return new DataProbe(
		sheet,
		(r) => sheet.rows.get(r)?.hidden === true,
		(c) => getColInfo(sheet, c)?.hidden === true,
	);
}

/** Merged area containing a cell, or the cell itself. */
function mergeAt(sheet: Worksheet, row: number, col: number): RangeAddress {
	for (const m of sheet.merges) {
		if (row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) {
			return m;
		}
	}
	return { top: row, left: col, bottom: row, right: col };
}

/** Select one cell (or extend / add, like Shift+click / Ctrl+click). */
export function selectCell(
	store: ViewerStore,
	row: number,
	col: number,
	opts: { extend?: boolean; add?: boolean } = {},
): void {
	const sheet = store.activeSheet;
	const sel = store.selectionOf();
	if (opts.extend) {
		const range = expandToMerges(sheet.merges, rangeFromCorners(sel.anchor, { row, col }));
		const ranges = [...sel.ranges.slice(0, -1), range];
		store.setSelection({ ...sel, ranges });
		return;
	}
	const m = mergeAt(sheet, row, col);
	const active = { row: m.top, col: m.left };
	if (opts.add) {
		store.setSelection({ active, anchor: active, ranges: [...sel.ranges, m] });
		return;
	}
	store.setSelection({ active, anchor: active, ranges: [m] });
}

/** Extend the current range so that its moving corner is at (row, col) (mouse drag). */
export function dragSelectionTo(store: ViewerStore, row: number, col: number): void {
	const sheet = store.activeSheet;
	const sel = store.selectionOf();
	const range = expandToMerges(sheet.merges, rangeFromCorners(sel.anchor, { row, col }));
	const prev = currentRange(sel);
	if (
		prev.top === range.top &&
		prev.left === range.left &&
		prev.bottom === range.bottom &&
		prev.right === range.right
	) {
		return;
	}
	store.setSelection({ ...sel, ranges: [...sel.ranges.slice(0, -1), range] });
}

/** The moving corner of the current range (opposite the anchor). */
function movingCorner(sel: SelectionState): CellAddress {
	const r = currentRange(sel);
	return {
		row: sel.anchor.row === r.top ? r.bottom : r.top,
		col: sel.anchor.col === r.left ? r.right : r.left,
	};
}

/** Arrow-key navigation: move (or extend with Shift), jumping with Ctrl. */
export function moveActive(
	store: ViewerStore,
	dir: Direction,
	opts: { extend?: boolean; jump?: boolean } = {},
): void {
	const sheet = store.activeSheet;
	const probe = probeFor(sheet);
	const sel = store.selectionOf();
	if (opts.extend) {
		const corner = movingCorner(sel);
		const target = opts.jump ? jumpToEdge(probe, corner, dir) : moveOne(probe, corner, dir);
		const range = expandToMerges(sheet.merges, rangeFromCorners(sel.anchor, target));
		store.setSelection({ ...sel, ranges: [range] });
		return;
	}
	// Leave a merged area from its edge in the direction of travel.
	const m = mergeAt(sheet, sel.active.row, sel.active.col);
	const from: CellAddress = {
		row: dir === 'down' ? m.bottom : dir === 'up' ? m.top : sel.active.row,
		col: dir === 'right' ? m.right : dir === 'left' ? m.left : sel.active.col,
	};
	const target = opts.jump ? jumpToEdge(probe, from, dir) : moveOne(probe, from, dir);
	selectCell(store, target.row, target.col);
}

/** Enter / Tab: move within a multi-cell selection, or move the single cell. */
export function advanceActive(store: ViewerStore, dir: Direction): void {
	const sel = store.selectionOf();
	const r = currentRange(sel);
	const multi = sel.ranges.length > 1 || r.top !== r.bottom || r.left !== r.right;
	const merge = mergeAt(store.activeSheet, sel.active.row, sel.active.col);
	const onlyMerge =
		sel.ranges.length === 1 &&
		merge.top === r.top &&
		merge.left === r.left &&
		merge.bottom === r.bottom &&
		merge.right === r.right;
	if (multi && !onlyMerge) {
		store.setSelection(cycleActiveWithinSelection(sel, dir));
		return;
	}
	moveActive(store, dir);
}

/** Page Up / Page Down (rows) or Alt+Page Up / Down (columns). */
export function pageMove(
	store: ViewerStore,
	dir: Direction,
	pageSize: number,
	extend = false,
): void {
	const sheet = store.activeSheet;
	const probe = probeFor(sheet);
	const sel = store.selectionOf();
	let target: CellAddress = extend ? movingCorner(sel) : { ...sel.active };
	for (let i = 0; i < pageSize; i++) {
		target = moveOne(probe, target, dir);
	}
	if (extend) {
		store.setSelection({ ...sel, ranges: [rangeFromCorners(sel.anchor, target)] });
	} else {
		selectCell(store, target.row, target.col);
	}
}

export function goHome(store: ViewerStore, ctrl: boolean, extend: boolean): void {
	const sel = store.selectionOf();
	const pane = store.activeSheet.view.pane;
	const frozenRow = pane?.state === 'frozen' ? Math.round(pane.ySplit) : 0;
	const frozenCol = pane?.state === 'frozen' ? Math.round(pane.xSplit) : 0;
	const target = ctrl
		? { row: frozenRow, col: frozenCol }
		: { row: sel.active.row, col: frozenCol };
	if (extend) {
		store.setSelection({ ...sel, ranges: [rangeFromCorners(sel.anchor, target)] });
	} else {
		selectCell(store, target.row, target.col);
	}
}

/** Ctrl+End: the last used cell. */
export function goEnd(store: ViewerStore, extend: boolean): void {
	const target = probeFor(store.activeSheet).lastUsedCell();
	const sel = store.selectionOf();
	if (extend) {
		store.setSelection({ ...sel, ranges: [rangeFromCorners(sel.anchor, target)] });
	} else {
		selectCell(store, target.row, target.col);
	}
}

export function selectColumns(
	store: ViewerStore,
	from: number,
	to: number,
	opts: { add?: boolean } = {},
): void {
	const sel = store.selectionOf();
	const top = store.activeSheet.view.pane?.state === 'frozen' ? 0 : 0;
	const active = { row: top, col: Math.min(from, to) };
	const range = wholeColumns(from, to);
	store.setSelection({
		active: opts.add ? active : { row: sel.active.row >= 0 ? top : top, col: from },
		anchor: { row: 0, col: from },
		ranges: opts.add ? [...sel.ranges, range] : [range],
	});
}

export function selectRows(
	store: ViewerStore,
	from: number,
	to: number,
	opts: { add?: boolean } = {},
): void {
	const sel = store.selectionOf();
	const range = wholeRows(from, to);
	store.setSelection({
		active: { row: from, col: 0 },
		anchor: { row: from, col: 0 },
		ranges: opts.add ? [...sel.ranges, range] : [range],
	});
}

/** Ctrl+A: current region first, the whole sheet on the second press. */
export function selectAllOrRegion(store: ViewerStore): void {
	const sel = store.selectionOf();
	const region = currentRegion(probeFor(store.activeSheet), sel.active);
	const cur = currentRange(sel);
	const isRegion =
		cur.top === region.top &&
		cur.left === region.left &&
		cur.bottom === region.bottom &&
		cur.right === region.right;
	const single = region.top === region.bottom && region.left === region.right;
	if (isRegion || single) {
		store.setSelection({ active: sel.active, anchor: { row: 0, col: 0 }, ranges: [selectAll()] });
	} else {
		store.setSelection({
			active: sel.active,
			anchor: { row: region.top, col: region.left },
			ranges: [region],
		});
	}
}

export function selectEntireColumn(store: ViewerStore): void {
	const r = currentRange(store.selectionOf());
	const sel = store.selectionOf();
	store.setSelection({ ...sel, ranges: [wholeColumns(r.left, r.right)] });
}

export function selectEntireRow(store: ViewerStore): void {
	const r = currentRange(store.selectionOf());
	const sel = store.selectionOf();
	store.setSelection({ ...sel, ranges: [wholeRows(r.top, r.bottom)] });
}

/**
 * Go To a reference typed in the Name Box or the Go To dialog: `B7`,
 * `A1:C10`, `Sheet2!D4`, or a defined name. Returns false when invalid.
 */
export function goToReference(store: ViewerStore, text: string): boolean {
	const wb = store.workbook;
	let ref = text.trim();
	let sheetIndex = store.state.activeSheet;
	const name = wb.definedNames.find(
		(n) =>
			n.name.toLowerCase() === ref.toLowerCase() &&
			(n.localSheet === undefined || n.localSheet === sheetIndex),
	);
	if (name) {
		ref = name.formula;
	}
	const bang = ref.lastIndexOf('!');
	if (bang > 0) {
		let sheetName = ref.slice(0, bang);
		if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
			sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
		}
		const idx = findSheetIndex(wb, sheetName);
		if (idx < 0) {
			return false;
		}
		sheetIndex = idx;
		ref = ref.slice(bang + 1);
	}
	const range = parseRange(ref);
	if (!range) {
		return false;
	}
	if (sheetIndex !== store.state.activeSheet) {
		store.setUi({ activeSheet: sheetIndex });
	}
	const active = { row: range.top, col: range.left };
	store.setSelection({ active, anchor: active, ranges: [range] }, sheetIndex);
	return true;
}

export { singleCellSelection };
