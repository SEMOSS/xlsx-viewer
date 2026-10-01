import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import {
	adjustFormulaForStructureChange,
	adjustRangeForStructureChange,
	type StructureChange,
} from '../formula/language';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { ChartModel } from '../types/chart';
import type { DrawingAnchor, DrawingObject } from '../types/drawing';
import type { ColInfo, RowInfo, Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { shiftPageBreaks } from './page-setup-ops-breaks';
import type { EditTransaction } from './transaction';

type Axis = 'rows' | 'cols';

function changeFor(
	sheet: Worksheet,
	axis: Axis,
	insert: boolean,
	at: number,
	count: number,
): StructureChange {
	return {
		kind: insert
			? axis === 'rows'
				? 'insertRows'
				: 'insertCols'
			: axis === 'rows'
				? 'deleteRows'
				: 'deleteCols',
		sheet: sheet.name,
		at,
		count,
	};
}

/** New index of a line after the change, or null when it was deleted. */
function shiftIndex(
	index: number,
	insert: boolean,
	at: number,
	count: number,
	max: number,
): number | null {
	if (insert) {
		if (index < at) {
			return index;
		}
		const n = index + count;
		return n > max ? null : n;
	}
	if (index < at) {
		return index;
	}
	if (index < at + count) {
		return null;
	}
	return index - count;
}

function adjustRange(range: RangeAddress, change: StructureChange): RangeAddress | null {
	return adjustRangeForStructureChange(range, change);
}

function adjustRanges(ranges: RangeAddress[], change: StructureChange): RangeAddress[] {
	const out: RangeAddress[] = [];
	for (const r of ranges) {
		const a = adjustRange(r, change);
		if (a) {
			out.push(a);
		}
	}
	return out;
}

function adjustFormula(formula: string, change: StructureChange, formulaSheet: string): string {
	return adjustFormulaForStructureChange(formula, change, formulaSheet).formula;
}

function shiftCells(
	sheet: Worksheet,
	axis: Axis,
	insert: boolean,
	at: number,
	count: number,
): void {
	const max = axis === 'rows' ? MAX_ROW_INDEX : MAX_COL_INDEX;
	const next = new Map<number, Cell>();
	const remap = new Map<number, number>();
	for (const [key, cell] of sheet.cells) {
		const row = keyRow(key);
		const col = keyCol(key);
		const moved = shiftIndex(axis === 'rows' ? row : col, insert, at, count, max);
		if (moved === null) {
			continue;
		}
		const newKey = axis === 'rows' ? cellKey(moved, col) : cellKey(row, moved);
		remap.set(key, newKey);
		next.set(newKey, cell);
	}
	// Fix array anchors that moved.
	for (const [key, cell] of next) {
		if (cell.arrayAnchor !== undefined) {
			const anchor = remap.get(cell.arrayAnchor);
			if (anchor === undefined) {
				next.delete(key);
			} else if (anchor !== cell.arrayAnchor) {
				next.set(key, { ...cell, arrayAnchor: anchor });
			}
		}
	}
	sheet.cells = next;
}

function shiftRows(sheet: Worksheet, insert: boolean, at: number, count: number): void {
	const next = new Map<number, RowInfo>();
	for (const [row, info] of sheet.rows) {
		const moved = shiftIndex(row, insert, at, count, MAX_ROW_INDEX);
		if (moved !== null) {
			next.set(moved, info);
		}
	}
	if (insert && at > 0) {
		// Inserted rows take the formatting (height and style) of the row above, like Excel.
		const above = sheet.rows.get(at - 1);
		if (above && (above.customFormat || above.height !== undefined)) {
			for (let r = at; r < at + count; r++) {
				next.set(r, {
					height: above.height,
					customHeight: above.customHeight,
					style: above.style,
					customFormat: above.customFormat,
				});
			}
		}
	}
	sheet.rows = next;
}

function shiftCols(sheet: Worksheet, insert: boolean, at: number, count: number): void {
	const out: ColInfo[] = [];
	for (const c of sheet.cols) {
		if (insert) {
			if (c.max < at) {
				out.push(c);
			} else if (c.min >= at) {
				if (c.min + count <= MAX_COL_INDEX) {
					out.push({ ...c, min: c.min + count, max: Math.min(MAX_COL_INDEX, c.max + count) });
				}
			} else {
				// The record spans the insertion point: the inserted columns inherit it.
				out.push({ ...c, max: Math.min(MAX_COL_INDEX, c.max + count) });
			}
		} else {
			const end = at + count - 1;
			if (c.max < at) {
				out.push(c);
			} else if (c.min > end) {
				out.push({ ...c, min: c.min - count, max: c.max - count });
			} else {
				const keptBefore = Math.max(0, at - c.min);
				const keptAfter = Math.max(0, c.max - end);
				const kept = keptBefore + keptAfter;
				if (kept > 0) {
					const min = Math.min(c.min, at);
					out.push({ ...c, min, max: min + kept - 1 });
				}
			}
		}
	}
	sheet.cols = out.sort((a, b) => a.min - b.min);
}

function shiftAnchor(
	anchor: DrawingAnchor,
	axis: Axis,
	insert: boolean,
	at: number,
	count: number,
): DrawingAnchor {
	const move = (p: { col: number; colOff: number; row: number; rowOff: number }) => {
		const index = axis === 'rows' ? p.row : p.col;
		if (index < at) {
			return p;
		}
		let next: number;
		if (insert) {
			next = index + count;
		} else {
			next = index < at + count ? at : index - count;
		}
		return axis === 'rows'
			? { ...p, row: next, rowOff: !insert && index < at + count ? 0 : p.rowOff }
			: { ...p, col: next, colOff: !insert && index < at + count ? 0 : p.colOff };
	};
	if (anchor.type === 'twoCell') {
		if (anchor.editAs === 'absolute') {
			return anchor;
		}
		if (anchor.editAs === 'oneCell') {
			const from = move(anchor.from);
			const dRow = from.row - anchor.from.row;
			const dCol = from.col - anchor.from.col;
			return {
				...anchor,
				from,
				to: { ...anchor.to, row: anchor.to.row + dRow, col: anchor.to.col + dCol },
			};
		}
		return { ...anchor, from: move(anchor.from), to: move(anchor.to) };
	}
	if (anchor.type === 'oneCell') {
		return { ...anchor, from: move(anchor.from) };
	}
	return anchor;
}

function adjustChart(chart: ChartModel, change: StructureChange, sheetName: string): ChartModel {
	let changed = false;
	const fix = (ref?: string) => {
		if (!ref) {
			return ref;
		}
		const next = adjustFormula(ref, change, sheetName);
		if (next !== ref) {
			changed = true;
		}
		return next;
	};
	const plots = chart.plots.map((p) => ({
		...p,
		series: p.series.map((s) => ({
			...s,
			name: s.name ? { ...s.name, ref: fix(s.name.ref) } : s.name,
			categories: s.categories ? { ...s.categories, ref: fix(s.categories.ref) } : s.categories,
			values: s.values ? { ...s.values, ref: fix(s.values.ref) } : s.values,
			bubbleSizes: s.bubbleSizes
				? { ...s.bubbleSizes, ref: fix(s.bubbleSizes.ref) }
				: s.bubbleSizes,
		})),
	}));
	return changed ? { ...chart, plots, dirty: true } : chart;
}

function adjustDrawing(
	obj: DrawingObject,
	axis: Axis,
	insert: boolean,
	at: number,
	count: number,
	apply: boolean,
): DrawingObject {
	const anchor = apply ? shiftAnchor(obj.anchor, axis, insert, at, count) : obj.anchor;
	return anchor === obj.anchor ? obj : { ...obj, anchor, dirty: true };
}

/**
 * Insert or delete whole rows / columns on one sheet, updating every
 * dependent structure in the workbook. `at` is 0-based.
 */
export function changeStructure(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	axis: Axis,
	insert: boolean,
	at: number,
	count: number,
): void {
	if (count <= 0) {
		return;
	}
	const kind = insert
		? axis === 'rows'
			? 'insertRows'
			: 'insertColumns'
		: axis === 'rows'
			? 'deleteRows'
			: 'deleteColumns';
	tx.touchSheet(sheetIndex, true, kind);
	for (let i = 0; i < workbook.sheets.length; i++) {
		if (i !== sheetIndex) {
			tx.touchSheet(i, true, 'references');
		}
	}
	tx.touchWorkbook();
	const sheet = workbook.sheets[sheetIndex];
	const change = changeFor(sheet, axis, insert, at, count);

	shiftCells(sheet, axis, insert, at, count);
	if (axis === 'rows') {
		shiftRows(sheet, insert, at, count);
	} else {
		shiftCols(sheet, insert, at, count);
	}
	if (insert && axis === 'rows' && at > 0) {
		// Copy cell formats of the row above into the inserted rows.
		for (const [key, cell] of [...sheet.cells]) {
			if (keyRow(key) === at - 1 && cell.style) {
				for (let r = at; r < at + count; r++) {
					sheet.cells.set(cellKey(r, keyCol(key)), { value: null, style: cell.style });
				}
			}
		}
	} else if (insert && axis === 'cols' && at > 0) {
		for (const [key, cell] of [...sheet.cells]) {
			if (keyCol(key) === at - 1 && cell.style) {
				for (let c = at; c < at + count; c++) {
					sheet.cells.set(cellKey(keyRow(key), c), { value: null, style: cell.style });
				}
			}
		}
	}

	// Formulas everywhere (any sheet may reference this one).
	for (const s of workbook.sheets) {
		let touched = false;
		for (const [key, cell] of s.cells) {
			let next = cell;
			if (cell.formula) {
				const f = adjustFormula(cell.formula, change, s.name);
				if (f !== cell.formula) {
					next = { ...next, formula: f };
				}
			}
			if (s === sheet && cell.array) {
				const ref = adjustRange(cell.array.ref, change);
				if (ref && ref !== cell.array.ref) {
					next = { ...next, array: { ...cell.array, ref } };
				}
			}
			if (next !== cell) {
				s.cells.set(key, next);
				touched = true;
			}
		}
		if (s !== sheet) {
			s.drawings = s.drawings.map((d) =>
				d.kind === 'chart' ? { ...d, chart: adjustChart(d.chart, change, s.name) } : d,
			);
			if (s.drawings.some((d) => d.kind === 'chart' && d.chart.dirty)) {
				s.drawingsDirty = true;
			}
		}
		if (touched) {
			s.dirty = true;
		}
	}
	for (const name of workbook.definedNames) {
		const scope =
			name.localSheet !== undefined
				? (workbook.sheets[name.localSheet]?.name ?? sheet.name)
				: sheet.name;
		name.formula = adjustFormula(name.formula, change, scope);
	}

	// Sheet features.
	sheet.merges = adjustRanges(sheet.merges, change).filter(
		(m) => m.top !== m.bottom || m.left !== m.right,
	);
	const breaks = axis === 'rows' ? 'rowBreaks' : 'colBreaks';
	if (sheet[breaks]) {
		sheet[breaks] = shiftPageBreaks(sheet[breaks], insert, at, count);
	}
	sheet.hyperlinks = sheet.hyperlinks.flatMap((h) => {
		const ref = adjustRange(h.ref, change);
		return ref ? [{ ...h, ref }] : [];
	});
	sheet.conditionalFormats = sheet.conditionalFormats.flatMap((cf) => {
		const ranges = adjustRanges(cf.ranges, change);
		if (ranges.length === 0) {
			return [];
		}
		return [
			{
				...cf,
				ranges,
				rules: cf.rules.map((rule) => ({
					...rule,
					formulas: rule.formulas?.map((f) => adjustFormula(f, change, sheet.name)),
				})),
			},
		];
	});
	sheet.dataValidations = sheet.dataValidations.flatMap((dv) => {
		const ranges = adjustRanges(dv.ranges, change);
		if (ranges.length === 0) {
			return [];
		}
		return [
			{
				...dv,
				ranges,
				formula1:
					dv.formula1 !== undefined ? adjustFormula(dv.formula1, change, sheet.name) : undefined,
				formula2:
					dv.formula2 !== undefined ? adjustFormula(dv.formula2, change, sheet.name) : undefined,
			},
		];
	});
	if (sheet.autoFilter) {
		const ref = adjustRange(sheet.autoFilter.ref, change);
		sheet.autoFilter = ref ? { ...sheet.autoFilter, ref } : undefined;
	}
	sheet.tables = sheet.tables.flatMap((t) => {
		const ref = adjustRange(t.ref, change);
		return ref
			? [
					{
						...t,
						ref,
						dirty: true,
						autoFilter: t.autoFilter
							? {
									...t.autoFilter,
									ref: { ...ref, bottom: ref.bottom - (t.totalsRowCount ? 1 : 0) },
								}
							: t.autoFilter,
					},
				]
			: [];
	});
	const lineOf = (p: { row: number; col: number }) => (axis === 'rows' ? p.row : p.col);
	const movePoint = <T extends { row: number; col: number }>(p: T): T | null => {
		const moved = shiftIndex(
			lineOf(p),
			insert,
			at,
			count,
			axis === 'rows' ? MAX_ROW_INDEX : MAX_COL_INDEX,
		);
		if (moved === null) {
			return null;
		}
		return axis === 'rows' ? { ...p, row: moved } : { ...p, col: moved };
	};
	const notesBefore = sheet.notes.length + sheet.threadedComments.length;
	sheet.notes = sheet.notes.flatMap((n) => {
		const ref = movePoint(n.ref);
		return ref ? [{ ...n, ref }] : [];
	});
	sheet.threadedComments = sheet.threadedComments.flatMap((c) => {
		const ref = movePoint(c.ref);
		return ref ? [{ ...c, ref }] : [];
	});
	if (sheet.notes.length + sheet.threadedComments.length > 0 || notesBefore > 0) {
		sheet.commentsDirty = true;
	}
	if (sheet.drawings.length) {
		sheet.drawings = sheet.drawings.map((d) => {
			const moved = adjustDrawing(d, axis, insert, at, count, true);
			return moved.kind === 'chart'
				? { ...moved, chart: adjustChart(moved.chart, change, sheet.name) }
				: moved;
		});
		sheet.drawingsDirty = true;
	}
	sheet.sparklineGroups = sheet.sparklineGroups.map((g) => ({
		...g,
		sparklines: g.sparklines.flatMap((s) => {
			const location = movePoint(s.location);
			return location
				? [{ ...s, location, formula: adjustFormula(s.formula, change, sheet.name) }]
				: [];
		}),
	}));
	if (sheet.view.selection) {
		const active = movePoint(sheet.view.selection.activeCell) ?? sheet.view.selection.activeCell;
		sheet.view = {
			...sheet.view,
			selection: {
				...sheet.view.selection,
				activeCell: active,
				ranges: adjustRanges(sheet.view.selection.ranges, change),
			},
		};
	}
	sheet.dirty = true;
}

export function insertRows(
	tx: EditTransaction,
	wb: Workbook,
	sheetIndex: number,
	at: number,
	count = 1,
): void {
	changeStructure(tx, wb, sheetIndex, 'rows', true, at, count);
}

export function deleteRows(
	tx: EditTransaction,
	wb: Workbook,
	sheetIndex: number,
	at: number,
	count = 1,
): void {
	changeStructure(tx, wb, sheetIndex, 'rows', false, at, count);
}

export function insertColumns(
	tx: EditTransaction,
	wb: Workbook,
	sheetIndex: number,
	at: number,
	count = 1,
): void {
	changeStructure(tx, wb, sheetIndex, 'cols', true, at, count);
}

export function deleteColumns(
	tx: EditTransaction,
	wb: Workbook,
	sheetIndex: number,
	at: number,
	count = 1,
): void {
	changeStructure(tx, wb, sheetIndex, 'cols', false, at, count);
}
