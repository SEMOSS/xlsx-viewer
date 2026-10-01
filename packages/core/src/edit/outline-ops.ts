import { computeOutline, type OutlineGroup } from '../sheet/outline';
import { updateCols, updateRow } from '../sheet/worksheet';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import type { EditTransaction } from './transaction';

export type OutlineAxis = 'rows' | 'cols';

const MAX_LEVEL = 7;

function begin(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	axis: OutlineAxis,
): Worksheet {
	tx.touchSheet(sheetIndex, false, axis === 'rows' ? 'formatRows' : 'formatColumns');
	const sheet = workbook.sheets[sheetIndex];
	sheet.dirty = true;
	return sheet;
}

function setLines(
	sheet: Worksheet,
	axis: OutlineAxis,
	start: number,
	end: number,
	update: (info: { outlineLevel?: number; hidden?: boolean; collapsed?: boolean }) => {
		outlineLevel?: number;
		hidden?: boolean;
		collapsed?: boolean;
	},
): void {
	if (axis === 'rows') {
		for (let r = start; r <= end; r++) {
			updateRow(sheet, r, (info) => ({ ...info, ...update(info) }));
		}
	} else {
		updateCols(sheet, start, end, (info) => ({ ...info, ...update(info) }));
	}
}

function lineLevel(sheet: Worksheet, axis: OutlineAxis, index: number): number {
	if (axis === 'rows') {
		return sheet.rows.get(index)?.outlineLevel ?? 0;
	}
	return sheet.cols.find((c) => index >= c.min && index <= c.max)?.outlineLevel ?? 0;
}

function refreshMaxLevel(sheet: Worksheet, axis: OutlineAxis): void {
	const levels = computeOutline(sheet, axis).levels;
	if (axis === 'rows') {
		sheet.outlineLevelRow = levels || undefined;
	} else {
		sheet.outlineLevelCol = levels || undefined;
	}
}

const clean = (v: number) => (v > 0 ? v : undefined);

/** Data > Group: put the rows (columns) one outline level deeper (Excel allows 7 levels). */
export function groupLines(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	axis: OutlineAxis,
	start: number,
	end: number,
): void {
	const sheet = begin(tx, workbook, sheetIndex, axis);
	setLines(sheet, axis, start, end, (info) => ({
		outlineLevel: clean(Math.min(MAX_LEVEL, (info.outlineLevel ?? 0) + 1)),
	}));
	refreshMaxLevel(sheet, axis);
}

/** Data > Ungroup: one level up; lines already at the top level are left alone. */
export function ungroupLines(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	axis: OutlineAxis,
	start: number,
	end: number,
): void {
	const sheet = begin(tx, workbook, sheetIndex, axis);
	setLines(sheet, axis, start, end, (info) => ({
		outlineLevel: clean((info.outlineLevel ?? 0) - 1),
	}));
	refreshMaxLevel(sheet, axis);
}

/**
 * The +/- outline button (Hide Detail / Show Detail): collapsing hides the
 * whole group; expanding shows its lines again, except those inside nested
 * groups that are themselves collapsed.
 */
export function setGroupCollapsed(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	axis: OutlineAxis,
	group: Pick<OutlineGroup, 'level' | 'start' | 'end' | 'summary'>,
	collapsed: boolean,
): void {
	const sheet = begin(tx, workbook, sheetIndex, axis);
	if (collapsed) {
		setLines(sheet, axis, group.start, group.end, () => ({ hidden: true }));
	} else {
		const inner = computeOutline(sheet, axis).groups.filter(
			(g) => g.level > group.level && g.start >= group.start && g.end <= group.end && g.collapsed,
		);
		for (let i = group.start; i <= group.end; i++) {
			const hidden = inner.some((g) => i >= g.start && i <= g.end);
			setLines(sheet, axis, i, i, () => ({ hidden: hidden || undefined }));
		}
	}
	if (group.summary >= 0) {
		setLines(sheet, axis, group.summary, group.summary, () => ({
			collapsed: collapsed || undefined,
		}));
	}
}

/** The outline level buttons (1, 2, 3...): show levels below `level`, collapse the rest. */
export function showOutlineLevel(
	tx: EditTransaction,
	workbook: Workbook,
	sheetIndex: number,
	axis: OutlineAxis,
	level: number,
): void {
	const sheet = begin(tx, workbook, sheetIndex, axis);
	const { groups } = computeOutline(sheet, axis);
	const lines = new Set<number>();
	for (const g of groups) {
		for (let i = g.start; i <= g.end; i++) {
			lines.add(i);
		}
	}
	for (const i of lines) {
		const hidden = lineLevel(sheet, axis, i) >= level;
		setLines(sheet, axis, i, i, () => ({ hidden: hidden || undefined }));
	}
	for (const g of groups) {
		if (g.summary >= 0) {
			setLines(sheet, axis, g.summary, g.summary, () => ({
				collapsed: g.level >= level || undefined,
			}));
		}
	}
}
