import { describe, expect, it } from 'vitest';

import { MAX_ROW_INDEX } from '../address/constants';
import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { getCell } from '../sheet/worksheet';
import type { BorderEdge, BorderSpec } from '../types/style';
import type { Workbook } from '../types/workbook';
import { applyBorderEdges, describeRangeBorders } from './border-ops';
import { applyBorderPreset } from './style-ops';
import { EditTransaction } from './transaction';

const r = (a1: string) => parseRange(a1)!;
const thin: BorderEdge = { style: 'thin' };
const thickRed: BorderEdge = { style: 'thick', color: { rgb: 'FFFF0000' } };

function borderAt(wb: Workbook, a1: string): BorderSpec {
	const { top, left } = r(a1);
	const style = getCell(wb.sheets[0], top, left)?.style ?? 0;
	return wb.styles.borders[wb.styles.cellXfs[style].borderId];
}

function apply(wb: Workbook, ref: string, change: Parameters<typeof applyBorderEdges>[4]) {
	const tx = new EditTransaction(wb, 'Borders');
	applyBorderEdges(tx, wb, 0, r(ref), change);
	return tx.commit();
}

describe('applyBorderEdges', () => {
	it('draws the outline on the sides of a range only', () => {
		const wb = createWorkbook();
		apply(wb, 'B2:C3', { top: thin, bottom: thin, left: thin, right: thin });
		expect(borderAt(wb, 'B2')).toEqual({ top: thin, left: thin });
		expect(borderAt(wb, 'C2')).toEqual({ top: thin, right: thin });
		expect(borderAt(wb, 'B3')).toEqual({ bottom: thin, left: thin });
		expect(borderAt(wb, 'C3')).toEqual({ bottom: thin, right: thin });
		expect(getCell(wb.sheets[0], 0, 1)).toBeUndefined();
	});

	it('draws inside lines on both cells of each inner gridline', () => {
		const wb = createWorkbook();
		apply(wb, 'A1:B2', { insideHorizontal: thickRed, insideVertical: thin });
		expect(borderAt(wb, 'A1')).toEqual({ bottom: thickRed, right: thin });
		expect(borderAt(wb, 'B1')).toEqual({ bottom: thickRed, left: thin });
		expect(borderAt(wb, 'A2')).toEqual({ top: thickRed, right: thin });
		expect(borderAt(wb, 'B2')).toEqual({ top: thickRed, left: thin });
	});

	it('removes edges with null and keeps edges that are not mentioned', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'all');
		applyBorderPreset(tx, wb, 0, r('A1:A1'), 'all', thin);
		tx.commit();
		apply(wb, 'A1', { top: null, right: thickRed });
		expect(borderAt(wb, 'A1')).toEqual({ bottom: thin, left: thin, right: thickRed });
	});

	it('sets and clears diagonals', () => {
		const wb = createWorkbook();
		apply(wb, 'A1', { diagonalDown: thin });
		expect(borderAt(wb, 'A1')).toEqual({ diagonal: thin, diagonalDown: true });
		apply(wb, 'A1', { diagonalUp: thin });
		expect(borderAt(wb, 'A1')).toMatchObject({
			diagonal: thin,
			diagonalDown: true,
			diagonalUp: true,
		});
		apply(wb, 'A1', { diagonalUp: null, diagonalDown: null });
		const b = borderAt(wb, 'A1');
		expect(b.diagonal).toBeUndefined();
		expect(b.diagonalUp).toBe(false);
		expect(b.diagonalDown).toBe(false);
	});

	it('keeps the shared gridline consistent with the neighbour outside the range', () => {
		const wb = createWorkbook();
		apply(wb, 'B1', { bottom: thin });
		expect(borderAt(wb, 'B1')).toEqual({ bottom: thin });
		apply(wb, 'B2', { top: null });
		expect(borderAt(wb, 'B1')).toEqual({});
		apply(wb, 'B1', { bottom: thin });
		apply(wb, 'B2', { top: thickRed });
		expect(borderAt(wb, 'B1')).toEqual({ bottom: thickRed });
		expect(borderAt(wb, 'B2')).toEqual({ top: thickRed });
	});

	it('formats whole columns through the column style', () => {
		const wb = createWorkbook();
		apply(wb, 'B:C', { left: thin, right: thin, insideVertical: thickRed, top: thin });
		const sheet = wb.sheets[0];
		const colStyle = (col: number) =>
			sheet.cols.find((c) => col >= c.min && col <= c.max)?.style ?? 0;
		const colBorder = (col: number) => wb.styles.borders[wb.styles.cellXfs[colStyle(col)].borderId];
		expect(colBorder(1)).toEqual({ left: thin, right: thickRed });
		expect(colBorder(2)).toEqual({ left: thickRed, right: thin });
		// Only the first row gets the top edge; no per-cell records further down.
		expect(borderAt(wb, 'B1').top).toEqual(thin);
		expect(getCell(sheet, MAX_ROW_INDEX, 1)).toBeUndefined();
		expect(sheet.cells.size).toBe(2);
	});

	it('is undoable as part of the transaction', () => {
		const wb = createWorkbook();
		const entry = apply(wb, 'A1:B2', { top: thin, insideVertical: thin });
		expect(wb.sheets[0].cells.size).toBe(4);
		entry.undo();
		expect(wb.sheets[0].cells.size).toBe(0);
		entry.redo();
		expect(borderAt(wb, 'A1')).toEqual({ top: thin, right: thin });
	});

	it('does nothing when no edge changes', () => {
		const wb = createWorkbook();
		apply(wb, 'A1:B2', {});
		expect(wb.sheets[0].cells.size).toBe(0);
	});
});

describe('describeRangeBorders', () => {
	it('reports one line, none or mixed per edge', () => {
		const wb = createWorkbook();
		apply(wb, 'A1:B2', { top: thin, left: thickRed, insideHorizontal: thin });
		apply(wb, 'B1', { right: thin });
		const state = describeRangeBorders(wb, 0, r('A1:B2'));
		expect(state.top).toEqual(thin);
		expect(state.left).toEqual(thickRed);
		expect(state.insideHorizontal).toEqual(thin);
		expect(state.bottom).toBeNull();
		expect(state.insideVertical).toBeNull();
		expect(state.right).toBe('mixed');
		expect(state.diagonalUp).toBeNull();
	});

	it('reads diagonals and samples whole columns quickly', () => {
		const wb = createWorkbook();
		apply(wb, 'C:C', { diagonalUp: thin, left: thin });
		const state = describeRangeBorders(wb, 0, r('C:C'));
		expect(state.diagonalUp).toEqual(thin);
		expect(state.diagonalDown).toBeNull();
		expect(state.left).toEqual(thin);
	});
});
