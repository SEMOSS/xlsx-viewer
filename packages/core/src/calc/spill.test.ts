import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { ERRORS } from '../values/errors';
import { CalcEngine } from './engine';
import {
	calculated,
	cellAt,
	edit,
	evalIn,
	locate,
	makeWorkbook,
	refs,
	valueAt,
	valuesAt,
} from './test-utils/workbook';

const B1 = cellKey(0, 1);

describe('dynamic-array spills', () => {
	it('spills, grows and shrinks with the result', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: 3, B1: '=SEQUENCE(A1)', C1: '=B3' } });
		expect(valuesAt(wb, 'B1:B4')).toEqual([[1], [2], [3], [null]]);
		expect(cellAt(wb, 'B1')?.array).toEqual({
			ref: { top: 0, left: 1, bottom: 2, right: 1 },
			dynamic: true,
		});
		expect(cellAt(wb, 'B3')).toEqual({ value: 3, arrayAnchor: B1 });
		expect(valueAt(wb, 'C1')).toBe(3);
		expect(engine.getSpillRange(locate(wb, 'B1'))).toEqual({
			top: 0,
			left: 1,
			bottom: 2,
			right: 1,
		});
		expect(engine.isDynamicResult(locate(wb, 'B1'))).toBe(true);

		const grow = edit(engine, 'A1', 5);
		expect(valuesAt(wb, 'B1:B6')).toEqual([[1], [2], [3], [4], [5], [null]]);
		expect(refs(wb, grow.changed)).toEqual(['Sheet1!B4', 'Sheet1!B5']);

		const shrink = edit(engine, 'A1', 2);
		expect(valuesAt(wb, 'B1:B6')).toEqual([[1], [2], [null], [null], [null], [null]]);
		expect(refs(wb, shrink.changed)).toEqual(['Sheet1!C1', 'Sheet1!B3', 'Sheet1!B4', 'Sheet1!B5']);
		expect(valueAt(wb, 'C1')).toBe(0);
		expect(cellAt(wb, 'B3')).toBeUndefined();

		edit(engine, 'A1', 1);
		expect(cellAt(wb, 'B1')?.array?.ref).toEqual({ top: 0, left: 1, bottom: 0, right: 1 });
		expect(cellAt(wb, 'B2')).toBeUndefined();
	});

	it('shows #SPILL! when blocked and spills again when unblocked', () => {
		const { wb, engine } = calculated({ Sheet1: { B1: '=SEQUENCE(3)', B3: 'x', D1: '=SUM(B1#)' } });
		expect(valueAt(wb, 'B1')).toEqual(ERRORS.SPILL);
		expect(valueAt(wb, 'B2')).toBe(null);
		expect(valueAt(wb, 'D1')).toEqual(ERRORS.REF);
		expect(engine.getBlockedSpillRange(locate(wb, 'B1'))).toEqual({
			top: 0,
			left: 1,
			bottom: 2,
			right: 1,
		});
		const r = edit(engine, 'B3', undefined);
		expect(valuesAt(wb, 'B1:B3')).toEqual([[1], [2], [3]]);
		expect(valueAt(wb, 'D1')).toBe(6);
		expect(refs(wb, r.changed)).toEqual(['Sheet1!B1', 'Sheet1!D1', 'Sheet1!B2', 'Sheet1!B3']);
	});

	it('is blocked by typing into the spill range and recovers on clear', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: '={1;2;3}' } });
		edit(engine, 'A2', 'typed');
		expect(valuesAt(wb, 'A1:A3')).toEqual([[ERRORS.SPILL], ['typed'], [null]]);
		edit(engine, 'A2', undefined);
		expect(valuesAt(wb, 'A1:A3')).toEqual([[1], [2], [3]]);
		edit(engine, 'A3', '=1+1');
		expect(valueAt(wb, 'A1')).toEqual(ERRORS.SPILL);
		expect(valueAt(wb, 'A2')).toBe(null);
		expect(valueAt(wb, 'A3')).toBe(2);
	});

	it('removes spilled values when the anchor is cleared, keeping styles', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: '=SEQUENCE(1,3)' } }, {}, (w) => {
			w.sheets[0].cells.set(cellKey(0, 2), { value: null, style: 4 });
		});
		expect(cellAt(wb, 'C1')).toEqual({ value: 3, arrayAnchor: 0, style: 4 });
		const r = edit(engine, 'A1', undefined);
		expect(cellAt(wb, 'B1')).toBeUndefined();
		expect(cellAt(wb, 'C1')).toEqual({ value: null, style: 4 });
		expect(refs(wb, r.changed)).toEqual(['Sheet1!B1', 'Sheet1!C1']);
	});

	it('recalculates readers of spilled cells, including cells the spill grows into', () => {
		const { wb, engine } = calculated({
			Sheet1: { A1: 2, B1: '=SEQUENCE(A1)', C1: '=SUM(B4:B5)', D1: '=ROWS(B1#)', E1: '=B1#*10' },
		});
		expect(valueAt(wb, 'C1')).toBe(0);
		expect(valueAt(wb, 'D1')).toBe(2);
		expect(valuesAt(wb, 'E1:E2')).toEqual([[10], [20]]);
		edit(engine, 'A1', 5);
		expect(valueAt(wb, 'C1')).toBe(9);
		expect(valueAt(wb, 'D1')).toBe(5);
		expect(valuesAt(wb, 'E1:E5')).toEqual([[10], [20], [30], [40], [50]]);
	});

	it('lets a blocked formula spill once another spill shrinks away', () => {
		const { wb, engine } = calculated({
			Sheet1: { A1: 3, B1: '=SEQUENCE(A1)', C2: '=SEQUENCE(1,1)', A3: '=SEQUENCE(1,2)' },
		});
		// A3 wants A3:B3, but B3 belongs to B1's spill.
		expect(valueAt(wb, 'A3')).toEqual(ERRORS.SPILL);
		edit(engine, 'A1', 2);
		expect(valuesAt(wb, 'A3:B3')).toEqual([[1, 2]]);
		expect(valuesAt(wb, 'B1:B2')).toEqual([[1], [2]]);
	});

	it('detects merged cells, tables and the grid edge as obstacles', () => {
		const merged = calculated({ Sheet1: { A1: '=SEQUENCE(2)' } }, {}, (w) =>
			w.sheets[0].merges.push({ top: 1, left: 0, bottom: 1, right: 1 }),
		);
		expect(valueAt(merged.wb, 'A1')).toEqual(ERRORS.SPILL);
		const edge = calculated({ Sheet1: { XFD1: '=SEQUENCE(1,2)' } });
		expect(valueAt(edge.wb, 'XFD1')).toEqual(ERRORS.SPILL);
		const table = calculated({ Sheet1: { A1: '=SEQUENCE(3)' } }, {}, (w) =>
			w.sheets[0].tables.push({
				id: 1,
				name: 'T',
				displayName: 'T',
				ref: { top: 2, left: 0, bottom: 3, right: 0 },
				headerRowCount: 1,
				totalsRowCount: 0,
				columns: [{ id: 1, name: 'c' }],
			}),
		);
		expect(valueAt(table.wb, 'A1')).toEqual(ERRORS.SPILL);
	});

	it('restores spills after an undo puts the old cells back', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: '=SEQUENCE(3)' } });
		const before = cellAt(wb, 'A2');
		edit(engine, 'A2', 'blocker');
		expect(valueAt(wb, 'A1')).toEqual(ERRORS.SPILL);
		// Undo restores the previous cell object (still owned by the anchor).
		wb.sheets[0].cells.set(cellKey(1, 0), before!);
		engine.onCellsChanged([locate(wb, 'A2')]);
		expect(valuesAt(wb, 'A1:A3')).toEqual([[1], [2], [3]]);
	});

	it('supports spill references to loaded dynamic arrays and errors for others', () => {
		const wb = makeWorkbook({
			Sheet1: {
				A1: { f: 'SEQUENCE(2)', dynamic: true, spill: 'A1:A2', v: 1 },
				A2: 2,
				B1: { f: 'SUM(A1#)', v: 0 },
				C1: { f: 'A2#', v: 0 },
				D1: '=7',
				E1: '=D1#',
			},
		});
		wb.sheets[0].cells.set(cellKey(1, 0), { value: 2, arrayAnchor: 0 });
		const engine = new CalcEngine(wb);
		engine.rebuild();
		engine.recalculateAll();
		expect(valueAt(wb, 'B1')).toBe(3);
		expect(valueAt(wb, 'C1')).toEqual(ERRORS.REF);
		expect(valueAt(wb, 'E1')).toEqual(ERRORS.REF);
		expect(evalIn(engine, 'ANCHORARRAY(A1)')).toEqual([[1], [2]]);
	});

	it('ignores orphaned cells of arrays that no longer exist', () => {
		const { wb } = calculated({ Sheet1: { A1: '=SEQUENCE(3)' } }, {}, (w) => {
			// Left behind by an unnotified edit: claims to belong to a missing anchor at B1.
			w.sheets[0].cells.set(cellKey(1, 0), { value: 99, arrayAnchor: cellKey(0, 1) });
		});
		expect(valuesAt(wb, 'A1:A3')).toEqual([[1], [2], [3]]);
	});

	it('spills references and blank cells as zeros', () => {
		const { wb } = calculated({
			Sheet1: { A1: 1, A3: 3, B1: '=A1:A3', C1: '=FILTER(A1:A3,A1:A3<>5)' },
		});
		expect(valuesAt(wb, 'B1:B3')).toEqual([[1], [0], [3]]);
		expect(valuesAt(wb, 'C1:C3')).toEqual([[1], [0], [3]]);
	});
});
