import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { createWorkbook } from '../sheet/workbook';
import { CalcEngine } from './engine';

/**
 * Performance sanity checks. Targets: rebuild of 100k formulas < 2 s, an
 * edit with 10k dependents < 100 ms, recalculateAll of 100k simple formulas
 * < 1.5 s. The limits asserted here are generous (slow CI machines); typical
 * timings are an order of magnitude lower.
 */

function time<T>(f: () => T): { ms: number; value: T } {
	const start = performance.now();
	const value = f();
	return { ms: performance.now() - start, value };
}

describe('performance', () => {
	it('rebuilds and recalculates 100k formulas quickly', () => {
		const wb = createWorkbook();
		const cells = wb.sheets[0].cells;
		for (let i = 0; i < 100_000; i++) {
			cells.set(cellKey(i, 0), { value: i });
			cells.set(cellKey(i, 1), { value: i * 2 + 1, formula: `A${i + 1}*2+1` });
			cells.set(cellKey(i, 2), {
				value: i * 2 + 1,
				formula: `IF(B${i + 1}>10,B${i + 1}-A${i + 1},0)`,
			});
		}
		const engine = new CalcEngine(wb);
		const rebuild = time(() => engine.rebuild());
		expect(rebuild.ms).toBeLessThan(6000);
		const all = time(() => engine.recalculateAll());
		expect(all.ms).toBeLessThan(4500);
		expect(cells.get(cellKey(99_999, 2))?.value).toBe(100_000);
	});

	it('propagates an edit to 10k dependents quickly', () => {
		const wb = createWorkbook();
		const cells = wb.sheets[0].cells;
		cells.set(cellKey(0, 0), { value: 1 });
		for (let i = 0; i < 10_000; i++) {
			cells.set(cellKey(i, 1), { value: null, formula: `$A$1+${i}` });
			cells.set(cellKey(i, 2), { value: null, formula: `B${i + 1}*2` });
		}
		const engine = new CalcEngine(wb);
		engine.rebuild();
		cells.set(cellKey(0, 0), { value: 5 });
		const edit = time(() => engine.onCellsChanged([{ sheet: 0, row: 0, col: 0 }]));
		expect(edit.value.changed.length).toBe(20_000);
		expect(edit.ms).toBeLessThan(500);
		expect(cells.get(cellKey(9_999, 2))?.value).toBe(20_008);
	});

	it('keeps unrelated edits cheap in a large model', () => {
		const wb = createWorkbook();
		const cells = wb.sheets[0].cells;
		for (let i = 0; i < 50_000; i++) {
			cells.set(cellKey(i, 0), { value: i });
			cells.set(cellKey(i, 1), { value: null, formula: `SUM(A${i + 1}:A${i + 3})` });
		}
		const engine = new CalcEngine(wb);
		engine.rebuild();
		cells.set(cellKey(25_000, 0), { value: -1 });
		const edit = time(() => engine.onCellsChanged([{ sheet: 0, row: 25_000, col: 0 }]));
		expect(edit.value.changed.length).toBe(3);
		expect(edit.ms).toBeLessThan(50);
	});

	it('handles 100k-long dependency chains', () => {
		const wb = createWorkbook();
		const cells = wb.sheets[0].cells;
		cells.set(cellKey(0, 0), { value: 1 });
		for (let i = 1; i < 100_000; i++) {
			cells.set(cellKey(i, 0), { value: null, formula: `A${i}+1` });
		}
		const engine = new CalcEngine(wb);
		const build = time(() => engine.rebuild());
		expect(cells.get(cellKey(99_999, 0))?.value).toBe(100_000);
		expect(build.ms).toBeLessThan(6000);
		cells.set(cellKey(0, 0), { value: 0 });
		const edit = time(() => engine.onCellsChanged([{ sheet: 0, row: 0, col: 0 }]));
		expect(cells.get(cellKey(99_999, 0))?.value).toBe(99_999);
		expect(edit.ms).toBeLessThan(3000);
	});

	it('aggregates whole columns without visiting empty rows', () => {
		const wb = createWorkbook();
		const cells = wb.sheets[0].cells;
		for (let i = 0; i < 20_000; i++) {
			cells.set(cellKey(i * 50, 0), { value: 1 });
		}
		for (let i = 0; i < 1_000; i++) {
			cells.set(cellKey(i, 3), { value: null, formula: 'SUM(A:A)+COUNT(A:A)' });
		}
		const engine = new CalcEngine(wb);
		const build = time(() => engine.rebuild());
		expect(cells.get(cellKey(0, 3))?.value).toBe(40_000);
		expect(build.ms).toBeLessThan(5000);
	});

	it('keeps running totals over formula columns free of quadratic bookkeeping', () => {
		const wb = createWorkbook();
		const cells = wb.sheets[0].cells;
		cells.set(cellKey(0, 0), { value: 1 });
		for (let i = 0; i < 5_000; i++) {
			if (i > 0) {
				cells.set(cellKey(i, 0), { value: null, formula: `A${i}+1` });
			}
			cells.set(cellKey(i, 1), { value: null, formula: `SUM($A$1:A${i + 1})` });
		}
		const engine = new CalcEngine(wb);
		const build = time(() => engine.rebuild());
		expect(cells.get(cellKey(4_999, 1))?.value).toBe(12_502_500);
		expect(build.ms).toBeLessThan(3000);
		cells.set(cellKey(0, 0), { value: 2 });
		const edit = time(() => engine.onCellsChanged([{ sheet: 0, row: 0, col: 0 }]));
		expect(cells.get(cellKey(4_999, 1))?.value).toBe(12_507_500);
		expect(edit.ms).toBeLessThan(3000);
	});
});
