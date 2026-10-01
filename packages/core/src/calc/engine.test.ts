import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { ERRORS } from '../values/errors';
import { CalcEngine } from './engine';
import {
	calculated,
	cellAt,
	edit,
	locate,
	makeWorkbook,
	refs,
	valueAt,
} from './test-utils/workbook';

describe('CalcEngine: recalculation', () => {
	it('computes multi-sheet models in dependency order', () => {
		const { wb } = calculated({
			Inputs: { A1: 100, A2: 0.2 },
			Model: { A1: '=Inputs!A1*(1+Inputs!A2)', A2: '=A1-Inputs!A1', A3: '=Summary!A1/2' },
			Summary: { A1: '=SUM(Model!A1:A2)' },
		});
		expect(valueAt(wb, 'Model!A1')).toBe(120);
		expect(valueAt(wb, 'Model!A2')).toBe(20);
		expect(valueAt(wb, 'Summary!A1')).toBe(140);
		expect(valueAt(wb, 'Model!A3')).toBe(70);
	});

	it('recomputes only dependents on edit and reports changed cells', () => {
		const { wb, engine } = calculated({
			Sheet1: { A1: 1, A2: 2, B1: '=A1*10', B2: '=A2*10', C1: '=B1+B2', D1: '=1+1' },
		});
		const r = edit(engine, 'A1', 5);
		expect(refs(wb, r.changed)).toEqual(['Sheet1!B1', 'Sheet1!C1']);
		expect(valueAt(wb, 'C1')).toBe(70);
		expect(r.circular).toEqual([]);
		const same = edit(engine, 'A2', 2);
		expect(same.changed).toEqual([]);
	});

	it('tracks ranges over whole columns and rows', () => {
		const { wb, engine } = calculated({
			Sheet1: { A1: 1, A2: 2, C1: '=SUM(A:A)', C2: '=COUNT(5:5)', C3: '=SUM(Sheet2!B:B)' },
			Sheet2: { B7: 7 },
		});
		expect(valueAt(wb, 'C1')).toBe(3);
		edit(engine, 'A50000', 10);
		expect(valueAt(wb, 'C1')).toBe(13);
		edit(engine, 'Z5', 1);
		edit(engine, 'X5', 1);
		expect(valueAt(wb, 'C2')).toBe(2);
		edit(engine, 'Sheet2!B1000', 3);
		expect(valueAt(wb, 'C3')).toBe(10);
	});

	it('handles new, changed and removed formulas', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: 3, B1: '=A1*2', C1: '=B1+1' } });
		edit(engine, 'B1', '=A1*3');
		expect(valueAt(wb, 'C1')).toBe(10);
		edit(engine, 'B1', 100);
		expect(valueAt(wb, 'C1')).toBe(101);
		edit(engine, 'A1', 4);
		expect(valueAt(wb, 'C1')).toBe(101);
		edit(engine, 'B1', undefined);
		expect(valueAt(wb, 'C1')).toBe(1);
		edit(engine, 'D1', '=C1*2');
		expect(valueAt(wb, 'D1')).toBe(2);
	});

	it('handles long chains without recursion', () => {
		const n = 50_000;
		const wb = makeWorkbook({ Sheet1: { A1: 1 } });
		const cells = wb.sheets[0].cells;
		// Inserted in reverse so the scheduler cannot rely on cell order.
		for (let i = n; i >= 2; i--) {
			cells.set(cellKey(i - 1, 0), { value: null, formula: `A${i - 1}+1` });
		}
		const engine = new CalcEngine(wb);
		engine.rebuild();
		expect(cells.get(cellKey(n - 1, 0))?.value).toBe(n);
		const r = edit(engine, 'A1', 10);
		expect(r.changed.length).toBe(n - 1);
		expect(cells.get(cellKey(n - 1, 0))?.value).toBe(n + 9);
	});

	it('resolves dynamic references (INDIRECT, OFFSET) in the right order', () => {
		const { wb, engine } = calculated({
			Sheet1: {
				A1: '=INDIRECT("B"&C1)*2',
				B1: '=D1+1',
				B2: '=D1+2',
				C1: 1,
				D1: 10,
				E1: '=SUM(OFFSET(B1,0,0,2))',
			},
		});
		expect(valueAt(wb, 'A1')).toBe(22);
		expect(valueAt(wb, 'E1')).toBe(23);
		edit(engine, 'C1', 2);
		expect(valueAt(wb, 'A1')).toBe(24);
		edit(engine, 'D1', 0);
		expect(valueAt(wb, 'A1')).toBe(4);
		expect(valueAt(wb, 'E1')).toBe(3);
	});

	it('rebuild trusts cached values and computes formulas without one', () => {
		const wb = makeWorkbook({
			Sheet1: { A1: 1, B1: { f: 'A1+1', v: 99 }, C1: { f: 'A1+2' }, D1: { f: 'B1*2', v: 7 } },
		});
		const engine = new CalcEngine(wb);
		engine.rebuild();
		expect(valueAt(wb, 'B1')).toBe(99);
		expect(valueAt(wb, 'C1')).toBe(3);
		expect(valueAt(wb, 'D1')).toBe(7);
		const r = engine.recalculateAll();
		expect(valueAt(wb, 'B1')).toBe(2);
		expect(valueAt(wb, 'D1')).toBe(4);
		expect(refs(wb, r.changed)).toEqual(['Sheet1!B1', 'Sheet1!D1']);
	});

	it('honours fullCalcOnLoad on the first rebuild only', () => {
		const wb = makeWorkbook({ Sheet1: { A1: 1, B1: { f: 'A1+1', v: 99 } } });
		wb.calc.fullCalcOnLoad = true;
		const engine = new CalcEngine(wb);
		engine.rebuild();
		expect(valueAt(wb, 'B1')).toBe(2);
		wb.sheets[0].cells.set(cellKey(0, 1), { value: 50, formula: 'A1+1' });
		engine.rebuild();
		expect(valueAt(wb, 'B1')).toBe(50);
	});

	it('sets sheet.dirty only for sheets whose values changed', () => {
		const wb = makeWorkbook({
			Sheet1: { A1: 1, B1: { f: 'A1*2', v: 2 } },
			Sheet2: { A1: { f: '5', v: 5 } },
		});
		wb.sheets.forEach((s) => (s.dirty = false));
		const engine = new CalcEngine(wb);
		engine.rebuild();
		engine.recalculateAll();
		expect(wb.sheets.map((s) => s.dirty)).toEqual([false, false]);
		wb.sheets[0].cells.set(cellKey(0, 0), { value: 3 });
		engine.onCellsChanged([{ sheet: 0, row: 0, col: 0 }]);
		expect(wb.sheets.map((s) => s.dirty)).toEqual([true, false]);
	});

	it('never mutates cell objects', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: 1, B1: '=A1+1', C1: '=SEQUENCE(2)' } });
		const before = cellAt(wb, 'B1')!;
		const snapshot = { ...before };
		edit(engine, 'A1', 2);
		expect(before).toEqual(snapshot);
		expect(cellAt(wb, 'B1')).not.toBe(before);
		expect(Object.isFrozen(ERRORS.NA)).toBe(true);
	});

	it('rebuilds by itself when sheets, names or tables change', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: 2, B1: '=A1*Factor' } });
		expect(valueAt(wb, 'B1')).toEqual(ERRORS.NAME);
		wb.definedNames.push({ name: 'Factor', formula: '10' });
		edit(engine, 'A1', 3);
		expect(valueAt(wb, 'B1')).toBe(30);
		wb.definedNames[0] = { name: 'Factor', formula: '100' };
		engine.onCellsChanged([]);
		expect(valueAt(wb, 'B1')).toBe(300);
		wb.sheets[0].name = 'Renamed';
		expect(engine.evaluateFormula('Renamed!A1', 0, 5, 5)).toMatchObject({ kind: 'ref' });
		expect(engine.recalculateAll().changed).toEqual([]);
	});

	it('survives the cell map being replaced (undo of a sheet snapshot)', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: 1, A2: 2, B1: '=SUM(A:A)' } });
		const saved = new Map(wb.sheets[0].cells);
		edit(engine, 'A3', 3);
		expect(valueAt(wb, 'B1')).toBe(6);
		wb.sheets[0].cells = new Map(saved);
		engine.onCellsChanged([locate(wb, 'A3')]);
		expect(valueAt(wb, 'B1')).toBe(3);
	});
});

describe('CalcEngine: graph queries', () => {
	const { engine, wb } = calculated(
		{
			Sheet1: { A1: 1, A2: 2, B1: '=SUM(A1:A2)+Sheet2!A1', C1: '=B1*Rate', D1: '=A1' },
			Sheet2: { A1: 5 },
		},
		{},
		(w) => w.definedNames.push({ name: 'Rate', formula: 'Sheet2!$B$1' }),
	);

	it('lists direct precedents with names expanded', () => {
		expect(engine.getPrecedents(locate(wb, 'B1'))).toEqual([
			{ sheet: 1, range: { top: 0, left: 0, bottom: 0, right: 0 } },
			{ sheet: 0, range: { top: 0, left: 0, bottom: 1, right: 0 } },
		]);
		expect(engine.getPrecedents(locate(wb, 'C1'))).toEqual([
			{ sheet: 0, range: { top: 0, left: 1, bottom: 0, right: 1 } },
			{ sheet: 1, range: { top: 0, left: 1, bottom: 0, right: 1 } },
		]);
		expect(engine.getPrecedents(locate(wb, 'A1'))).toEqual([]);
	});

	it('lists direct dependents', () => {
		expect(refs(wb, engine.getDependents(locate(wb, 'A1')))).toEqual(['Sheet1!B1', 'Sheet1!D1']);
		expect(refs(wb, engine.getDependents(locate(wb, 'Sheet2!B1')))).toEqual(['Sheet1!C1']);
		expect(refs(wb, engine.getDependents(locate(wb, 'B1')))).toEqual(['Sheet1!C1']);
		expect(engine.getDependents(locate(wb, 'Z9'))).toEqual([]);
	});
});
