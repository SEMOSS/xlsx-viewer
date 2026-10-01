import { describe, expect, it } from 'vitest';

import { CalcEngine } from './engine';
import { calculated, edit, makeWorkbook, refs, valueAt, valuesAt } from './test-utils/workbook';

describe('circular references', () => {
	it('shows 0 and reports the cells of a cycle', () => {
		const { wb, result } = calculated({
			Sheet1: { A1: '=B1+1', B1: '=C1+1', C1: '=A1+1', D1: '=A1*2', E1: 5 },
		});
		expect(valuesAt(wb, 'A1:D1')).toEqual([[0, 0, 0, 0]]);
		expect(refs(wb, result.circular)).toEqual(['Sheet1!A1', 'Sheet1!B1', 'Sheet1!C1']);
	});

	it('detects self references, including through ranges', () => {
		const { wb, result } = calculated({
			Sheet1: { A1: 1, A2: '=A2+1', A5: '=SUM(A:A)', B1: '=ROW(B1)' },
		});
		expect(valueAt(wb, 'A2')).toBe(0);
		expect(valueAt(wb, 'A5')).toBe(0);
		expect(valueAt(wb, 'B1')).toBe(1);
		expect(refs(wb, result.circular)).toEqual(['Sheet1!A2', 'Sheet1!A5']);
	});

	it('detects cycles across sheets and through names', () => {
		const { wb, result } = calculated(
			{ Sheet1: { A1: '=Sheet2!A1' }, Sheet2: { A1: '=Loopy' } },
			{},
			(w) => w.definedNames.push({ name: 'Loopy', formula: 'Sheet1!$A$1*2' }),
		);
		expect(valueAt(wb, 'A1')).toBe(0);
		expect(refs(wb, result.circular)).toEqual(['Sheet1!A1', 'Sheet2!A1']);
	});

	it('detects cycles that only appear at run time (INDIRECT)', () => {
		const { wb, result } = calculated({ Sheet1: { A1: '=INDIRECT("B1")+1', B1: '=A1+1' } });
		expect(valuesAt(wb, 'A1:B1')).toEqual([[0, 0]]);
		expect(result.circular.length).toBe(2);
	});

	it('recovers when the cycle is broken', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: '=B1+1', B1: '=A1+1' } });
		const r = edit(engine, 'B1', 10);
		expect(valueAt(wb, 'A1')).toBe(11);
		expect(r.circular).toEqual([]);
		const again = edit(engine, 'B1', '=A1+1');
		expect(refs(wb, again.circular)).toEqual(['Sheet1!A1', 'Sheet1!B1']);
	});

	it('detects a dynamic array reading its own spill area', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: '=SEQUENCE(3)+A2' } });
		expect(valueAt(wb, 'A1')).toBe(0);
		expect(valueAt(wb, 'A2')).toBe(null);
		const again = engine.recalculateAll();
		expect(refs(wb, again.circular)).toEqual(['Sheet1!A1']);
		expect(valueAt(wb, 'A1')).toBe(0);
	});

	it('detects cycles through a spill', () => {
		const { wb, result } = calculated({ Sheet1: { A1: '=SEQUENCE(B1)', B1: '=SUM(A2:A3)+2' } });
		expect(valueAt(wb, 'A1')).toBe(0);
		expect(result.circular.length).toBe(2);
	});
});

describe('iterative calculation', () => {
	it('iterates a cycle until it converges', () => {
		const { wb, result } = calculated({ Sheet1: { A1: '=B1/2+1', B1: '=A1' } }, {}, (w) => {
			w.calc.iterate = true;
			w.calc.iterateCount = 200;
			w.calc.iterateDelta = 1e-9;
		});
		expect(valueAt(wb, 'A1')).toBeCloseTo(2, 8);
		expect(valueAt(wb, 'B1')).toBeCloseTo(2, 8);
		expect(result.circular).toEqual([]);
	});

	it('stops after iterateCount iterations per recalculation', () => {
		const wb = makeWorkbook({ Sheet1: { A1: '=A1+1', B1: '=A1*10' } });
		wb.calc.iterate = true;
		wb.calc.iterateCount = 10;
		const engine = new CalcEngine(wb);
		engine.rebuild();
		expect(valueAt(wb, 'A1')).toBe(10);
		expect(valueAt(wb, 'B1')).toBe(100);
		// Like F9 in Excel, every recalculation iterates again from the current values.
		engine.recalculateAll();
		expect(valueAt(wb, 'A1')).toBe(20);
	});

	it('uses Excel defaults (100 iterations, delta 0.001)', () => {
		const wb = makeWorkbook({ Sheet1: { A1: '=A1+1' } });
		wb.calc.iterate = true;
		new CalcEngine(wb).rebuild();
		expect(valueAt(wb, 'A1')).toBe(100);
	});

	it('solves a circular interest model', () => {
		// Interest on the average balance, which itself includes the interest.
		const { wb } = calculated(
			{ Sheet1: { A1: 1000, B1: '=(A1+C1)*0.1', C1: '=A1+B1' } },
			{},
			(w) => {
				w.calc.iterate = true;
				w.calc.iterateDelta = 1e-12;
			},
		);
		const interest = valueAt(wb, 'B1') as number;
		expect(interest).toBeCloseTo((1000 * 0.2) / 0.9, 6);
	});
});
