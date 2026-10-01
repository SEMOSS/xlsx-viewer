import { describe, expect, it } from 'vitest';

import { CalcEngine } from '../../calc/engine';
import {
	calculated,
	cellAt,
	evalIn,
	makeWorkbook,
	valueAt,
	valuesAt,
} from '../../calc/test-utils/workbook';
import { ERRORS } from '../../values/errors';

/**
 * The same formula behaves differently by mode: dynamic-array formulas
 * spill, legacy formulas implicitly intersect references and use the first
 * element of arrays, CSE array formulas fill their block.
 */
describe('array modes', () => {
	const sheet = { A1: 1, A2: 2, A3: 3, A4: 4 };

	it('dynamic formulas spill array results', () => {
		const { wb } = calculated({ Sheet1: { ...sheet, B1: '=A1:A3*10', C2: '=ROW(A1:A2)' } });
		expect(valuesAt(wb, 'B1:B3')).toEqual([[10], [20], [30]]);
		// Spilled cells point at the anchor's cell key (B1 = row 0, column 1).
		expect(cellAt(wb, 'B2')?.arrayAnchor).toBe(1);
		expect(valuesAt(wb, 'C2:C3')).toEqual([[1], [2]]);
	});

	it('legacy formulas implicitly intersect references', () => {
		const { wb } = calculated({
			Sheet1: {
				...sheet,
				B2: { f: 'A1:A4*10', v: 0 },
				B6: { f: 'A1:A4*10', v: 0 },
				C3: { f: 'A1:A4', v: 0 },
				D1: { f: '{5,6,7}', v: 0 },
				E3: { f: 'SUM(A1:A4)', v: 0 },
				F3: { f: 'ABS(A1:A4)', v: 0 },
				G1: { f: 'A1:D1', v: 0 },
			},
		});
		expect(valueAt(wb, 'B2')).toBe(20);
		expect(valueAt(wb, 'B6')).toEqual(ERRORS.VALUE);
		expect(valueAt(wb, 'C3')).toBe(3);
		expect(valueAt(wb, 'D1')).toBe(5);
		expect(valueAt(wb, 'E3')).toBe(10);
		expect(valueAt(wb, 'F3')).toBe(3);
		expect(valueAt(wb, 'G1')).toEqual(ERRORS.VALUE);
		expect(valueAt(wb, 'B3')).toBe(null);
	});

	it('legacy formulas still evaluate arrays element-wise, and array parameters as arrays', () => {
		const { wb } = calculated({
			Sheet1: {
				...sheet,
				B1: 'x',
				B2: 'y',
				B3: 'x',
				B4: 'z',
				C1: { f: 'SUM(COUNTIF(B1:B4,{"x","z"}))', v: 0 },
				C2: { f: 'SUMPRODUCT(1/COUNTIF(B1:B4,B1:B4))', v: 0 },
				C3: { f: 'SUMPRODUCT((A1:A4>1)*A1:A4)', v: 0 },
				C4: { f: 'SUM(IF({TRUE,FALSE,TRUE},{1,2,3}))', v: 0 },
				C5: { f: 'SUM(A1:A4*2)', v: 0 },
				C9: { f: 'MAX(ABS({-7,3}))', v: 0 },
				D1: { f: 'INDEX(A1:A4*10,3)', v: 0 },
				D2: { f: 'LOOKUP(2,1/(B1:B4="x"),A1:A4)', v: 0 },
				D3: { f: 'MATCH(1,INDEX((B1:B4="x")*(A1:A4>1),0),0)', v: 0 },
			},
		});
		expect(valueAt(wb, 'D1')).toBe(30);
		expect(valueAt(wb, 'D2')).toBe(3);
		expect(valueAt(wb, 'D3')).toBe(3);
		expect(valueAt(wb, 'C1')).toBe(3);
		expect(valueAt(wb, 'C2')).toBe(3);
		expect(valueAt(wb, 'C3')).toBe(9);
		expect(valueAt(wb, 'C4')).toBe(4);
		// A range operand in a legacy formula is implicitly intersected (row 5: none).
		expect(valueAt(wb, 'C5')).toEqual(ERRORS.VALUE);
		expect(valueAt(wb, 'C9')).toBe(7);
	});

	it('CSE arrays fill their block and pad with #N/A', () => {
		const { wb } = calculated({
			Sheet1: {
				...sheet,
				B1: { f: 'A1:A3*10', cse: 'B1:C4' },
				D1: { f: 'SUM(A1:A4*A1:A4)', cse: 'D1' },
				E1: { f: '{1,2}', cse: 'E1:F3' },
				G1: { f: '7', cse: 'G1:G2' },
			},
		});
		expect(valuesAt(wb, 'B1:C4')).toEqual([
			[10, 10],
			[20, 20],
			[30, 30],
			[ERRORS.NA, ERRORS.NA],
		]);
		expect(valueAt(wb, 'D1')).toBe(30);
		expect(valuesAt(wb, 'E1:F3')).toEqual([
			[1, 2],
			[1, 2],
			[1, 2],
		]);
		expect(valuesAt(wb, 'G1:G2')).toEqual([[7], [7]]);
		expect(cellAt(wb, 'C2')?.arrayAnchor).toBe(1);
	});

	it('recalculates readers of CSE blocks after the block', () => {
		const wb = makeWorkbook({
			Sheet1: {
				A1: { f: 'SUM(C2:C3)', v: 5 },
				B1: 1,
				C1: { f: 'B1*{1;2;3}', cse: 'C1:C3', v: 1 },
				C2: 2,
				C3: 3,
			},
		});
		wb.sheets[0].cells.set(1, { value: 10 });
		const engine = new CalcEngine(wb);
		engine.rebuild();
		engine.recalculateAll();
		expect(valuesAt(wb, 'C1:C3')).toEqual([[10], [20], [30]]);
		expect(valueAt(wb, 'A1')).toBe(50);
	});

	it('evaluateFormula honours dynamicArrays: false', () => {
		const { engine } = calculated({ Sheet1: sheet });
		expect(evalIn(engine, 'A1:A4*2', 'Sheet1!C3', false)).toBe(6);
		expect(evalIn(engine, 'A1:A4*2', 'Sheet1!C3', true)).toEqual([[2], [4], [6], [8]]);
		expect(evalIn(engine, 'A1:A4', 'Sheet1!C9', false)).toEqual(ERRORS.VALUE);
		expect(evalIn(engine, '{1,2}', 'Sheet1!C3', false)).toBe(1);
	});

	it('a blank or empty result is shown as 0', () => {
		const { wb } = calculated({
			Sheet1: { A1: '=Z1', A2: '=Z1:Z2', A4: '=IF(TRUE,)', B1: { f: 'Z9', v: 5 } },
		});
		expect(valueAt(wb, 'A1')).toBe(0);
		expect(valuesAt(wb, 'A2:A3')).toEqual([[0], [0]]);
		expect(valueAt(wb, 'A4')).toBe(0);
		expect(valueAt(wb, 'B1')).toBe(0);
	});

	it('LAMBDA values and unions are errors in cells', () => {
		const { wb } = calculated({
			Sheet1: { A1: '=LAMBDA(x,x)', A2: '=(B1,C1)', A3: '=LAMBDA(x,x)(4)' },
		});
		expect(valueAt(wb, 'A1')).toEqual(ERRORS.CALC);
		expect(valueAt(wb, 'A2')).toEqual(ERRORS.VALUE);
		expect(valueAt(wb, 'A3')).toBe(4);
	});
});
