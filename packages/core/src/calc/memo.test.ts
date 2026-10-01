import { describe, expect, it } from 'vitest';

import { getDefaultRegistry } from '../formula/registry';
import type { FormulaValue, FunctionContext, RefValue } from '../formula/types';
import { isRefValue } from '../formula/types';
import { ERRORS } from '../values/errors';
import { calculated, edit, valueAt } from './test-utils/workbook';

/** Context method exposed by the evaluator's frames. */
type WithMemo = FunctionContext & { memo<T>(ref: RefValue, key: string, build: () => T): T };

let builds = 0;
getDefaultRegistry().register({
	name: 'TESTMEMOSUM',
	minArgs: 1,
	maxArgs: 1,
	params: ['ref'],
	impl: (args, ctx): FormulaValue => {
		const ref = args[0];
		if (!isRefValue(ref)) {
			return ERRORS.VALUE;
		}
		return (ctx as WithMemo).memo(ref, 'sum', () => {
			builds++;
			let total = 0;
			ctx.forEachCell(ref, (v) => {
				total += typeof v === 'number' ? v : 0;
			});
			return total;
		});
	},
});

describe('reference memo for functions', () => {
	it('builds derived data once per pass and rebuilds after the range changes', () => {
		const cells: Record<string, string | number> = { A1: 1, A2: 2, A3: '=A1+A2' };
		for (let i = 1; i <= 50; i++) {
			cells[`C${i}`] = '=TESTMEMOSUM($A$1:$A$3)';
		}
		builds = 0;
		const { wb, engine } = calculated({ Sheet1: cells });
		expect(valueAt(wb, 'C50')).toBe(6);
		// One build during rebuild, one during recalculateAll: never per cell.
		expect(builds).toBe(2);
		builds = 0;
		edit(engine, 'A1', 10);
		expect(valueAt(wb, 'C1')).toBe(24);
		expect(valueAt(wb, 'C50')).toBe(24);
		expect(builds).toBe(1);
		builds = 0;
		engine.evaluateFormula('TESTMEMOSUM(A1:A3)', 0, 9, 9);
		engine.evaluateFormula('TESTMEMOSUM(A1:A3)', 0, 9, 9);
		expect(builds).toBe(2);
	});
});
