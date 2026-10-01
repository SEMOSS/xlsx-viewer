import { expect } from 'vitest';

import type { FormulaValue, FunctionArg, Scalar } from '../../../types';
import { isArrayValue } from '../../../types';
import { call, fn, MockContext } from '../../../test-utils/mock-context';
import { STATISTICAL_FUNCTIONS } from '..';

/** Call a statistical function by name the way the evaluator does. */
export function stat(name: string, args: FunctionArg[], ctx = new MockContext()): FormulaValue {
	return call(fn(STATISTICAL_FUNCTIONS, name), args, ctx);
}

/** Assert a finite number within a relative tolerance of `expected`. */
export function expectClose(actual: unknown, expected: number, rel = 1e-13): void {
	expect(typeof actual, `expected a number, got ${JSON.stringify(actual)}`).toBe('number');
	const a = actual as number;
	const tolerance = expected === 0 ? rel : rel * Math.abs(expected);
	expect(Math.abs(a - expected), `${a} vs ${expected}`).toBeLessThanOrEqual(tolerance);
}

/** Rows of an array result (fails when the value is not an array). */
export function rowsOf(v: FormulaValue): Scalar[][] {
	expect(isArrayValue(v), `expected an array, got ${JSON.stringify(v)}`).toBe(true);
	return isArrayValue(v) ? v.values : [];
}

/** Vertical array literal. */
export function column(...values: Scalar[]) {
	return { kind: 'array' as const, rows: values.length, cols: 1, values: values.map((v) => [v]) };
}

/** Horizontal array literal. */
export function row(...values: Scalar[]) {
	return { kind: 'array' as const, rows: 1, cols: values.length, values: [values] };
}
