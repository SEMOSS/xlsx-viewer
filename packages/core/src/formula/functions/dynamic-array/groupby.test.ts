import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call } from '../../test-utils/mock-context';
import type { ArrayValue, FormulaValue, Scalar } from '../../types';
import { isArrayValue } from '../../types';
import { GROUPBY } from './groupby';
import { LambdaContext, name } from './test-utils/lambda-context';

const REGIONS = arr(['East'], ['West'], ['East'], ['west'], ['North'], ['East']);
const PRODUCTS = arr(['Apples'], ['Pears'], ['Pears'], ['Apples'], ['Apples'], ['Apples']);
const SALES = arr([10], [20], [5], [7], [3], [2]);
const BOTH = arr(
	['East', 'Apples'],
	['West', 'Pears'],
	['East', 'Pears'],
	['west', 'Apples'],
	['North', 'Apples'],
	['East', 'Apples'],
);
const MISSING = { type: 'missing' } as const;

function g(args: unknown[], ctx = new LambdaContext()): unknown {
	const out = call(GROUPBY, args as FormulaValue[], ctx);
	return isArrayValue(out) ? out.values : out;
}

function withHeader(head: Scalar, a: ArrayValue): ArrayValue {
	return arr([head, ...new Array<Scalar>(a.cols - 1).fill(null)], ...a.values);
}

describe('GROUPBY basics', () => {
	it('groups, sorts keys ascending and adds a grand total', () => {
		expect(g([REGIONS, SALES, name('SUM')])).toEqual([
			['East', 17],
			['North', 3],
			['West', 27],
			['Total', 47],
		]);
	});

	it('groups text case-insensitively, keeping the first spelling', () => {
		expect(g([arr(['b'], ['B'], ['a']), arr([1], [2], [3]), name('COUNT')])).toEqual([
			['a', 1],
			['b', 2],
			['Total', 3],
		]);
	});

	it('detects and hides headers automatically for one level', () => {
		const regions = withHeader('Region', REGIONS);
		const sales = withHeader('Sales', SALES);
		expect(g([regions, sales, name('SUM')])).toEqual([
			['East', 17],
			['North', 3],
			['West', 27],
			['Total', 47],
		]);
		expect(g([regions, sales, name('SUM'), 3])).toEqual([
			['Region', 'Sales'],
			['East', 17],
			['North', 3],
			['West', 27],
			['Total', 47],
		]);
		expect(g([regions, sales, name('MAX'), 1, 0])).toEqual([
			['East', 10],
			['North', 3],
			['West', 20],
		]);
	});

	it('generates headers with field_headers 2', () => {
		expect(g([REGIONS, SALES, name('SUM'), 2, 0])).toEqual([
			['Row Field 1', 'Value 1'],
			['East', 17],
			['North', 3],
			['West', 27],
		]);
	});

	it('places totals by total_depth', () => {
		expect(g([REGIONS, SALES, name('SUM'), MISSING, 0])).toEqual([
			['East', 17],
			['North', 3],
			['West', 27],
		]);
		expect(g([REGIONS, SALES, name('SUM'), MISSING, -1])).toEqual([
			['Total', 47],
			['East', 17],
			['North', 3],
			['West', 27],
		]);
	});

	it('sorts by keys or aggregates with sort_order', () => {
		const firstColumn = (v: unknown) => (v as Scalar[][]).map((r) => r[0]);
		expect(firstColumn(g([REGIONS, SALES, name('SUM'), MISSING, 0, -1]))).toEqual([
			'West',
			'North',
			'East',
		]);
		expect(firstColumn(g([REGIONS, SALES, name('SUM'), MISSING, 0, 2]))).toEqual([
			'North',
			'East',
			'West',
		]);
		expect(firstColumn(g([REGIONS, SALES, name('SUM'), MISSING, 0, -2]))).toEqual([
			'West',
			'East',
			'North',
		]);
	});

	it('aggregates several value columns', () => {
		const values = arr([10, 1], [20, 2], [5, 3], [7, 4], [3, 5], [2, 6]);
		expect(g([REGIONS, values, name('SUM')])).toEqual([
			['East', 17, 10],
			['North', 3, 5],
			['West', 27, 6],
			['Total', 47, 21],
		]);
	});

	it('filters rows and sorts blank keys last', () => {
		const keep = arr([false], [true], [false], [true], [true], [false]);
		expect(g([REGIONS, SALES, name('SUM'), MISSING, MISSING, MISSING, keep])).toEqual([
			['North', 3],
			['West', 27],
			['Total', 30],
		]);
		expect(g([arr([2], [null], [1]), arr([1], [2], [3]), name('SUM')])).toEqual([
			[1, 3],
			[2, 1],
			[null, 2],
			['Total', 6],
		]);
	});
});

describe('GROUPBY hierarchies', () => {
	it('adds first-level subtotals by default', () => {
		expect(g([BOTH, SALES, name('SUM')])).toEqual([
			['East', 'Apples', 12],
			['East', 'Pears', 5],
			['East', null, 17],
			['North', 'Apples', 3],
			['North', null, 3],
			['West', 'Apples', 7],
			['West', 'Pears', 20],
			['West', null, 27],
			['Total', null, 47],
		]);
	});

	it('puts subtotals and the total first for negative depths', () => {
		expect(g([BOTH, SALES, name('SUM'), MISSING, -2])).toEqual([
			['Total', null, 47],
			['East', null, 17],
			['East', 'Apples', 12],
			['East', 'Pears', 5],
			['North', null, 3],
			['North', 'Apples', 3],
			['West', null, 27],
			['West', 'Apples', 7],
			['West', 'Pears', 20],
		]);
		expect(g([BOTH, SALES, name('SUM'), MISSING, 1])).toHaveLength(6);
	});

	it('sorts a level by its field number within the hierarchy', () => {
		const out = g([BOTH, SALES, name('SUM'), MISSING, 0, -2]) as Scalar[][];
		expect(out.map((r) => `${r[0]}/${r[1]}`)).toEqual([
			'East/Pears',
			'East/Apples',
			'North/Apples',
			'West/Pears',
			'West/Apples',
		]);
	});

	it('sorts fields independently with field_relationship 1', () => {
		expect(g([BOTH, SALES, name('SUM'), MISSING, MISSING, -2, MISSING, 1])).toEqual([
			['East', 'Pears', 5],
			['West', 'Pears', 20],
			['East', 'Apples', 12],
			['North', 'Apples', 3],
			['West', 'Apples', 7],
			['Total', null, 47],
		]);
	});

	it('shows headers automatically for several levels', () => {
		const fields = arr(['Region', 'Product'], ...BOTH.values);
		const out = g([fields, withHeader('Sales', SALES), name('SUM'), MISSING, 0]) as Scalar[][];
		expect(out[0]).toEqual(['Region', 'Product', 'Sales']);
		expect(out).toHaveLength(6);
	});
});

describe('GROUPBY functions', () => {
	const col = (fn: string) => (g([REGIONS, SALES, name(fn)]) as Scalar[][]).map((r) => r[1]);

	it('supports the common eta-reduced aggregations', () => {
		expect(col('AVERAGE')).toEqual([17 / 3, 3, 13.5, 47 / 6]);
		expect(col('COUNT')).toEqual([3, 1, 2, 6]);
		expect(col('COUNTA')).toEqual([3, 1, 2, 6]);
		expect(col('MAX')).toEqual([10, 3, 20, 20]);
		expect(col('MIN')).toEqual([2, 3, 7, 2]);
		expect(col('PRODUCT')).toEqual([100, 3, 140, 42000]);
		expect(col('MEDIAN')).toEqual([5, 3, 13.5, 6]);
		expect(col('PERCENTOF')).toEqual([17 / 47, 3 / 47, 27 / 47, 1]);
		expect(col('var.p')[1]).toBe(0);
		expect(col('STDEV.S')[1]).toBe(ERRORS.DIV0);
		expect(col('STDEV')[2]).toBeCloseTo(9.19239, 5);
	});

	it('keeps errors in the affected groups only', () => {
		const sales = arr([10], [ERRORS.NA], [5], [7], [3], [2]);
		expect(g([REGIONS, sales, name('SUM')])).toEqual([
			['East', 17],
			['North', 3],
			['West', ERRORS.NA],
			['Total', ERRORS.NA],
		]);
		expect((g([REGIONS, sales, name('COUNTA')]) as Scalar[][])[2]).toEqual(['West', 2]);
	});

	it('calls LAMBDAs with the group (and the whole column for two parameters)', () => {
		const ctx = new LambdaContext();
		const count = ctx.lambda(['x'], ([x]) => (x as ArrayValue).rows);
		const share = ctx.lambda(
			['part', 'all'],
			([p, a]) => (p as ArrayValue).rows / (a as ArrayValue).rows,
		);
		ctx.bindings.set('COUNTER', count);
		ctx.bindings.set('SHARE', share);
		expect(g([REGIONS, SALES, name('Counter')], ctx)).toEqual([
			['East', 3],
			['North', 1],
			['West', 2],
			['Total', 6],
		]);
		expect((g([REGIONS, SALES, name('share')], ctx) as Scalar[][])[0]).toEqual(['East', 0.5]);
		const pair = ctx.lambda(['x'], () => arr([1, 2]));
		ctx.bindings.set('PAIR', pair);
		expect((g([REGIONS, SALES, name('PAIR')], ctx) as Scalar[][])[0]).toEqual([
			'East',
			ERRORS.CALC,
		]);
	});

	it('lets a LET / LAMBDA binding shadow a built-in name', () => {
		const ctx = new LambdaContext();
		const fake = ctx.lambda(['x'], () => 'shadowed');
		ctx.bindings.set('SUM', fake);
		ctx.scope = new Map([['SUM', fake]]);
		expect((g([REGIONS, SALES, name('SUM')], ctx) as Scalar[][])[0]).toEqual(['East', 'shadowed']);
	});

	it('rejects names that are neither built-in nor LAMBDAs', () => {
		const ctx = new LambdaContext();
		expect(g([REGIONS, SALES, name('NOPE')], ctx)).toBe(ERRORS.NAME);
		ctx.bindings.set('FIVE', 5);
		expect(g([REGIONS, SALES, name('FIVE')], ctx)).toBe(ERRORS.VALUE);
		expect(g([REGIONS, SALES, { type: 'number', value: 1 }], ctx)).toBe(ERRORS.VALUE);
	});
});

describe('GROUPBY validation', () => {
	it('rejects mismatched or invalid arguments', () => {
		expect(g([REGIONS, arr([1], [2]), name('SUM')])).toBe(ERRORS.VALUE);
		expect(g([REGIONS, SALES, name('SUM'), 5])).toBe(ERRORS.VALUE);
		expect(g([REGIONS, SALES, name('SUM'), MISSING, 'x'])).toBe(ERRORS.VALUE);
		expect(g([REGIONS, SALES, name('SUM'), MISSING, MISSING, 0])).toBe(ERRORS.VALUE);
		expect(g([REGIONS, SALES, name('SUM'), MISSING, MISSING, 3])).toBe(ERRORS.VALUE);
		expect(g([BOTH, SALES, name('SUM'), MISSING, MISSING, arr([1, 3])])).toBe(ERRORS.VALUE);
		expect(g([REGIONS, SALES, name('SUM'), MISSING, MISSING, MISSING, arr([true])])).toBe(
			ERRORS.VALUE,
		);
		expect(g([REGIONS, SALES, name('SUM'), MISSING, MISSING, MISSING, MISSING, 2])).toBe(
			ERRORS.VALUE,
		);
		expect(g([ERRORS.REF, SALES, name('SUM')])).toBe(ERRORS.VALUE);
		expect(g([ERRORS.REF, arr([1]), name('SUM')])).toEqual([
			[ERRORS.REF, 1],
			['Total', 1],
		]);
	});

	it('is #CALC! when no rows remain', () => {
		const none = arr([false], [false], [false], [false], [false], [false]);
		expect(g([REGIONS, SALES, name('SUM'), MISSING, MISSING, MISSING, none])).toBe(ERRORS.CALC);
	});

	it('groups 100k rows quickly', () => {
		const n = 100_000;
		const keys = arr(...Array.from({ length: n }, (_, i) => [`k${i % 100}`]));
		const vals = arr(...Array.from({ length: n }, (_, i) => [i]));
		const started = Date.now();
		const out = g([keys, vals, name('SUM')]) as Scalar[][];
		expect(Date.now() - started).toBeLessThan(1500);
		expect(out).toHaveLength(101);
		expect(out[100]).toEqual(['Total', (n * (n - 1)) / 2]);
	});
});
