import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call } from '../../test-utils/mock-context';
import type { FormulaValue, Scalar } from '../../types';
import { isArrayValue } from '../../types';
import { PIVOTBY } from './pivotby';
import { LambdaContext, name } from './test-utils/lambda-context';

const REGIONS = arr(['East'], ['West'], ['East'], ['west'], ['North'], ['East']);
const PRODUCTS = arr(['Apples'], ['Pears'], ['Pears'], ['Apples'], ['Apples'], ['Apples']);
const SALES = arr([10], [20], [5], [7], [3], [2]);
const MISSING = { type: 'missing' } as const;

function p(args: unknown[], ctx = new LambdaContext()): unknown {
	const out = call(PIVOTBY, args as FormulaValue[], ctx);
	return isArrayValue(out) ? out.values : out;
}

describe('PIVOTBY', () => {
	it('pivots rows by columns with totals on both axes', () => {
		expect(p([REGIONS, PRODUCTS, SALES, name('SUM')])).toEqual([
			[null, 'Apples', 'Pears', 'Total'],
			['East', 12, 5, 17],
			['North', 3, null, 3],
			['West', 7, 20, 27],
			['Total', 22, 25, 47],
		]);
	});

	it('controls row and column totals separately', () => {
		expect(p([REGIONS, PRODUCTS, SALES, name('SUM'), MISSING, 0, MISSING, 0])).toEqual([
			[null, 'Apples', 'Pears'],
			['East', 12, 5],
			['North', 3, null],
			['West', 7, 20],
		]);
		expect(p([REGIONS, PRODUCTS, SALES, name('SUM'), MISSING, -1, MISSING, -1])).toEqual([
			[null, 'Total', 'Apples', 'Pears'],
			['Total', 47, 22, 25],
			['East', 17, 12, 5],
			['North', 3, 3, null],
			['West', 27, 7, 20],
		]);
	});

	it('sorts rows and columns', () => {
		const out = p([REGIONS, PRODUCTS, SALES, name('SUM'), MISSING, 0, -1, 0, -2]) as Scalar[][];
		expect(out[0]).toEqual([null, 'Pears', 'Apples']);
		expect(out.slice(1).map((r) => r[0])).toEqual(['West', 'North', 'East']);
	});

	it('measures PERCENTOF against column, row or grand totals', () => {
		const col = p([REGIONS, PRODUCTS, SALES, name('PERCENTOF')]) as Scalar[][];
		expect(col[1]).toEqual(['East', 12 / 22, 5 / 25, 17 / 47]);
		expect(col[4]).toEqual(['Total', 1, 1, 1]);
		const row = p([REGIONS, PRODUCTS, SALES, name('PERCENTOF'), ...Array(6).fill(MISSING), 1]);
		expect((row as Scalar[][])[1]).toEqual(['East', 12 / 17, 5 / 17, 1]);
		const grand = p([REGIONS, PRODUCTS, SALES, name('PERCENTOF'), ...Array(6).fill(MISSING), 2]);
		expect((grand as Scalar[][])[1]).toEqual(['East', 12 / 47, 5 / 47, 17 / 47]);
		expect(p([REGIONS, PRODUCTS, SALES, name('PERCENTOF'), ...Array(6).fill(MISSING), 3])).toBe(
			ERRORS.VALUE,
		);
	});

	it('adds a value-name row for several value columns', () => {
		const values = arr(['Sales', 'Qty'], [10, 1], [20, 2], [5, 3], [7, 4], [3, 5], [2, 6]);
		const regions = arr(['Region'], ...REGIONS.values);
		const products = arr(['Product'], ...PRODUCTS.values);
		const out = p([regions, products, values, name('SUM'), 3, 0, MISSING, 0]) as Scalar[][];
		expect(out[0]).toEqual([null, 'Apples', 'Apples', 'Pears', 'Pears']);
		expect(out[1]).toEqual(['Region', 'Sales', 'Qty', 'Sales', 'Qty']);
		expect(out[2]).toEqual(['East', 12, 7, 5, 3]);
	});

	it('nests several row and column levels', () => {
		const rows = arr(['A', 'x'], ['A', 'y'], ['B', 'x'], ['A', 'x']);
		const cols = arr([2023, 'Q1'], [2024, 'Q1'], [2023, 'Q2'], [2024, 'Q1']);
		const out = p([
			rows,
			cols,
			arr([1], [2], [4], [8]),
			name('SUM'),
			0,
			2,
			MISSING,
			1,
		]) as Scalar[][];
		expect(out[0]).toEqual([null, null, 2023, 2023, 2024, 'Total']);
		expect(out[1]).toEqual([null, null, 'Q1', 'Q2', 'Q1', null]);
		expect(out.slice(2)).toEqual([
			['A', 'x', 1, null, 8, 9],
			['A', 'y', null, null, 2, 2],
			['A', null, 1, null, 10, 11],
			['B', 'x', null, 4, null, 4],
			['B', null, null, 4, null, 4],
			['Total', null, 1, 4, 10, 15],
		]);
	});

	it('filters rows and calls LAMBDAs per cell', () => {
		const keep = arr([true], [true], [false], [true], [true], [true]);
		const out = p([REGIONS, PRODUCTS, SALES, name('COUNT'), ...Array(5).fill(MISSING), keep]);
		expect((out as Scalar[][])[1]).toEqual(['East', 2, null, 2]);
		const ctx = new LambdaContext();
		ctx.bindings.set(
			'ROWS2',
			ctx.lambda(['v'], ([v]) => (isArrayValue(v) ? v.rows * 2 : 0)),
		);
		const doubled = p([REGIONS, PRODUCTS, SALES, name('ROWS2')], ctx) as Scalar[][];
		expect(doubled[4]).toEqual(['Total', 8, 4, 12]);
	});

	it('validates its arguments', () => {
		expect(p([REGIONS, arr(['x']), SALES, name('SUM')])).toBe(ERRORS.VALUE);
		expect(p([REGIONS, PRODUCTS, SALES, name('UNKNOWN')])).toBe(ERRORS.NAME);
		expect(p([REGIONS, PRODUCTS, SALES, name('SUM'), 9])).toBe(ERRORS.VALUE);
		expect(p([REGIONS, PRODUCTS, SALES, name('SUM'), MISSING, MISSING, 4])).toBe(ERRORS.VALUE);
		expect(p([REGIONS, PRODUCTS, SALES, name('SUM'), MISSING, MISSING, MISSING, 'x'])).toBe(
			ERRORS.VALUE,
		);
		const none = arr([false], [false], [false], [false], [false], [false]);
		expect(p([REGIONS, PRODUCTS, SALES, name('SUM'), ...Array(5).fill(MISSING), none])).toBe(
			ERRORS.CALC,
		);
		expect(p([REGIONS, PRODUCTS, SALES, name('SUM'), ...Array(5).fill(MISSING), arr([true])])).toBe(
			ERRORS.VALUE,
		);
	});
});
