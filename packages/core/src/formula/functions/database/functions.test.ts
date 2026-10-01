import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import type { Scalar } from '../../types';
import { block, CountingContext } from '../lookup/test-utils/helpers';
import {
	DAVERAGE,
	DCOUNT,
	DCOUNTA,
	DGET,
	DMAX,
	DMIN,
	DPRODUCT,
	DSTDEV,
	DSTDEVP,
	DSUM,
	DVAR,
	DVARP,
} from './functions';
import { DATABASE_FUNCTIONS } from './index';

/**
 * The example from Microsoft's D-function documentation: criteria in A1:F3
 * (="=Apple" / ="=Pear" evaluate to the text "=Apple" / "=Pear"), the
 * database in A5:E11.
 */
const DOC_SHEET: Record<string, Scalar> = {
	...block([
		['Tree', 'Height', 'Age', 'Yield', 'Profit', 'Height'],
		['=Apple', '>10', null, null, null, '<16'],
		['=Pear'],
	]),
	...block(
		[
			['Tree', 'Height', 'Age', 'Yield', 'Profit'],
			['Apple', 18, 20, 14, 105],
			['Pear', 12, 12, 10, 96],
			['Cherry', 13, 14, 9, 105],
			['Apple', 14, 15, 10, 75],
			['Pear', 9, 8, 8, 76.8],
			['Apple', 8, 9, 6, 45],
		],
		5,
	),
};

const ctx = new MockContext(DOC_SHEET);
const db = ctx.ref('A5:E11');

describe('D-functions: documented examples', () => {
	it('counts', () => {
		expect(call(DCOUNT, [db, 'Age', ctx.ref('A1:F2')], ctx)).toBe(1);
		expect(call(DCOUNTA, [db, 'Profit', ctx.ref('A1:F2')], ctx)).toBe(1);
	});

	it('finds extremes', () => {
		expect(call(DMAX, [db, 'Profit', ctx.ref('A1:A3')], ctx)).toBe(105);
		expect(call(DMIN, [db, 'Profit', ctx.ref('A1:B2')], ctx)).toBe(75);
	});

	it('sums and multiplies', () => {
		expect(call(DSUM, [db, 'Profit', ctx.ref('A1:A2')], ctx)).toBe(225);
		expect(call(DSUM, [db, 'Profit', ctx.ref('A1:F2')], ctx)).toBe(75);
		expect(call(DPRODUCT, [db, 'Yield', ctx.ref('A1:F2')], ctx)).toBe(10);
	});

	it('averages, including with the database as its own criteria', () => {
		expect(call(DAVERAGE, [db, 'Yield', ctx.ref('A1:B2')], ctx)).toBe(12);
		expect(call(DAVERAGE, [db, 3, db], ctx)).toBe(13);
	});

	it('computes deviations and variances', () => {
		const crit = ctx.ref('A1:A3');
		expect(call(DSTDEV, [db, 'Yield', crit], ctx)).toBeCloseTo(2.96648, 5);
		expect(call(DSTDEVP, [db, 'Yield', crit], ctx)).toBeCloseTo(2.6533, 5);
		expect(call(DVAR, [db, 'Yield', crit], ctx)).toBeCloseTo(8.8, 10);
		expect(call(DVARP, [db, 'Yield', crit], ctx)).toBeCloseTo(7.04, 10);
	});

	it('rejects DGET with several matches', () => {
		expect(call(DGET, [db, 'Yield', ctx.ref('A1:A3')], ctx)).toBe(ERRORS.NUM);
	});
});

describe('D-functions: aggregation edge cases', () => {
	// Criteria cells for these tests live in H1:H4.
	const c = new MockContext({
		...DOC_SHEET,
		H1: 'Tree',
		H2: 'Cherry',
		I1: 'Tree',
		I2: 'Plum',
		J1: 'Tree',
		J2: '=Apple',
		K1: 'Height',
		K2: '>15',
	});
	const one = c.ref('H1:H2');
	const none = c.ref('I1:I2');

	it('returns the field of the single matching record with DGET', () => {
		expect(call(DGET, [db, 'Profit', one], c)).toBe(105);
		expect(call(DGET, [db, 'tree', one], c)).toBe('Cherry');
		expect(call(DGET, [db, 'Profit', none], c)).toBe(ERRORS.VALUE);
		expect(call(DGET, [db, 'Profit', c.ref('J1:K2')], c)).toBe(105);
	});

	it('handles empty selections per function', () => {
		expect(call(DSUM, [db, 'Profit', none], c)).toBe(0);
		expect(call(DCOUNT, [db, 'Profit', none], c)).toBe(0);
		expect(call(DMAX, [db, 'Profit', none], c)).toBe(0);
		expect(call(DMIN, [db, 'Profit', none], c)).toBe(0);
		expect(call(DPRODUCT, [db, 'Profit', none], c)).toBe(0);
		expect(call(DAVERAGE, [db, 'Profit', none], c)).toBe(ERRORS.DIV0);
		expect(call(DVARP, [db, 'Profit', none], c)).toBe(ERRORS.DIV0);
		expect(call(DSTDEVP, [db, 'Profit', none], c)).toBe(ERRORS.DIV0);
	});

	it('needs two values for sample statistics', () => {
		expect(call(DSTDEV, [db, 'Profit', one], c)).toBe(ERRORS.DIV0);
		expect(call(DVAR, [db, 'Profit', one], c)).toBe(ERRORS.DIV0);
		expect(call(DSTDEVP, [db, 'Profit', one], c)).toBe(0);
		expect(call(DVARP, [db, 'Profit', one], c)).toBe(0);
	});

	it('counts whole records when the field is omitted', () => {
		const apples = c.ref('J1:J2');
		expect(call(DCOUNT, [db, null, apples], c)).toBe(3);
		expect(call(DCOUNT, [db, { type: 'missing' }, apples], c)).toBe(3);
		expect(call(DCOUNTA, [db, null, apples], c)).toBe(3);
		expect(call(DSUM, [db, null, apples], c)).toBe(ERRORS.VALUE);
		expect(call(DGET, [db, null, one], c)).toBe(ERRORS.VALUE);
	});

	it('accepts field labels case-insensitively and 1-based numbers', () => {
		const apples = c.ref('J1:J2');
		expect(call(DSUM, [db, 'PROFIT', apples], c)).toBe(225);
		expect(call(DSUM, [db, 5, apples], c)).toBe(225);
		expect(call(DSUM, [db, 5.9, apples], c)).toBe(225);
		expect(call(DSUM, [db, c.ref('E5'), apples], c)).toBe(225);
	});

	it('rejects fields that name no column', () => {
		const apples = c.ref('J1:J2');
		expect(call(DSUM, [db, 'Weight', apples], c)).toBe(ERRORS.VALUE);
		expect(call(DSUM, [db, 0, apples], c)).toBe(ERRORS.VALUE);
		expect(call(DSUM, [db, 6, apples], c)).toBe(ERRORS.VALUE);
		expect(call(DSUM, [db, true, apples], c)).toBe(ERRORS.VALUE);
		expect(call(DSUM, [db, ERRORS.NA, apples], c)).toBe(ERRORS.NA);
	});
});

describe('D-functions: values in the field column', () => {
	const c = new MockContext({
		...block([
			['Name', 'Value'],
			['a', 10],
			['a', 'text'],
			['a', true],
			['a', null],
			['a', 5],
			['b', ERRORS.DIV0],
		]),
		...block(
			[
				['Name', 'Name'],
				['a', null],
				[null, 'b'],
			],
			1,
			4,
		),
		G1: 'Name',
		G2: 'b',
	});
	const data = c.ref('A1:B7');
	const onlyA = c.ref('E1:E2');
	const withB = c.ref('E1:F3');

	it('aggregates only numbers', () => {
		expect(call(DSUM, [data, 'Value', onlyA], c)).toBe(15);
		expect(call(DAVERAGE, [data, 'Value', onlyA], c)).toBe(7.5);
		expect(call(DCOUNT, [data, 'Value', onlyA], c)).toBe(2);
		expect(call(DCOUNTA, [data, 'Value', onlyA], c)).toBe(4);
		expect(call(DMAX, [data, 'Value', onlyA], c)).toBe(10);
		expect(call(DMIN, [data, 'Value', onlyA], c)).toBe(5);
		expect(call(DPRODUCT, [data, 'Value', onlyA], c)).toBe(50);
	});

	it('propagates errors from matching records only', () => {
		expect(call(DSUM, [data, 'Value', withB], c)).toBe(ERRORS.DIV0);
		expect(call(DMAX, [data, 'Value', withB], c)).toBe(ERRORS.DIV0);
		expect(call(DCOUNT, [data, 'Value', withB], c)).toBe(2);
		expect(call(DCOUNTA, [data, 'Value', withB], c)).toBe(5);
		expect(call(DGET, [data, 'Value', c.ref('G1:G2')], c)).toBe(ERRORS.DIV0);
		// F2 is blank, so the criteria row F2 has no condition and matches every record.
		expect(call(DGET, [data, 'Value', c.ref('F1:F3')], c)).toBe(ERRORS.NUM);
	});

	it('returns a blank field from DGET as blank', () => {
		const ctx2 = new MockContext(block([['K', 'V'], ['x', null], ['y', 2], ['K'], ['x']]));
		expect(call(DGET, [ctx2.ref('A1:B3'), 'V', ctx2.ref('A4:A5')], ctx2)).toBeNull();
	});
});

describe('D-functions: inputs', () => {
	it('works with arrays for the database and criteria', () => {
		const data = arr(['k', 'n'], ['x', 1], ['y', 2], ['x', 3]);
		expect(call(DSUM, [data, 'n', arr(['k'], ['x'])])).toBe(4);
		expect(call(DCOUNT, [data, 2, arr(['n'], ['>1'])])).toBe(2);
	});

	it('propagates errors and rejects criteria without a condition row', () => {
		expect(call(DSUM, [ERRORS.REF, 'n', arr(['k'], ['x'])])).toBe(ERRORS.REF);
		expect(call(DSUM, [arr(['k', 'n'], ['x', 1]), 'n', ERRORS.NA])).toBe(ERRORS.NA);
		expect(call(DSUM, [arr(['k', 'n'], ['x', 1]), 'n', arr(['k'])])).toBe(ERRORS.VALUE);
		const union = { kind: 'refs' as const, refs: [ctx.ref('A1'), ctx.ref('C3')] };
		expect(call(DSUM, [union, 'n', arr(['k'], ['x'])], ctx)).toBe(ERRORS.VALUE);
	});

	it('handles whole-column databases lazily on 100k records', () => {
		const cells: Record<string, Scalar> = { A1: 'Region', B1: 'Amount', C1: 'Note' };
		for (let r = 2; r <= 100_001; r++) {
			cells[`A${r}`] = r % 4 === 0 ? 'North' : 'South';
			cells[`B${r}`] = r;
			cells[`C${r}`] = 'x';
		}
		const big = new CountingContext({ ...cells, E1: 'Region', E2: 'North' });
		const started = Date.now();
		const total = call(DSUM, [big.ref('A:C'), 'Amount', big.ref('E1:E2')], big);
		expect(Date.now() - started).toBeLessThan(1500);
		let expected = 0;
		for (let r = 4; r <= 100_001; r += 4) {
			expected += r;
		}
		expect(total).toBe(expected);
		// Labels (3), the criteria cells (2), the region of every record and the
		// amount of the matching quarter: the Note column is never read.
		expect(big.reads).toBeLessThan(100_000 + 25_000 + 20);
	});

	it('registers all twelve functions without file prefixes', () => {
		expect(DATABASE_FUNCTIONS.map((d) => d.name).sort()).toEqual([
			'DAVERAGE',
			'DCOUNT',
			'DCOUNTA',
			'DGET',
			'DMAX',
			'DMIN',
			'DPRODUCT',
			'DSTDEV',
			'DSTDEVP',
			'DSUM',
			'DVAR',
			'DVARP',
		]);
		for (const def of DATABASE_FUNCTIONS) {
			expect(def.filePrefix).toBeUndefined();
			expect(def.params).toEqual(['any', 'value', 'any']);
		}
	});
});
