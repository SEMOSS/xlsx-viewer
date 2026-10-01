import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { rowsOf, stat } from '../test-utils/expect';

// Excel's COUNTIF sample: A2:A5 fruit, B2:B5 quantities.
const fruit = () =>
	new MockContext({
		A2: 'apples',
		A3: 'oranges',
		A4: 'peaches',
		A5: 'apples',
		B2: 32,
		B3: 54,
		B4: 75,
		B5: 86,
	});

describe('COUNTIF', () => {
	it('matches the Excel examples', () => {
		const ctx = fruit();
		expect(stat('COUNTIF', [ctx.ref('A2:A5'), 'apples'], ctx)).toBe(2);
		expect(stat('COUNTIF', [ctx.ref('A2:A5'), ctx.ref('A4')], ctx)).toBe(1);
		expect(stat('COUNTIF', [ctx.ref('B2:B5'), '>55'], ctx)).toBe(2);
		expect(stat('COUNTIF', [ctx.ref('B2:B5'), '<>75'], ctx)).toBe(3);
		expect(stat('COUNTIF', [ctx.ref('A2:A5'), '*'], ctx)).toBe(4);
		expect(stat('COUNTIF', [ctx.ref('A2:A5'), '?????es'], ctx)).toBe(2);
		expect(stat('COUNTIF', [ctx.ref('B2:B5'), 32], ctx)).toBe(1);
	});

	it('counts blank cells across the whole range for blank criteria', () => {
		const ctx = fruit();
		ctx.setCells({ A6: '' });
		expect(stat('COUNTIF', [ctx.ref('A2:A10'), ''], ctx)).toBe(5);
		expect(stat('COUNTIF', [ctx.ref('A2:A10'), '='], ctx)).toBe(4);
		expect(stat('COUNTIF', [ctx.ref('A2:A10'), '<>'], ctx)).toBe(4);
		expect(stat('COUNTIF', [ctx.ref('C:C'), ''], ctx)).toBe(1048576);
	});

	it('matches error criteria instead of propagating them', () => {
		const ctx = new MockContext({ A1: ERRORS.NA, A2: 1, A3: ERRORS.NA });
		expect(stat('COUNTIF', [ctx.ref('A1:A3'), ERRORS.NA], ctx)).toBe(2);
		expect(stat('COUNTIF', [ctx.ref('A1:A3'), '#N/A'], ctx)).toBe(2);
	});

	it('is lifted over an array of criteria and accepts arrays as ranges', () => {
		const ctx = fruit();
		expect(rowsOf(stat('COUNTIF', [ctx.ref('A2:A5'), arr(['apples', 'peaches'])], ctx))).toEqual([
			[2, 1],
		]);
		expect(stat('COUNTIF', [arr([1, 2, 3, 4]), '>2'])).toBe(2);
	});
});

describe('COUNTIFS', () => {
	const sales = () =>
		new MockContext({
			A2: 'Davidoski',
			A3: 'Burke',
			A4: 'Sundaram',
			A5: 'Levitan',
			B2: 'Yes',
			B3: 'Yes',
			B4: 'Yes',
			B5: 'No',
			C2: 'No',
			C3: 'Yes',
			C4: 'Yes',
			C5: 'Yes',
			D2: 'No',
			D3: 'No',
			D4: 'Yes',
			D5: 'Yes',
		});

	it('matches the Excel examples', () => {
		const ctx = sales();
		expect(stat('COUNTIFS', [ctx.ref('B2:D2'), '=Yes'], ctx)).toBe(1);
		expect(stat('COUNTIFS', [ctx.ref('B2:B5'), '=Yes', ctx.ref('C2:C5'), '=Yes'], ctx)).toBe(2);
		expect(stat('COUNTIFS', [ctx.ref('B5:D5'), '=Yes', ctx.ref('B3:D3'), '=Yes'], ctx)).toBe(1);
	});

	it('counts positions blank in every range', () => {
		const ctx = new MockContext({ A1: 1, B1: 1 });
		expect(stat('COUNTIFS', [ctx.ref('A1:A5'), '', ctx.ref('B1:B5'), ''], ctx)).toBe(4);
		expect(stat('COUNTIFS', [ctx.ref('A:A'), '', ctx.ref('B:B'), ''], ctx)).toBe(1048575);
	});

	it('requires ranges of one shape and complete pairs', () => {
		const ctx = sales();
		expect(stat('COUNTIFS', [ctx.ref('B2:B5'), 'Yes', ctx.ref('C2:C4'), 'Yes'], ctx)).toBe(
			ERRORS.VALUE,
		);
		expect(stat('COUNTIFS', [ctx.ref('B2:B5'), 'Yes', ctx.ref('C2:C5')], ctx)).toBe(ERRORS.VALUE);
	});
});

// Excel's AVERAGEIF sample: property values and commissions.
const property = () =>
	new MockContext({
		A2: 100000,
		A3: 200000,
		A4: 300000,
		A5: 400000,
		B2: 7000,
		B3: 14000,
		B4: 21000,
		B5: 28000,
	});

describe('AVERAGEIF', () => {
	it('matches the Excel examples', () => {
		const ctx = property();
		expect(stat('AVERAGEIF', [ctx.ref('B2:B5'), '<23000'], ctx)).toBe(14000);
		expect(stat('AVERAGEIF', [ctx.ref('A2:A5'), '<250000'], ctx)).toBe(150000);
		expect(stat('AVERAGEIF', [ctx.ref('A2:A5'), '<95000'], ctx)).toBe(ERRORS.DIV0);
		expect(stat('AVERAGEIF', [ctx.ref('A2:A5'), '>250000', ctx.ref('B2:B5')], ctx)).toBe(24500);
	});

	it('resizes the average range from its top-left cell', () => {
		const ctx = property();
		expect(stat('AVERAGEIF', [ctx.ref('A2:A5'), '>250000', ctx.ref('B2')], ctx)).toBe(24500);
	});

	it('ignores non-numeric averaged cells but propagates their errors', () => {
		const ctx = new MockContext({ A1: 'x', A2: 'x', A3: 'y', B1: 'text', B2: 4, B3: ERRORS.NA });
		expect(stat('AVERAGEIF', [ctx.ref('A1:A3'), 'x', ctx.ref('B1:B3')], ctx)).toBe(4);
		expect(stat('AVERAGEIF', [ctx.ref('A1:A3'), 'y', ctx.ref('B1:B3')], ctx)).toBe(ERRORS.NA);
	});
});

describe('AVERAGEIFS / MAXIFS / MINIFS', () => {
	it('matches the Excel AVERAGEIFS example', () => {
		const ctx = new MockContext({ B2: 75, B3: 94, B4: 86, B5: 'Incomplete' });
		const grades = ctx.ref('B2:B5');
		expect(stat('AVERAGEIFS', [grades, grades, '>70', grades, '<90'], ctx)).toBe(80.5);
		expect(stat('AVERAGEIFS', [grades, grades, '>95'], ctx)).toBe(ERRORS.DIV0);
	});

	it('matches the Excel MAXIFS / MINIFS examples', () => {
		const ctx = new MockContext({
			A2: 89,
			A3: 93,
			A4: 96,
			A5: 85,
			A6: 91,
			A7: 88,
			B2: 1,
			B3: 2,
			B4: 2,
			B5: 3,
			B6: 1,
			B7: 1,
		});
		expect(stat('MAXIFS', [ctx.ref('A2:A7'), ctx.ref('B2:B7'), 1], ctx)).toBe(91);
		expect(stat('MINIFS', [ctx.ref('A2:A7'), ctx.ref('B2:B7'), 1], ctx)).toBe(88);
		expect(stat('MAXIFS', [ctx.ref('A2:A7'), ctx.ref('B2:B7'), 5], ctx)).toBe(0);
		expect(
			stat('MINIFS', [ctx.ref('A2:A7'), ctx.ref('B2:B7'), '>1', ctx.ref('A2:A7'), '<95'], ctx),
		).toBe(85);
	});

	it('requires every range to match the value range shape', () => {
		const ctx = property();
		expect(stat('AVERAGEIFS', [ctx.ref('B2:B5'), ctx.ref('A2:A4'), '>0'], ctx)).toBe(ERRORS.VALUE);
		expect(stat('MAXIFS', [ctx.ref('B2:B5'), ctx.ref('A2:B5'), '>0'], ctx)).toBe(ERRORS.VALUE);
		expect(stat('MINIFS', [ctx.ref('B2:B5'), ctx.ref('A2:A5')], ctx)).toBe(ERRORS.VALUE);
	});

	it('works over whole columns', () => {
		const ctx = property();
		expect(stat('AVERAGEIFS', [ctx.ref('B:B'), ctx.ref('A:A'), '>=300000'], ctx)).toBe(24500);
		expect(stat('MAXIFS', [ctx.ref('B:B'), ctx.ref('A:A'), '<300000'], ctx)).toBe(14000);
	});
});
