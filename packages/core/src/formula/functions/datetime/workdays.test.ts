import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { DATETIME_FUNCTIONS } from './index';

const ctx = new MockContext({ A1: 41235, A2: 41247, A3: 41295, B1: 41235, B3: 'not a date' });
const f = (name: string, ...args: FunctionArg[]) => call(fn(DATETIME_FUNCTIONS, name), args, ctx);
const date = (y: number, m: number, d: number) => f('DATE', y, m, d) as number;

describe('NETWORKDAYS', () => {
	const start = date(2012, 10, 1);
	const end = date(2013, 3, 1);

	it('matches the documented examples', () => {
		expect(f('NETWORKDAYS', start, end)).toBe(110);
		expect(f('NETWORKDAYS', start, end, 41235)).toBe(109);
		expect(f('NETWORKDAYS', start, end, ctx.ref('A1:A3'))).toBe(107);
		expect(f('NETWORKDAYS', start, end, arr([41235, 41247, 41295]))).toBe(107);
	});

	it('counts backwards with a negative sign', () => {
		expect(f('NETWORKDAYS', end, start)).toBe(-110);
		expect(f('NETWORKDAYS', end, start, ctx.ref('A1:A3'))).toBe(-107);
	});

	it('handles single days, weekends and duplicate holidays', () => {
		const monday = date(2024, 1, 8);
		const saturday = date(2024, 1, 6);
		expect(f('NETWORKDAYS', monday, monday)).toBe(1);
		expect(f('NETWORKDAYS', saturday, saturday)).toBe(0);
		expect(f('NETWORKDAYS', saturday, monday)).toBe(1);
		expect(f('NETWORKDAYS', monday, monday + 6, arr([monday, monday, saturday]))).toBe(4);
		expect(f('NETWORKDAYS', start, end, ctx.ref('B1:B2'))).toBe(109);
		expect(f('NETWORKDAYS', start + 0.9, end + 0.2)).toBe(110);
	});

	it('validates holidays and dates', () => {
		expect(f('NETWORKDAYS', start, end, ctx.ref('B1:B3'))).toBe(ERRORS.VALUE);
		expect(f('NETWORKDAYS', start, end, arr([ERRORS.NA]))).toBe(ERRORS.NA);
		expect(f('NETWORKDAYS', start, end, arr([true]))).toBe(ERRORS.VALUE);
		expect(f('NETWORKDAYS', start, end, arr([-5]))).toBe(ERRORS.NUM);
		expect(f('NETWORKDAYS', start, end, '11/22/2012')).toBe(109);
		expect(f('NETWORKDAYS', -1, end)).toBe(ERRORS.NUM);
		expect(f('NETWORKDAYS', 'x', end)).toBe(ERRORS.VALUE);
	});
});

describe('NETWORKDAYS.INTL', () => {
	const jan1 = date(2006, 1, 1);
	const holidays = arr(['2006/1/2', '2006/1/16']);

	it('matches the documented examples', () => {
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 1, 31))).toBe(22);
		expect(f('NETWORKDAYS.INTL', date(2006, 2, 28), date(2006, 1, 31))).toBe(-21);
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 2, 1), 7, holidays)).toBe(22);
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 2, 1), '0010001', holidays)).toBe(20);
	});

	it.each([
		[1, 22],
		[2, 21],
		[3, 21],
		[4, 22],
		[5, 23],
		[6, 23],
		[7, 23],
		[11, 26],
		[12, 26],
		[13, 26],
		[14, 27],
		[15, 27],
		[16, 27],
		[17, 27],
	])('weekend code %d over January 2006 gives %d working days', (code, days) => {
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 1, 31), code)).toBe(days);
	});

	it('accepts weekend masks', () => {
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 1, 31), '0000011')).toBe(22);
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 1, 31), '0000000')).toBe(31);
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 1, 31), '1111111')).toBe(0);
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 1, 31), '1000000')).toBe(26);
		expect(f('NETWORKDAYS.INTL', jan1, date(2006, 1, 31), null, holidays)).toBe(20);
	});

	it('validates the weekend argument', () => {
		const end = date(2006, 1, 31);
		expect(f('NETWORKDAYS.INTL', jan1, end, 0)).toBe(ERRORS.NUM);
		expect(f('NETWORKDAYS.INTL', jan1, end, 8)).toBe(ERRORS.NUM);
		expect(f('NETWORKDAYS.INTL', jan1, end, 18)).toBe(ERRORS.NUM);
		expect(f('NETWORKDAYS.INTL', jan1, end, '000001')).toBe(ERRORS.VALUE);
		expect(f('NETWORKDAYS.INTL', jan1, end, '0000012')).toBe(ERRORS.VALUE);
		expect(f('NETWORKDAYS.INTL', jan1, end, '1')).toBe(ERRORS.VALUE);
		expect(f('NETWORKDAYS.INTL', jan1, end, true)).toBe(ERRORS.VALUE);
		expect(f('NETWORKDAYS.INTL', jan1, end, ERRORS.REF)).toBe(ERRORS.REF);
	});
});

describe('WORKDAY', () => {
	const start = date(2008, 10, 1);
	const friday = date(2024, 1, 5);
	const saturday = date(2024, 1, 6);
	const monday = date(2024, 1, 8);

	it('matches the documented examples', () => {
		expect(f('WORKDAY', start, 151)).toBe(date(2009, 4, 30));
		expect(f('WORKDAY', start, 151, arr([39778], [39786], [39834]))).toBe(date(2009, 5, 5));
	});

	it('steps over weekends and holidays in both directions', () => {
		expect(f('WORKDAY', friday, 1)).toBe(monday);
		expect(f('WORKDAY', monday, -1)).toBe(friday);
		expect(f('WORKDAY', saturday, 0)).toBe(saturday);
		expect(f('WORKDAY', saturday, 1)).toBe(monday);
		expect(f('WORKDAY', saturday, -1)).toBe(friday);
		expect(f('WORKDAY', friday, 1, monday)).toBe(monday + 1);
		expect(f('WORKDAY', friday, 1, arr([monday, monday + 1]))).toBe(monday + 2);
		expect(f('WORKDAY', monday + 2, -2, arr([monday, monday + 1]))).toBe(friday - 1);
		expect(f('WORKDAY', friday, 5)).toBe(friday + 7);
		expect(f('WORKDAY', friday, 10.9)).toBe(friday + 14);
		expect(f('WORKDAY', friday, 1, friday)).toBe(monday);
	});

	it('agrees with NETWORKDAYS over long spans', () => {
		const holidays = arr([45306, 45341, 45439, 45477, 45544, 45603, 45622, 45651]);
		for (const n of [1, 7, 50, 251, 1000, 10000]) {
			const end = f('WORKDAY', 45292, n, holidays) as number;
			expect(f('NETWORKDAYS', 45293, end, holidays)).toBe(n);
			expect(f('NETWORKDAYS', end, end, holidays)).toBe(1);
			const back = f('WORKDAY', end, -n, holidays) as number;
			expect(f('NETWORKDAYS', back, end - 1, holidays)).toBe(n);
		}
	});

	it('rejects results out of range', () => {
		expect(f('WORKDAY', date(9999, 12, 30), 5)).toBe(ERRORS.NUM);
		expect(f('WORKDAY', 3, -10)).toBe(ERRORS.NUM);
		expect(f('WORKDAY', 45292, 1e12)).toBe(ERRORS.NUM);
		expect(f('WORKDAY', 'x', 1)).toBe(ERRORS.VALUE);
	});
});

describe('WORKDAY.INTL', () => {
	const jan1 = date(2012, 1, 1);

	it('matches the documented examples', () => {
		expect(f('WORKDAY.INTL', jan1, 30, 0)).toBe(ERRORS.NUM);
		expect(f('WORKDAY.INTL', jan1, 90, 11)).toBe(41013);
		expect(f('WORKDAY.INTL', jan1, 30, 17)).toBe(40944);
	});

	it('supports masks and holidays', () => {
		expect(f('WORKDAY.INTL', jan1, 1, '0000000')).toBe(jan1 + 1);
		expect(f('WORKDAY.INTL', jan1, 1, '0111111')).toBe(jan1 + 1);
		expect(f('WORKDAY.INTL', jan1, 2, '0111111')).toBe(jan1 + 8);
		expect(f('WORKDAY.INTL', jan1, 1, '0000000', jan1 + 1)).toBe(jan1 + 2);
		expect(f('WORKDAY.INTL', jan1, 5, '1111111')).toBe(ERRORS.VALUE);
		expect(f('WORKDAY.INTL', jan1, 1, null, jan1 + 1)).toBe(jan1 + 2);
	});

	it('carries the _xlfn. file prefix', () => {
		expect(fn(DATETIME_FUNCTIONS, 'WORKDAY.INTL').filePrefix).toBe('_xlfn.');
		expect(fn(DATETIME_FUNCTIONS, 'NETWORKDAYS.INTL').filePrefix).toBe('_xlfn.');
		expect(fn(DATETIME_FUNCTIONS, 'WORKDAY').filePrefix).toBeUndefined();
	});
});
