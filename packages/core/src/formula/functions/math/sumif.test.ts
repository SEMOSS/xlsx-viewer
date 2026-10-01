import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import type { AstNode } from '../../ast';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { MATH_FUNCTIONS } from './index';

// Microsoft's documented SUMIF / SUMIFS examples.
const ctx = new MockContext({
	A2: 100000,
	A3: 200000,
	A4: 300000,
	A5: 400000,
	B2: 7000,
	B3: 14000,
	B4: 21000,
	B5: 28000,
	C2: 250000,
	// Category / food / sales table.
	E2: 'Vegetables',
	E3: 'Vegetables',
	E4: 'Fruits',
	E6: 'Vegetables',
	E7: 'Fruits',
	F2: 'Tomatoes',
	F3: 'Celery',
	F4: 'Oranges',
	F5: 'Butter',
	F6: 'Carrots',
	F7: 'Apples',
	G2: 2300,
	G3: 5500,
	G4: 800,
	G5: 400,
	G6: 4200,
	G7: 1200,
	// SUMIFS table.
	I2: 5,
	I3: 4,
	I4: 15,
	I5: 3,
	I6: 22,
	I7: 12,
	I8: 10,
	I9: 33,
	J2: 'Apples',
	J3: 'Apples',
	J4: 'Artichokes',
	J5: 'Artichokes',
	J6: 'Bananas',
	J7: 'Bananas',
	J8: 'Carrots',
	J9: 'Carrots',
	K2: 'Tom',
	K3: 'Sarah',
	K4: 'Tom',
	K5: 'Sarah',
	K6: 'Tom',
	K7: 'Sarah',
	K8: 'Tom',
	K9: 'Sarah',
	// Errors and mixed values.
	M1: 1,
	M2: ERRORS.NA,
	M3: 3,
	N1: 'a',
	N2: 'b',
	N3: 'a',
});
const f = (name: string, ...args: FunctionArg[]) => call(fn(MATH_FUNCTIONS, name), args, ctx);
const missing: AstNode = { type: 'missing' };

describe('SUMIF', () => {
	it('sums with numeric criteria', () => {
		expect(f('SUMIF', ctx.ref('A2:A5'), '>160000', ctx.ref('B2:B5'))).toBe(63000);
		expect(f('SUMIF', ctx.ref('A2:A5'), '>160000')).toBe(900000);
		expect(f('SUMIF', ctx.ref('A2:A5'), 300000, ctx.ref('B2:B5'))).toBe(21000);
		expect(f('SUMIF', ctx.ref('A2:A5'), '>250000', ctx.ref('B2:B5'))).toBe(49000);
		expect(f('SUMIF', ctx.ref('A2:A5'), '<=200000', ctx.ref('B2:B5'))).toBe(21000);
	});

	it('sums with text, wildcard and blank criteria', () => {
		expect(f('SUMIF', ctx.ref('E2:E7'), 'Fruits', ctx.ref('G2:G7'))).toBe(2000);
		expect(f('SUMIF', ctx.ref('E2:E7'), 'vegetables', ctx.ref('G2:G7'))).toBe(12000);
		expect(f('SUMIF', ctx.ref('F2:F7'), '*es', ctx.ref('G2:G7'))).toBe(4300);
		expect(f('SUMIF', ctx.ref('E2:E7'), '', ctx.ref('G2:G7'))).toBe(400);
		expect(f('SUMIF', ctx.ref('E2:E7'), '<>Fruits', ctx.ref('G2:G7'))).toBe(12400);
	});

	it('resizes the sum range to the criteria range', () => {
		expect(f('SUMIF', ctx.ref('A2:A5'), '>160000', ctx.ref('B2'))).toBe(63000);
		expect(f('SUMIF', ctx.ref('E2:E7'), 'Fruits', ctx.ref('G2:G3'))).toBe(2000);
	});

	it('treats an explicitly empty sum range as omitted', () => {
		expect(f('SUMIF', ctx.ref('A2:A5'), '>160000', missing)).toBe(900000);
	});

	it('propagates errors from summed cells only', () => {
		expect(f('SUMIF', ctx.ref('N1:N3'), 'a', ctx.ref('M1:M3'))).toBe(4);
		expect(f('SUMIF', ctx.ref('N1:N3'), 'b', ctx.ref('M1:M3'))).toBe(ERRORS.NA);
		expect(f('SUMIF', ctx.ref('M1:M3'), '>0')).toBe(4);
		expect(f('SUMIF', ctx.ref('M1:M3'), '#N/A', ctx.ref('A2:A4'))).toBe(200000);
		expect(f('SUMIF', ctx.ref('M1:M3'), ERRORS.NA, ctx.ref('A2:A4'))).toBe(200000);
	});

	it('lifts over criteria arrays', () => {
		expect(f('SUMIF', ctx.ref('E2:E7'), arr(['Fruits', 'Vegetables']), ctx.ref('G2:G7'))).toEqual(
			arr([2000, 12000]),
		);
	});

	it('accepts arrays and whole columns, rejects scalars', () => {
		expect(f('SUMIF', arr([1, 2, 3]), '>1')).toBe(5);
		expect(f('SUMIF', ctx.ref('A:A'), '>250000', ctx.ref('B:B'))).toBe(49000);
		expect(f('SUMIF', 5, 5)).toBe(ERRORS.VALUE);
		expect(f('SUMIF', ERRORS.REF, 5)).toBe(ERRORS.REF);
	});
});

describe('SUMIFS', () => {
	it('sums with several criteria', () => {
		const range = (a1: string) => ctx.ref(a1);
		expect(f('SUMIFS', range('I2:I9'), range('J2:J9'), '=A*', range('K2:K9'), 'Tom')).toBe(20);
		expect(f('SUMIFS', range('I2:I9'), range('J2:J9'), '<>Bananas', range('K2:K9'), 'Tom')).toBe(
			30,
		);
		expect(f('SUMIFS', range('I2:I9'), range('K2:K9'), 'Sarah')).toBe(52);
		expect(f('SUMIFS', range('I2:I9'), range('I2:I9'), '>=10', range('I2:I9'), '<20')).toBe(37);
	});

	it('requires matching shapes and paired arguments', () => {
		expect(f('SUMIFS', ctx.ref('I2:I9'), ctx.ref('J2:J8'), 'Apples')).toBe(ERRORS.VALUE);
		expect(f('SUMIFS', ctx.ref('I2:I9'), ctx.ref('J2:J9'), 'Apples', ctx.ref('K2:K9'))).toBe(
			ERRORS.VALUE,
		);
		expect(f('SUMIFS', ctx.ref('I2:I9'), ctx.ref('J2:J9'))).toBe(ERRORS.VALUE);
	});

	it('lifts over criteria arrays and works on whole columns', () => {
		expect(f('SUMIFS', ctx.ref('I2:I9'), ctx.ref('K2:K9'), arr(['Tom'], ['Sarah']))).toEqual(
			arr([52], [52]),
		);
		expect(f('SUMIFS', ctx.ref('I:I'), ctx.ref('J:J'), 'Carrots')).toBe(43);
	});

	it('propagates errors from the summed cells', () => {
		expect(f('SUMIFS', ctx.ref('M1:M3'), ctx.ref('N1:N3'), 'b')).toBe(ERRORS.NA);
		expect(f('SUMIFS', ctx.ref('M1:M3'), ctx.ref('N1:N3'), 'a')).toBe(4);
	});
});
