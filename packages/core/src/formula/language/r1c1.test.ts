import { describe, expect, it } from 'vitest';

import { a1ToR1C1, r1c1ToA1 } from './r1c1';

describe('a1ToR1C1', () => {
	it.each([
		['=A1', 0, 0, '=RC'],
		['=A1', 1, 1, '=R[-1]C[-1]'],
		['=$A$1', 5, 5, '=R1C1'],
		['=A$1', 1, 1, '=R1C[-1]'],
		['=$A1', 1, 1, '=R[-1]C1'],
		['=C5', 0, 0, '=R[4]C[2]'],
		['=SUM(A1:B2)', 2, 2, '=SUM(R[-2]C[-2]:R[-1]C[-1])'],
		['=SUM(1:3)', 0, 0, '=SUM(R:R[2])'],
		['=SUM(5:5)', 0, 0, '=SUM(R[4])'],
		['=SUM($5:$5)', 0, 0, '=SUM(R5)'],
		['=SUM(A:C)', 0, 0, '=SUM(C:C[2])'],
		['=SUM($B:$B)', 0, 0, '=SUM(C2)'],
		["='My Sheet'!B2+Sheet2!$A$1", 0, 0, "='My Sheet'!R[1]C[1]+Sheet2!R1C1"],
		['=A1#*2', 0, 1, '=RC[-1]#*2'],
		['="A1"&Rate', 3, 3, '="A1"&Rate'],
	])('%s at (%i,%i) -> %s', (formula, row, col, expected) => {
		expect(a1ToR1C1(formula, row, col)).toBe(expected);
	});
});

describe('r1c1ToA1', () => {
	it.each([
		['=RC', 0, 0, '=A1'],
		['=R[-1]C[-1]', 1, 1, '=A1'],
		['=R1C1', 5, 5, '=$A$1'],
		['=R1C[-1]', 1, 1, '=A$1'],
		['=SUM(R[-2]C[-2]:R[-1]C[-1])', 2, 2, '=SUM(A1:B2)'],
		['=SUM(R:R[2])', 0, 0, '=SUM(1:3)'],
		['=SUM(R5)', 0, 0, '=SUM($5:$5)'],
		['=SUM(C2)', 0, 0, '=SUM($B:$B)'],
		['=SUM(C[1]:C[3])', 0, 0, '=SUM(B:D)'],
		['=ROUND(RC[1],0)+Rate', 0, 0, '=ROUND(B1,0)+Rate'],
		["=Sheet2!R1C1+'My Sheet'!RC", 1, 1, "=Sheet2!$A$1+'My Sheet'!B2"],
		['=r[1]c', 0, 0, '=A2'],
		['="RC"&A1', 0, 0, '="RC"&A1'],
	])('%s at (%i,%i) -> %s', (formula, row, col, expected) => {
		expect(r1c1ToA1(formula, row, col)).toBe(expected);
	});

	it('wraps relative references around the grid edges', () => {
		expect(r1c1ToA1('=R[-1]C', 0, 0)).toBe('=A1048576');
		expect(r1c1ToA1('=RC[-1]', 0, 0)).toBe('=XFD1');
	});

	it('round-trips with a1ToR1C1', () => {
		for (const f of [
			'=SUM(A1:B9)*$C$3',
			'=A$2+$B3',
			'=SUM(D:F)+SUM($2:4)',
			"='X Y'!C3:D4",
			'=IF(A1>0,B1,C1)',
		]) {
			for (const [row, col] of [
				[0, 0],
				[10, 3],
				[3, 10],
			]) {
				expect(r1c1ToA1(a1ToR1C1(f, row, col), row, col)).toBe(f);
			}
		}
	});
});
