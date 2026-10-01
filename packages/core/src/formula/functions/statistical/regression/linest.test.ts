import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../../values/errors';
import type { MissingNode } from '../../../ast';
import type { Scalar } from '../../../types';
import { arr, MockContext } from '../../../test-utils/mock-context';
import { column, expectClose, rowsOf, stat } from '../test-utils/expect';
import { leastSquares } from './least-squares';

const MISSING: MissingNode = { type: 'missing' };
const NA = ERRORS.NA;

function expectBlock(actual: Scalar[][], expected: (number | typeof NA)[][], rel = 1e-9): void {
	expect(actual.length).toBe(expected.length);
	expected.forEach((row, r) => {
		expect(actual[r].length).toBe(row.length);
		row.forEach((v, c) => {
			if (typeof v === 'number') {
				expectClose(actual[r][c], v, rel);
			} else {
				expect(actual[r][c]).toBe(v);
			}
		});
	});
}

describe('LINEST', () => {
	it('fits a line (Excel example 1)', () => {
		expect(rowsOf(stat('LINEST', [column(1, 9, 5, 7), column(0, 4, 2, 3)]))).toEqual([[2, 1]]);
	});

	it('reports exact-fit statistics, with F as #NUM!', () => {
		const out = rowsOf(stat('LINEST', [column(1, 9, 5, 7), column(0, 4, 2, 3), true, true]));
		expect(out).toEqual([
			[2, 1],
			[0, 0],
			[1, 0],
			[ERRORS.NUM, 2],
			[35, 0],
		]);
	});

	it('reports #NUM! for statistics that need a residual degree of freedom', () => {
		const out = rowsOf(stat('LINEST', [column(5, 7), column(1, 2), true, true]));
		expect(out[0]).toEqual([2, 3]);
		expect(out[1]).toEqual([ERRORS.NUM, ERRORS.NUM]);
		expect(out[3]).toEqual([ERRORS.NUM, 0]);
	});

	it('runs a multiple regression (Excel example 3, office buildings)', () => {
		const ctx = new MockContext();
		const data = [
			[2310, 2, 2, 20, 142000],
			[2333, 2, 2, 12, 144000],
			[2356, 3, 1.5, 33, 151000],
			[2379, 3, 2, 43, 150000],
			[2402, 2, 3, 53, 139000],
			[2425, 4, 2, 23, 169000],
			[2448, 2, 1.5, 99, 126000],
			[2471, 2, 2, 34, 142900],
			[2494, 3, 3, 23, 163000],
			[2517, 4, 4, 55, 169000],
			[2540, 2, 3, 22, 149000],
		];
		data.forEach((row, i) => {
			ctx.setCells({
				[`A${i + 2}`]: row[0],
				[`B${i + 2}`]: row[1],
				[`C${i + 2}`]: row[2],
				[`D${i + 2}`]: row[3],
				[`E${i + 2}`]: row[4],
			});
		});
		const out = rowsOf(stat('LINEST', [ctx.ref('E2:E12'), ctx.ref('A2:D12'), true, true], ctx));
		// Exact values (50-digit normal equations); Excel shows the same to 10 digits.
		expectBlock(out, [
			[
				-234.23716447120242, 2553.210660391538, 12529.768167086751, 27.641387366020286,
				52317.83050729132,
			],
			[
				13.268011475500364, 530.6691519303783, 400.0668381939531, 5.429374041545316,
				12237.361602862353,
			],
			[0.9967479933845101, 970.5784629285063, NA, NA, NA],
			[459.75367422539244, 6, NA, NA, NA],
			[1732393319.2292507, 5652135.31620397, NA, NA, NA],
		]);
	});

	it('fits two variables (exact reference)', () => {
		const y = column(3, 5, 6, 9, 11, 14);
		const x = arr([1, 2], [2, 1], [3, 4], [4, 3], [5, 7], [6, 5]);
		const out = rowsOf(stat('LINEST', [y, x, true, true]));
		expectBlock(
			out,
			[
				[-0.20021881838074398, 2.3544857768052516, 0.4934354485776805],
				[0.20863905175088468, 0.24091562538368305, 0.5812724706656144],
				[0.986467125143274, 0.6155651841911868, NA],
				[109.34119345524543, 3, NA],
				[82.86323851203501, 1.136761487964989, NA],
			],
			1e-12,
		);
	});

	it('forces the line through the origin when const is FALSE', () => {
		const out = rowsOf(stat('LINEST', [column(1, 9, 5, 7.5), column(0, 4, 2, 3), false, true]));
		expectBlock(
			out,
			[
				[2.3620689655172415, 0],
				[0.12902266850944624, NA],
				[0.9911284786397001, 0.6948083337796512],
				[335.1607142857143, 3],
				[161.80172413793105, 1.4482758620689655],
			],
			1e-13,
		);
	});

	it('drops collinear columns: zero coefficient, zero error, one more df', () => {
		const y = column(1, 3, 2, 5, 4);
		const x = arr([1, 2], [2, 4], [3, 6], [4, 8], [5, 10]);
		const out = rowsOf(stat('LINEST', [y, x, true, true]));
		const simple = rowsOf(stat('LINEST', [y, column(1, 2, 3, 4, 5), true, true]));
		// x2 = 2 x1 is dropped, reported first because the order is reversed
		expect(out[0][0]).toBe(0);
		expect(out[1][0]).toBe(0);
		expectClose(out[0][1], simple[0][0] as number, 1e-14);
		expectClose(out[0][2], simple[0][1] as number, 1e-14);
		expect(out[3][1]).toBe(3);
		// a constant column is collinear with the intercept
		const constant = rowsOf(stat('LINEST', [y, arr([1, 7], [2, 7], [3, 7], [4, 7], [5, 7])]));
		expect(constant[0][0]).toBe(0);
		expectClose(constant[0][1], 0.8, 1e-14);
	});

	it('defaults known_x to 1, 2, 3, ... and handles rows', () => {
		expect(rowsOf(stat('LINEST', [column(3, 5, 7)]))).toEqual([[2, 1]]);
		expect(rowsOf(stat('LINEST', [arr([3, 5, 7]), MISSING, true, false]))).toEqual([[2, 1]]);
		// y in a row: each row of x is a variable
		const out = rowsOf(stat('LINEST', [arr([3, 5, 7, 9]), arr([1, 2, 3, 4], [0, 1, 0, 1])]));
		expectClose(out[0][1], 2, 1e-14);
		expect(Math.abs(out[0][0] as number)).toBeLessThan(1e-14);
	});

	it('validates the data', () => {
		expect(stat('LINEST', [column(1, 2, 3), column(1, 2)])).toBe(ERRORS.REF);
		expect(stat('LINEST', [column(1, 'a', 3), column(1, 2, 3)])).toBe(ERRORS.VALUE);
		expect(stat('LINEST', [column(1, null, 3), column(1, 2, 3)])).toBe(ERRORS.VALUE);
		expect(stat('LINEST', [column(1, ERRORS.DIV0, 3), column(1, 2, 3)])).toBe(ERRORS.DIV0);
		const ctx = new MockContext({ A1: 1, A2: 2, B1: 1, B2: 2 });
		expect(stat('LINEST', [ctx.ref('A1:A3'), ctx.ref('B1:B3')], ctx)).toBe(ERRORS.VALUE);
	});
});

describe('LOGEST', () => {
	const units = column(33100, 47300, 69000, 102000, 150000, 220000);
	const months = column(11, 12, 13, 14, 15, 16);

	it('matches Excel', () => {
		const out = rowsOf(stat('LOGEST', [units, months, true, true]));
		expectBlock(
			out,
			[
				[1.4632756281161758, 495.30477015872754],
				[0.0026334028914250613, 0.035834282435717624],
				[0.9998086197758176, 0.011016314665073015],
				[20896.80109942032, 4],
				[2.5360188299385666, 0.00048543675519961113],
			],
			1e-11,
		);
	});

	it('sets b to 1 without a constant and rejects non-positive y', () => {
		const out = rowsOf(stat('LOGEST', [column(2, 4, 8), column(1, 2, 3), false]));
		expectClose(out[0][0], 2, 1e-14);
		expect(out[0][1]).toBe(1);
		expect(stat('LOGEST', [column(1, -2, 3), column(1, 2, 3)])).toBe(ERRORS.NUM);
	});
});

describe('leastSquares', () => {
	it('reports degrees of freedom and rank', () => {
		const fit = leastSquares([[1], [2], [3]], [2, 4, 6.5], true);
		expect(fit.rank).toBe(1);
		expect(fit.df).toBe(1);
		expect(fit.n).toBe(3);
		const saturated = leastSquares([[1], [2]], [5, 7], true);
		expect(saturated.df).toBe(0);
		expect(saturated.ssresid).toBe(0);
	});
});
