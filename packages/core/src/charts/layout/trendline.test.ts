import { describe, expect, it } from 'vitest';
import { fitTrendline, movingAverage, polyFit, trendlineLabel } from './trendline';

const xs = [1, 2, 3, 4, 5];

describe('trendlines', () => {
	it('fits linear and polynomial least squares', () => {
		const lin = fitTrendline(
			{ type: 'linear' },
			xs,
			xs.map((x) => 2 * x + 1),
		);
		expect(lin?.(10)).toBeCloseTo(21);
		const poly = fitTrendline(
			{ type: 'poly', order: 2 },
			xs,
			xs.map((x) => x * x - x),
		);
		expect(poly?.(6)).toBeCloseTo(30);
		expect(polyFit([1], [1], 2)).toBeNull();
	});

	it('fits exponential, logarithmic and power curves', () => {
		expect(
			fitTrendline(
				{ type: 'exp' },
				xs,
				xs.map((x) => 3 * Math.exp(0.5 * x)),
			)?.(6),
		).toBeCloseTo(3 * Math.exp(3));
		expect(
			fitTrendline(
				{ type: 'log' },
				xs,
				xs.map((x) => 2 + 3 * Math.log(x)),
			)?.(10),
		).toBeCloseTo(2 + 3 * Math.log(10));
		expect(
			fitTrendline(
				{ type: 'power' },
				xs,
				xs.map((x) => 4 * x ** 1.5),
			)?.(9),
		).toBeCloseTo(108);
		expect(fitTrendline({ type: 'exp' }, xs, [1, -1, 2, 3, 4])).toBeNull();
		expect(fitTrendline({ type: 'linear' }, [1], [1])).toBeNull();
	});

	it('computes moving averages and skips windows with gaps', () => {
		expect(movingAverage(xs, [2, 4, 6, 8, 10], 2)).toEqual([
			[2, 3],
			[3, 5],
			[4, 7],
			[5, 9],
		]);
		expect(movingAverage(xs, [2, NaN, 6, 8, 10], 2)).toEqual([
			[4, 7],
			[5, 9],
		]);
	});

	it('labels unnamed trendlines like Excel', () => {
		expect(trendlineLabel({ type: 'linear' }, 'Sales')).toBe('Linear (Sales)');
		expect(trendlineLabel({ type: 'movingAvg', period: 3 }, 'Sales')).toBe(
			'3 per. Mov. Avg. (Sales)',
		);
		expect(trendlineLabel({ type: 'exp', name: 'Growth' }, 'Sales')).toBe('Growth');
	});
});
