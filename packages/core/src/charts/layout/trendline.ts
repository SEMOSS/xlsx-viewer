/** Trendline fitting (linear, exponential, logarithmic, power, polynomial, moving average). */
import type { ChartTrendline } from '../../types/chart';

export type TrendFunction = (x: number) => number;

/** Solve a small linear system with Gaussian elimination (partial pivoting). */
function solve(a: number[][], b: number[]): number[] | null {
	const n = b.length;
	const m = a.map((row, i) => [...row, b[i]]);
	for (let c = 0; c < n; c++) {
		let pivot = c;
		for (let r = c + 1; r < n; r++) {
			if (Math.abs(m[r][c]) > Math.abs(m[pivot][c])) {
				pivot = r;
			}
		}
		if (Math.abs(m[pivot][c]) < 1e-12) {
			return null;
		}
		[m[c], m[pivot]] = [m[pivot], m[c]];
		for (let r = 0; r < n; r++) {
			if (r !== c) {
				const f = m[r][c] / m[c][c];
				for (let k = c; k <= n; k++) {
					m[r][k] -= f * m[c][k];
				}
			}
		}
	}
	return m.map((row, i) => row[n] / row[i]);
}

/** Least-squares polynomial coefficients (lowest degree first). */
export function polyFit(
	xs: readonly number[],
	ys: readonly number[],
	degree: number,
): number[] | null {
	const d = Math.max(1, Math.min(6, Math.round(degree)));
	if (xs.length <= d) {
		return null;
	}
	const a: number[][] = [];
	const b: number[] = [];
	for (let r = 0; r <= d; r++) {
		a.push([]);
		for (let c = 0; c <= d; c++) {
			let s = 0;
			for (const x of xs) {
				s += x ** (r + c);
			}
			a[r].push(s);
		}
		let t = 0;
		xs.forEach((x, i) => {
			t += ys[i] * x ** r;
		});
		b.push(t);
	}
	return solve(a, b);
}

/** Fit a trendline; returns the curve, or `null` when the data does not allow the fit. */
export function fitTrendline(
	trend: ChartTrendline,
	xs: readonly number[],
	ys: readonly number[],
): TrendFunction | null {
	const pairs = xs
		.map((x, i) => [x, ys[i]] as const)
		.filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
	const px = pairs.map((p) => p[0]);
	const py = pairs.map((p) => p[1]);
	switch (trend.type) {
		case 'linear': {
			const c = polyFit(px, py, 1);
			return c ? (x) => c[0] + c[1] * x : null;
		}
		case 'poly': {
			const c = polyFit(px, py, trend.order ?? 2);
			return c ? (x) => c.reduce((s, k, i) => s + k * x ** i, 0) : null;
		}
		case 'exp': {
			if (py.some((y) => y <= 0)) {
				return null;
			}
			const c = polyFit(px, py.map(Math.log), 1);
			return c ? (x) => Math.exp(c[0]) * Math.exp(c[1] * x) : null;
		}
		case 'log': {
			if (px.some((x) => x <= 0)) {
				return null;
			}
			const c = polyFit(px.map(Math.log), py, 1);
			return c ? (x) => c[0] + c[1] * Math.log(x) : null;
		}
		case 'power': {
			if (px.some((x) => x <= 0) || py.some((y) => y <= 0)) {
				return null;
			}
			const c = polyFit(px.map(Math.log), py.map(Math.log), 1);
			return c ? (x) => Math.exp(c[0]) * x ** c[1] : null;
		}
		default:
			return null;
	}
}

/** Moving-average points `[x, y]` (the average of the previous `period` values). */
export function movingAverage(
	xs: readonly number[],
	ys: readonly number[],
	period: number,
): Array<[number, number]> {
	const p = Math.max(2, Math.round(period));
	const out: Array<[number, number]> = [];
	for (let i = p - 1; i < ys.length; i++) {
		let sum = 0;
		let ok = true;
		for (let k = i - p + 1; k <= i; k++) {
			if (!Number.isFinite(ys[k])) {
				ok = false;
				break;
			}
			sum += ys[k];
		}
		if (ok) {
			out.push([xs[i], sum / p]);
		}
	}
	return out;
}

/** Legend text Excel shows for an unnamed trendline, e.g. `Linear (Sales)`. */
export function trendlineLabel(trend: ChartTrendline, seriesName: string): string {
	if (trend.name) {
		return trend.name;
	}
	const kind: Record<ChartTrendline['type'], string> = {
		linear: 'Linear',
		exp: 'Expon.',
		log: 'Log.',
		poly: 'Poly.',
		power: 'Power',
		movingAvg: `${Math.max(2, trend.period ?? 2)} per. Mov. Avg.`,
	};
	return `${kind[trend.type]} (${seriesName})`;
}
