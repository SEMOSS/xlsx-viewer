/**
 * Ordinary least squares y = b + sum m_j x_j for LINEST, LOGEST, TREND and
 * GROWTH. With a constant the columns are centred first (which also removes
 * the intercept from the system). A Householder QR factorisation then
 * processes the x columns in order and drops a column that is (numerically)
 * a linear combination of the ones before it, the way Excel reports
 * collinear columns: coefficient 0, standard error 0, and one more degree of
 * freedom.
 */

export interface LeastSquaresFit {
	/** Coefficient of each x column (0 for columns dropped as collinear). */
	coefficients: number[];
	intercept: number;
	/** se_j = sey * sqrt(factor_j); 0 for dropped columns. */
	coefficientFactors: number[];
	/** se_b = sey * sqrt(interceptFactor); NaN without a constant. */
	interceptFactor: number;
	/** Number of x columns kept in the model. */
	rank: number;
	n: number;
	df: number;
	ssreg: number;
	ssresid: number;
}

/** Relative residual norm below which a column counts as collinear. */
const COLLINEAR_TOLERANCE = 1e-12;

function dot(a: Float64Array, b: Float64Array, from: number): number {
	let s = 0;
	for (let i = from; i < a.length; i++) {
		s += a[i] * b[i];
	}
	return s;
}

/** Apply the reflection I - 2 v v' / (v' v) (v supported on rows >= offset) to `target`. */
function reflect(target: Float64Array, v: Float64Array, offset: number, vv: number): void {
	let d = 0;
	for (let i = 0; i < v.length; i++) {
		d += v[i] * target[offset + i];
	}
	const f = (2 * d) / vv;
	for (let i = 0; i < v.length; i++) {
		target[offset + i] -= f * v[i];
	}
}

/** x has one row per observation and one column per variable. */
export function leastSquares(
	x: readonly number[][],
	y: readonly number[],
	withConstant: boolean,
): LeastSquaresFit {
	const n = y.length;
	const k = x[0]?.length ?? 0;
	const xMeans = new Array<number>(k).fill(0);
	let yMean = 0;
	if (withConstant) {
		for (let i = 0; i < n; i++) {
			yMean += y[i];
			for (let j = 0; j < k; j++) {
				xMeans[j] += x[i][j];
			}
		}
		yMean /= n;
		for (let j = 0; j < k; j++) {
			xMeans[j] /= n;
		}
	}
	const cols = Array.from({ length: k }, (_, j) =>
		Float64Array.from(x, (row) => row[j] - xMeans[j]),
	);
	const b = Float64Array.from(y, (v) => v - yMean);
	const norms = cols.map((c) => Math.sqrt(dot(c, c, 0)));

	// Householder QR, skipping columns that add nothing new.
	const kept: number[] = [];
	for (let j = 0; j < k && kept.length < n; j++) {
		const r = kept.length;
		const col = cols[j];
		const s = Math.sqrt(dot(col, col, r));
		if (norms[j] === 0 || s <= COLLINEAR_TOLERANCE * norms[j]) {
			continue;
		}
		const alpha = col[r] > 0 ? -s : s;
		const v = col.slice(r);
		v[0] -= alpha;
		const vv = dot(v, v, 0);
		for (let l = j; l < k; l++) {
			reflect(cols[l], v, r, vv);
		}
		reflect(b, v, r, vv);
		kept.push(j);
	}
	const rank = kept.length;
	const R = (p: number, q: number) => cols[kept[q]][p];

	// Back substitution for the kept coefficients.
	const beta = new Array<number>(rank).fill(0);
	for (let p = rank - 1; p >= 0; p--) {
		let s = b[p];
		for (let q = p + 1; q < rank; q++) {
			s -= R(p, q) * beta[q];
		}
		beta[p] = s / R(p, p);
	}
	const coefficients = new Array<number>(k).fill(0);
	kept.forEach((j, q) => {
		coefficients[j] = beta[q];
	});
	if (k === 1 && rank === 1) {
		// Simple regression: the same arithmetic as SLOPE, so the functions agree exactly.
		let sxy = 0;
		let sxx = 0;
		for (let i = 0; i < n; i++) {
			const dx = x[i][0] - xMeans[0];
			sxy += dx * (y[i] - yMean);
			sxx += dx * dx;
		}
		coefficients[0] = sxy / sxx;
	}
	let intercept = 0;
	if (withConstant) {
		intercept = yMean;
		for (let j = 0; j < k; j++) {
			intercept -= coefficients[j] * xMeans[j];
		}
	}

	let ssresid = 0;
	let ssreg = 0;
	for (let i = 0; i < n; i++) {
		let fitted = intercept;
		for (let j = 0; j < k; j++) {
			fitted += coefficients[j] * x[i][j];
		}
		const res = y[i] - fitted;
		const dev = withConstant ? fitted - yMean : fitted;
		ssresid += res * res;
		ssreg += dev * dev;
	}
	// Rounding-level residuals of an exact fit count as zero.
	if (ssresid <= 1e-28 * (ssreg + ssresid)) {
		ssresid = 0;
	}

	// (X'X)^-1 = R^-1 R^-T: the diagonal gives the coefficient variance factors.
	const inv: number[][] = Array.from({ length: rank }, () => new Array<number>(rank).fill(0));
	for (let c = 0; c < rank; c++) {
		for (let p = c; p >= 0; p--) {
			let s = p === c ? 1 : 0;
			for (let q = p + 1; q <= c; q++) {
				s -= R(p, q) * inv[q][c];
			}
			inv[p][c] = s / R(p, p);
		}
	}
	const coefficientFactors = new Array<number>(k).fill(0);
	kept.forEach((j, p) => {
		let s = 0;
		for (let c = p; c < rank; c++) {
			s += inv[p][c] * inv[p][c];
		}
		coefficientFactors[j] = s;
	});
	let interceptFactor = NaN;
	if (withConstant) {
		// 1/n + m' (R'R)^-1 m with m the kept column means: solve R' w = m.
		const w = new Array<number>(rank).fill(0);
		let ww = 0;
		for (let p = 0; p < rank; p++) {
			let s = xMeans[kept[p]];
			for (let q = 0; q < p; q++) {
				s -= R(q, p) * w[q];
			}
			w[p] = s / R(p, p);
			ww += w[p] * w[p];
		}
		interceptFactor = 1 / n + ww;
	}
	if (k === 1 && rank === 1) {
		const f = 1 / (norms[0] * norms[0]);
		coefficientFactors[0] = f;
		interceptFactor = withConstant ? 1 / n + xMeans[0] * xMeans[0] * f : NaN;
	}
	return {
		coefficients,
		intercept,
		coefficientFactors,
		interceptFactor,
		rank,
		n,
		df: n - rank - (withConstant ? 1 : 0),
		ssreg,
		ssresid,
	};
}
