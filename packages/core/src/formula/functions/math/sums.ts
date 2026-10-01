import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { asArray, collectNumbers, flattenArray } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, numberArg } from './args';
import { power } from './arithmetic';

/** Compensated (Neumaier) summation: accurate sums of many terms. */
export class CompensatedSum {
	private sum = 0;
	private carry = 0;

	add(x: number): void {
		const t = this.sum + x;
		if (Math.abs(this.sum) >= Math.abs(x)) {
			this.carry += this.sum - t + x;
		} else {
			this.carry += x - t + this.sum;
		}
		this.sum = t;
	}

	get value(): number {
		const v = this.sum + this.carry;
		return Number.isNaN(v) && !Number.isNaN(this.sum) ? this.sum : v;
	}
}

/** Sum of a list of numbers (compensated). */
export function sumOf(values: readonly number[]): number {
	const acc = new CompensatedSum();
	for (const v of values) {
		acc.add(v);
	}
	return acc.value;
}

/** A SUM-style aggregate over `any` arguments (text and booleans inside ranges are ignored). */
function aggregate(name: string, reduce: (values: number[]) => number): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 255,
		params: ['any'],
		impl: (args: FunctionArg[], ctx: FunctionContext) => {
			const values = collectNumbers(args, ctx);
			return isCellError(values) ? values : finite(reduce(values));
		},
	};
}

function product(values: number[]): number {
	if (values.length === 0) {
		return 0;
	}
	let p = 1;
	for (const v of values) {
		p *= v;
	}
	return p;
}

const SUMPRODUCT: FunctionDefinition = {
	name: 'SUMPRODUCT',
	minArgs: 1,
	maxArgs: 255,
	params: ['array'],
	impl: (args) => {
		const arrays = args.map(asArray);
		const { rows, cols } = arrays[0];
		if (arrays.some((a) => a.rows !== rows || a.cols !== cols)) {
			return ERRORS.VALUE;
		}
		const acc = new CompensatedSum();
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				let p = 1;
				for (const a of arrays) {
					const v = a.values[r][c];
					if (isCellError(v)) {
						return v;
					}
					// Text, booleans and blanks count as 0.
					p *= typeof v === 'number' ? v : 0;
				}
				acc.add(p);
			}
		}
		return finite(acc.value);
	},
};

/** SUMX2MY2 / SUMX2PY2 / SUMXMY2: pairs where either value is not a number are skipped. */
function pairwise(name: string, term: (x: number, y: number) => number): FunctionDefinition {
	return {
		name,
		minArgs: 2,
		maxArgs: 2,
		params: ['array', 'array'],
		impl: (args) => {
			const xs = flattenArray(asArray(args[0]));
			const ys = flattenArray(asArray(args[1]));
			if (xs.length !== ys.length) {
				return ERRORS.NA;
			}
			const acc = new CompensatedSum();
			for (let i = 0; i < xs.length; i++) {
				const x = xs[i];
				const y = ys[i];
				if (isCellError(x)) {
					return x;
				}
				if (isCellError(y)) {
					return y;
				}
				if (typeof x === 'number' && typeof y === 'number') {
					acc.add(term(x, y));
				}
			}
			return finite(acc.value);
		},
	};
}

const SERIESSUM: FunctionDefinition = {
	name: 'SERIESSUM',
	minArgs: 4,
	maxArgs: 4,
	params: ['value', 'value', 'value', 'array'],
	impl: (args, ctx) => {
		const parts: number[] = [];
		for (let i = 0; i < 3; i++) {
			const v = numberArg(args[i], ctx);
			if (typeof v !== 'number') {
				return v;
			}
			parts.push(v);
		}
		const [x, n, m] = parts;
		const coefficients = flattenArray(asArray(args[3]));
		const acc = new CompensatedSum();
		for (let i = 0; i < coefficients.length; i++) {
			const a = coefficients[i];
			if (isCellError(a)) {
				return a;
			}
			if (typeof a !== 'number') {
				return ERRORS.VALUE;
			}
			const p: number | CellError = power(x, n + i * m);
			if (typeof p !== 'number') {
				return p;
			}
			acc.add(a * p);
		}
		return finite(acc.value);
	},
};

export const SUM_FUNCTIONS: FunctionDefinition[] = [
	aggregate('SUM', sumOf),
	aggregate('PRODUCT', product),
	aggregate('SUMSQ', (values) => sumOf(values.map((v) => v * v))),
	SUMPRODUCT,
	pairwise('SUMX2MY2', (x, y) => x * x - y * y),
	pairwise('SUMX2PY2', (x, y) => x * x + y * y),
	pairwise('SUMXMY2', (x, y) => (x - y) * (x - y)),
	SERIESSUM,
];
