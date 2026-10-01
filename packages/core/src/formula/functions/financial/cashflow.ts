/**
 * Cash-flow analysis: NPV, IRR, MIRR, XNPV and XIRR.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { asArray, collectNumbers, flattenArray, toNumber } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, numArg, optNum } from '../datetime/args';
import { maxSerial } from '../datetime/serial';
import { newton, scanForRoot } from './solver';

function npvImpl(args: FunctionArg[], ctx: FunctionContext) {
	const rate = numArg(args[0], ctx);
	if (isCellError(rate)) {
		return rate;
	}
	const values = collectNumbers(args.slice(1), ctx);
	if (isCellError(values)) {
		return values;
	}
	if (rate === -1) {
		return ERRORS.DIV0;
	}
	let npv = 0;
	for (let i = 0; i < values.length; i++) {
		npv += values[i] / Math.pow(1 + rate, i + 1);
	}
	return finite(npv);
}

function hasBothSigns(values: readonly number[]): boolean {
	return values.some((v) => v > 0) && values.some((v) => v < 0);
}

/** Internal rate of return of evenly spaced cash flows (Newton from `guess`, then a bracketing search). */
export function solveIrr(values: readonly number[], guess: number): number | null {
	if (!hasBothSigns(values)) {
		return null;
	}
	const fdf = (r: number): [number, number] => {
		if (r <= -1) {
			return [Number.NaN, Number.NaN];
		}
		let f = 0;
		let df = 0;
		for (let i = 0; i < values.length; i++) {
			const p = Math.pow(1 + r, i);
			f += values[i] / p;
			df -= (i * values[i]) / (p * (1 + r));
		}
		return [f, df];
	};
	const root = newton(fdf, guess, { maxIterations: 50, tolerance: 1e-12 });
	if (root !== null && root > -1) {
		return root;
	}
	return scanForRoot(fdf, guess);
}

function irrImpl(args: FunctionArg[], ctx: FunctionContext) {
	const values = collectNumbers([args[0]], ctx);
	const guess = optNum(args[1], ctx, 0.1);
	if (isCellError(values)) {
		return values;
	}
	if (isCellError(guess)) {
		return guess;
	}
	const r = solveIrr(values, guess);
	return r === null ? ERRORS.NUM : r;
}

function mirrImpl(args: FunctionArg[], ctx: FunctionContext) {
	const values = collectNumbers([args[0]], ctx);
	const financeRate = numArg(args[1], ctx);
	const reinvestRate = numArg(args[2], ctx);
	if (isCellError(values)) {
		return values;
	}
	if (isCellError(financeRate)) {
		return financeRate;
	}
	if (isCellError(reinvestRate)) {
		return reinvestRate;
	}
	const n = values.length;
	if (!hasBothSigns(values) || financeRate === -1 || reinvestRate === -1) {
		return ERRORS.DIV0;
	}
	let pvNegative = 0;
	let pvPositive = 0;
	for (let i = 0; i < n; i++) {
		const v = values[i];
		if (v > 0) {
			pvPositive += v / Math.pow(1 + reinvestRate, i);
		} else if (v < 0) {
			pvNegative += v / Math.pow(1 + financeRate, i);
		}
	}
	const ratio = (-pvPositive * Math.pow(1 + reinvestRate, n - 1)) / pvNegative;
	return finite(Math.pow(ratio, 1 / (n - 1)) - 1);
}

interface DatedFlows {
	values: number[];
	/** Years from the first date (365-day years). */
	times: number[];
}

/**
 * Pair up the values and dates of XNPV / XIRR. Pairs where both cells are
 * blank are skipped; anything else that is not numeric is #VALUE!, dates
 * before the first date are #NUM!, and differing sizes are #NUM!.
 */
function datedFlows(
	valuesArg: FunctionArg,
	datesArg: FunctionArg,
	ctx: FunctionContext,
): DatedFlows | CellError {
	const values = flattenArray(asArray(valuesArg));
	const dates = flattenArray(asArray(datesArg));
	if (values.length !== dates.length) {
		return ERRORS.NUM;
	}
	const out: DatedFlows = { values: [], times: [] };
	let first: number | undefined;
	for (let i = 0; i < values.length; i++) {
		const v = values[i];
		const d = dates[i];
		if (isCellError(v)) {
			return v;
		}
		if (isCellError(d)) {
			return d;
		}
		if (v === null && d === null) {
			continue;
		}
		if (typeof v !== 'number') {
			return ERRORS.VALUE;
		}
		if (d === null || typeof d === 'boolean') {
			return ERRORS.VALUE;
		}
		const serial = toNumber(d, ctx.date1904);
		if (isCellError(serial)) {
			return serial;
		}
		const day = Math.floor(serial);
		if (day < 0 || day > maxSerial(ctx.date1904)) {
			return ERRORS.NUM;
		}
		first ??= day;
		if (day < first) {
			return ERRORS.NUM;
		}
		out.values.push(v);
		out.times.push((day - first) / 365);
	}
	return out.values.length === 0 ? ERRORS.NUM : out;
}

function xnpvOf(rate: number, flows: DatedFlows): number {
	let sum = 0;
	for (let i = 0; i < flows.values.length; i++) {
		sum += flows.values[i] / Math.pow(1 + rate, flows.times[i]);
	}
	return sum;
}

function xnpvImpl(args: FunctionArg[], ctx: FunctionContext) {
	const rate = numArg(args[0], ctx);
	if (isCellError(rate)) {
		return rate;
	}
	const flows = datedFlows(args[1], args[2], ctx);
	if (isCellError(flows)) {
		return flows;
	}
	return rate <= -1 ? ERRORS.NUM : finite(xnpvOf(rate, flows));
}

function xirrImpl(args: FunctionArg[], ctx: FunctionContext) {
	const flows = datedFlows(args[0], args[1], ctx);
	const guess = optNum(args[2], ctx, 0.1);
	if (isCellError(flows)) {
		return flows;
	}
	if (isCellError(guess)) {
		return guess;
	}
	if (!hasBothSigns(flows.values) || guess <= -1) {
		return ERRORS.NUM;
	}
	const fdf = (r: number): [number, number] => {
		if (r <= -1) {
			return [Number.NaN, Number.NaN];
		}
		let f = 0;
		let df = 0;
		for (let i = 0; i < flows.values.length; i++) {
			const t = flows.times[i];
			const p = Math.pow(1 + r, t);
			f += flows.values[i] / p;
			df -= (t * flows.values[i]) / (p * (1 + r));
		}
		return [f, df];
	};
	let root = newton(fdf, guess, { maxIterations: 100, tolerance: 1e-12 });
	if (root === null || root <= -1) {
		root = scanForRoot(fdf, guess);
	}
	return root === null ? ERRORS.NUM : root;
}

export const CASHFLOW_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'NPV', minArgs: 2, maxArgs: Infinity, params: ['value', 'any'], impl: npvImpl },
	{ name: 'IRR', minArgs: 1, maxArgs: 2, params: ['any', 'value'], impl: irrImpl },
	{ name: 'MIRR', minArgs: 3, maxArgs: 3, params: ['any', 'value', 'value'], impl: mirrImpl },
	{ name: 'XNPV', minArgs: 3, maxArgs: 3, params: ['value', 'array', 'array'], impl: xnpvImpl },
	{ name: 'XIRR', minArgs: 2, maxArgs: 3, params: ['array', 'array', 'value'], impl: xirrImpl },
];
