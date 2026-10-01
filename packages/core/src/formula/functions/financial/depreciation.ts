/**
 * Depreciation: SLN, SYD, DB, DDB and VDB.
 */
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { boolArg, finite, readNumbers, REQUIRED } from '../datetime/args';

function slnImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	const [cost, salvage, life] = n;
	return life === 0 ? ERRORS.DIV0 : (cost - salvage) / life;
}

function sydImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, REQUIRED]);
	if (isCellError(n)) {
		return n;
	}
	const [cost, salvage, life, per] = n;
	if (life <= 0 || per <= 0 || per > life) {
		return ERRORS.NUM;
	}
	return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1));
}

/**
 * Fixed-declining balance (DB). The rate is rounded to three decimals; the
 * first and the extra last period are pro-rated by `month` / 12.
 */
function dbImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, REQUIRED, 12]);
	if (isCellError(n)) {
		return n;
	}
	const [cost, salvage, life, period] = n;
	const month = Math.trunc(n[4]);
	if (cost <= 0 || salvage < 0 || salvage > cost || life <= 0 || period <= 0) {
		return ERRORS.NUM;
	}
	if (month < 1 || month > 12 || period > life + 1) {
		return ERRORS.NUM;
	}
	const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000;
	const first = (cost * rate * month) / 12;
	if (Math.floor(period) === 1) {
		return first;
	}
	let total = first;
	let dep = 0;
	const last = Math.floor(Math.min(life, period));
	for (let i = 2; i <= last; i++) {
		dep = (cost - total) * rate;
		total += dep;
	}
	if (period > life) {
		dep = ((cost - total) * rate * (12 - month)) / 12;
	}
	return finite(dep);
}

/** Double-declining (or `factor`-declining) balance for one period, never below the salvage value. */
export function ddbOf(
	cost: number,
	salvage: number,
	life: number,
	period: number,
	factor: number,
): number {
	let rate = factor / life;
	let before: number;
	if (rate >= 1) {
		rate = 1;
		before = period === 1 ? cost : 0;
	} else {
		before = cost * Math.pow(1 - rate, period - 1);
	}
	const after = cost * Math.pow(1 - rate, period);
	const dep = after < salvage ? before - salvage : before - after;
	return dep < 0 ? 0 : dep;
}

function ddbImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, REQUIRED, 2]);
	if (isCellError(n)) {
		return n;
	}
	const [cost, salvage, life, period, factor] = n;
	if (
		cost < 0 ||
		salvage < 0 ||
		salvage > cost ||
		factor <= 0 ||
		life <= 0 ||
		period < 1 ||
		period > life
	) {
		return ERRORS.NUM;
	}
	return finite(ddbOf(cost, salvage, life, period, factor));
}

/**
 * Depreciation over the first `period` periods (which may be fractional),
 * switching to straight line once that gives more.
 */
function vdbWhole(
	cost: number,
	salvage: number,
	life: number,
	life1: number,
	period: number,
	factor: number,
): number {
	const end = Math.ceil(period);
	let remaining = cost - salvage;
	let switched = false;
	let sln = 0;
	let total = 0;
	for (let i = 1; i <= end; i++) {
		let term: number;
		if (switched) {
			term = sln;
		} else {
			const ddb = ddbOf(cost, salvage, life, i, factor);
			sln = remaining / (life1 - (i - 1));
			if (sln > ddb) {
				term = sln;
				switched = true;
			} else {
				term = ddb;
				remaining -= ddb;
			}
		}
		if (i === end) {
			term *= period + 1 - end;
		}
		total += term;
	}
	return total;
}

function vdbOf(
	cost: number,
	salvage: number,
	life: number,
	start: number,
	end: number,
	factor: number,
	noSwitch: boolean,
): number {
	const intStart = Math.floor(start);
	const intEnd = Math.ceil(end);
	if (noSwitch) {
		let total = 0;
		for (let i = intStart + 1; i <= intEnd; i++) {
			let term = ddbOf(cost, salvage, life, i, factor);
			if (i === intStart + 1) {
				term *= Math.min(end, intStart + 1) - start;
			} else if (i === intEnd) {
				term *= end + 1 - intEnd;
			}
			total += term;
		}
		return total;
	}
	let part = 0;
	if (start !== intStart) {
		const value = cost - vdbWhole(cost, salvage, life, life, intStart, factor);
		part += (start - intStart) * vdbWhole(value, salvage, life, life - intStart, 1, factor);
	}
	if (end !== intEnd) {
		const tempStart = intEnd - 1;
		const value = cost - vdbWhole(cost, salvage, life, life, tempStart, factor);
		part += (intEnd - end) * vdbWhole(value, salvage, life, life - tempStart, 1, factor);
	}
	const base = cost - vdbWhole(cost, salvage, life, life, intStart, factor);
	return vdbWhole(base, salvage, life, life - intStart, intEnd - intStart, factor) - part;
}

function vdbImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = readNumbers(args, ctx, [REQUIRED, REQUIRED, REQUIRED, REQUIRED, REQUIRED, 2]);
	const noSwitch = boolArg(args[6], false);
	if (isCellError(n)) {
		return n;
	}
	if (isCellError(noSwitch)) {
		return noSwitch;
	}
	const [cost, salvage, life, start, end, factor] = n;
	if (cost < 0 || salvage < 0 || salvage > cost || life <= 0 || factor <= 0) {
		return ERRORS.NUM;
	}
	if (start < 0 || end < start || end > life) {
		return ERRORS.NUM;
	}
	return finite(vdbOf(cost, salvage, life, start, end, factor, noSwitch));
}

export const DEPRECIATION_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'SLN', minArgs: 3, maxArgs: 3, params: ['value'], impl: slnImpl },
	{ name: 'SYD', minArgs: 4, maxArgs: 4, params: ['value'], impl: sydImpl },
	{ name: 'DB', minArgs: 4, maxArgs: 5, params: ['value'], impl: dbImpl },
	{ name: 'DDB', minArgs: 4, maxArgs: 5, params: ['value'], impl: ddbImpl },
	{ name: 'VDB', minArgs: 5, maxArgs: 7, params: ['value'], impl: vdbImpl },
];
