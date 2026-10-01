/**
 * French accounting depreciation: AMORLINC (linear) and AMORDEGRC
 * (degressive), pro-rated for the first period with YEARFRAC.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { intArg, readDates, readNumbers, REQUIRED } from '../datetime/args';
import { yearFrac, type Basis } from '../datetime/day-count';

interface AmorArgs {
	cost: number;
	purchased: number;
	firstPeriod: number;
	salvage: number;
	period: number;
	rate: number;
	basis: Basis;
}

/** Shared argument parsing: basis 2 is not supported by these functions. */
function readAmorArgs(args: FunctionArg[], ctx: FunctionContext): AmorArgs | CellError {
	const dates = readDates([args[1], args[2]], ctx, 2, ERRORS.VALUE);
	const n = readNumbers([args[0], args[3], args[4], args[5]], ctx, [
		REQUIRED,
		REQUIRED,
		REQUIRED,
		REQUIRED,
	]);
	const basis = intArg(args[6], ctx, 0);
	if (isCellError(dates)) {
		return dates;
	}
	if (isCellError(n)) {
		return n;
	}
	if (isCellError(basis)) {
		return basis;
	}
	const [cost, salvage, period, rate] = n;
	const [purchased, firstPeriod] = dates;
	if (basis !== 0 && basis !== 1 && basis !== 3 && basis !== 4) {
		return ERRORS.NUM;
	}
	if (
		cost < 0 ||
		salvage < 0 ||
		salvage > cost ||
		period < 0 ||
		rate <= 0 ||
		purchased > firstPeriod
	) {
		return ERRORS.NUM;
	}
	return { cost, purchased, firstPeriod, salvage, period: Math.trunc(period), rate, basis };
}

function amorlincImpl(args: FunctionArg[], ctx: FunctionContext) {
	const a = readAmorArgs(args, ctx);
	if (isCellError(a)) {
		return a;
	}
	const fullRate = a.cost * a.rate;
	const firstRate = yearFrac(a.purchased, a.firstPeriod, a.basis, ctx.date1904) * a.rate * a.cost;
	const fullPeriods = Math.trunc((a.cost - a.salvage - firstRate) / fullRate);
	let result = 0;
	if (a.period === 0) {
		result = firstRate;
	} else if (a.period <= fullPeriods) {
		result = fullRate;
	} else if (a.period === fullPeriods + 1) {
		result = a.cost - a.salvage - fullRate * fullPeriods - firstRate;
	}
	return result > 0 ? result : 0;
}

/**
 * Depreciation coefficient for the asset life 1 / rate. Excel returns
 * #NUM! for lives strictly between 0-1, 1-2, 2-3 and 4-5 years.
 */
function degressiveCoefficient(rate: number): number | null {
	const life = 1 / rate;
	const invalid = [
		[0, 1],
		[1, 2],
		[2, 3],
		[4, 5],
	].some(([lo, hi]) => life > lo && life < hi);
	if (invalid) {
		return null;
	}
	if (life < 3) {
		return 1;
	}
	if (life < 5) {
		return 1.5;
	}
	return life <= 6 ? 2 : 2.5;
}

function amordegrcImpl(args: FunctionArg[], ctx: FunctionContext) {
	const a = readAmorArgs(args, ctx);
	if (isCellError(a)) {
		return a;
	}
	const coefficient = degressiveCoefficient(a.rate);
	if (coefficient === null) {
		return ERRORS.NUM;
	}
	const rate = a.rate * coefficient;
	let dep = Math.round(yearFrac(a.purchased, a.firstPeriod, a.basis, ctx.date1904) * rate * a.cost);
	let value = a.cost - dep;
	let rest = value - a.salvage;
	for (let n = 0; n < a.period; n++) {
		dep = Math.round(rate * value);
		rest -= dep;
		if (rest < 0) {
			// The last two periods write off half of the remaining value each.
			return a.period - n <= 1 ? Math.round(value * 0.5) : 0;
		}
		value -= dep;
	}
	return dep;
}

export const AMORTIZATION_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'AMORDEGRC', minArgs: 6, maxArgs: 7, params: ['value'], impl: amordegrcImpl },
	{ name: 'AMORLINC', minArgs: 6, maxArgs: 7, params: ['value'], impl: amorlincImpl },
];
