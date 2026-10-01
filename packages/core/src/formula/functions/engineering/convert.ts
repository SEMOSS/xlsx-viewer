/**
 * CONVERT(number, from_unit, to_unit): unit conversion with metric and
 * binary prefixes. Unit names are case-sensitive; unknown units or units of
 * different categories are #N/A.
 */
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, textArg } from '../datetime/args';
import { engNum } from './args';
import {
	BINARY_PREFIXES,
	DECIMAL_PREFIXES,
	TEMPERATURES,
	UNITS,
	type UnitDef,
} from './convert-units';

export interface ResolvedUnit {
	/** Unit name without its prefix. */
	base: string;
	def: UnitDef;
	/** Prefix as a power of ten (already multiplied by the unit's dimension). */
	exp10: number;
	/** Binary prefix as a power of two. */
	exp2: number;
}

/** Prefixes sorted longest first so that `da` wins over `d`. */
const DECIMAL = [...DECIMAL_PREFIXES].sort((a, b) => b[0].length - a[0].length);

/** Look a unit up, splitting off a metric or binary prefix when the name is not a unit itself. */
export function resolveUnit(name: string): ResolvedUnit | null {
	const exact = UNITS.get(name);
	if (exact) {
		return { base: name, def: exact, exp10: 0, exp2: 0 };
	}
	for (const [prefix, exp2] of BINARY_PREFIXES) {
		const def = name.startsWith(prefix) ? UNITS.get(name.slice(prefix.length)) : undefined;
		if (def?.prefix && def.category === 'information') {
			return { base: name.slice(prefix.length), def, exp10: 0, exp2 };
		}
	}
	for (const [prefix, exp] of DECIMAL) {
		const def = name.startsWith(prefix) ? UNITS.get(name.slice(prefix.length)) : undefined;
		if (def?.prefix) {
			return { base: name.slice(prefix.length), def, exp10: exp * (def.power ?? 1), exp2: 0 };
		}
	}
	return null;
}

/** Correctly rounded 10^e. */
function pow10(e: number): number {
	return Number(`1e${e}`);
}

/**
 * Multiply by 10^exp10 * 2^exp2. Small negative exponents divide by the
 * (exact) positive power, which keeps results such as 1 in = 2.54 cm exact.
 */
function scale(value: number, exp10: number, exp2: number): number {
	let v = exp10 < 0 && exp10 >= -22 ? value / pow10(-exp10) : value * pow10(exp10);
	if (exp2 !== 0) {
		v *= 2 ** exp2;
	}
	return v;
}

/** Round to Excel's 15 significant digits, removing binary noise (12 in -> 1 ft, not 0.9999999999999998). */
function round15(n: number): number {
	return n === 0 || !Number.isFinite(n) ? n : Number(n.toPrecision(15));
}

/** Convert between two resolved units of the same category. */
export function convertUnits(value: number, from: ResolvedUnit, to: ResolvedUnit): number {
	if (from.def.category === 'temperature') {
		const a = TEMPERATURES.get(from.base)!;
		const b = TEMPERATURES.get(to.base)!;
		const celsius = ((scale(value, from.exp10, 0) - a.offset) * a.num) / a.den;
		return scale((celsius * b.den) / b.num + b.offset, -to.exp10, 0);
	}
	const base = value * from.def.factor;
	return scale(base, from.exp10 - to.exp10, from.exp2 - to.exp2) / to.def.factor;
}

function convertImpl(args: FunctionArg[], ctx: FunctionContext) {
	const value = engNum(args[0], ctx);
	const fromName = textArg(args[1]);
	const toName = textArg(args[2]);
	if (isCellError(value)) {
		return value;
	}
	if (isCellError(fromName)) {
		return fromName;
	}
	if (isCellError(toName)) {
		return toName;
	}
	const from = resolveUnit(fromName);
	const to = resolveUnit(toName);
	if (!from || !to || from.def.category !== to.def.category) {
		return ERRORS.NA;
	}
	return finite(round15(convertUnits(value, from, to)));
}

export const CONVERT_FUNCTIONS: FunctionDefinition[] = [
	{ name: 'CONVERT', minArgs: 3, maxArgs: 3, params: ['value'], impl: convertImpl },
];
