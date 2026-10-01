/**
 * Complex-number construction and arithmetic: COMPLEX, IMABS, IMAGINARY,
 * IMREAL, IMARGUMENT, IMCONJUGATE, IMSUM, IMSUB, IMPRODUCT, IMDIV, IMPOWER
 * and IMSQRT.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { forEachScalar } from '../../coerce';
import type { FunctionArg, FunctionContext, FunctionDefinition, Scalar } from '../../types';
import { isOmitted } from '../datetime/args';
import { engNum, engNumbers } from './args';
import {
	add,
	arg as argument,
	abs,
	commonSuffix,
	cx,
	div,
	formatComplex,
	mul,
	parseComplex,
	pow,
	sqrt,
	sub,
	type Complex,
	type Cx,
} from './complex';

/** Parse the first `count` scalar arguments as complex numbers. */
export function readComplexes(args: readonly FunctionArg[], count: number): Complex[] | CellError {
	const out: Complex[] = [];
	for (let i = 0; i < count; i++) {
		const z = parseComplex(isOmitted(args[i]) ? null : (args[i] as Scalar));
		if (isCellError(z)) {
			return z;
		}
		out.push(z);
	}
	return out;
}

/** Wrap an operation on parsed operands into an implementation that formats the result. */
export function complexResult(count: number, op: (z: Complex[]) => Cx | CellError) {
	return (args: FunctionArg[]) => {
		const zs = readComplexes(args, count);
		if (isCellError(zs)) {
			return zs;
		}
		const suffix = commonSuffix(zs);
		if (isCellError(suffix)) {
			return suffix;
		}
		const r = op(zs);
		return isCellError(r) ? r : formatComplex(r.re, r.im, suffix);
	};
}

/** A real-valued result of one complex argument. */
function realResult(op: (z: Complex) => number | CellError) {
	return (args: FunctionArg[]) => {
		const zs = readComplexes(args, 1);
		return isCellError(zs) ? zs : op(zs[0]);
	};
}

function complexImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = engNumbers(args, ctx, [null, null]);
	if (isCellError(n)) {
		return n;
	}
	let suffix: string = 'i';
	if (!isOmitted(args[2])) {
		const s = args[2];
		if (isCellError(s)) {
			return s;
		}
		if (typeof s !== 'string' || (s !== '' && s !== 'i' && s !== 'j')) {
			return ERRORS.VALUE;
		}
		suffix = s === '' ? 'i' : s;
	}
	return formatComplex(n[0], n[1], suffix as 'i' | 'j');
}

/** IMSUM / IMPRODUCT over any number of values, ranges and arrays (blank cells are skipped). */
function aggregate(seed: Cx, op: (a: Cx, b: Cx) => Cx) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const zs: Complex[] = [];
		let error: CellError | undefined;
		for (const a of args) {
			if (isOmitted(a)) {
				continue;
			}
			forEachScalar(a, ctx, (v, fromRef) => {
				if (error || (fromRef && v === null)) {
					return;
				}
				const z = parseComplex(v);
				if (isCellError(z)) {
					error = z;
				} else {
					zs.push(z);
				}
			});
		}
		if (error) {
			return error;
		}
		const suffix = commonSuffix(zs);
		if (isCellError(suffix)) {
			return suffix;
		}
		const r = zs.reduce<Cx>((acc, z) => op(acc, z), seed);
		return formatComplex(r.re, r.im, suffix);
	};
}

function powerImpl(args: FunctionArg[], ctx: FunctionContext) {
	const n = engNum(args[1], ctx);
	if (isCellError(n)) {
		return n;
	}
	return complexResult(1, ([z]) => pow(z, n))(args);
}

const value = (
	name: string,
	count: number,
	impl: FunctionDefinition['impl'],
	extra = 0,
): FunctionDefinition => ({
	name,
	minArgs: count,
	maxArgs: count + extra,
	params: ['value'],
	impl,
});

const variadic = (name: string, impl: FunctionDefinition['impl']): FunctionDefinition => ({
	name,
	minArgs: 1,
	maxArgs: 255,
	params: ['any'],
	impl,
});

export const COMPLEX_FUNCTIONS: FunctionDefinition[] = [
	value('COMPLEX', 2, complexImpl, 1),
	value(
		'IMABS',
		1,
		realResult((z) => abs(z)),
	),
	value(
		'IMAGINARY',
		1,
		realResult((z) => z.im),
	),
	value(
		'IMREAL',
		1,
		realResult((z) => z.re),
	),
	value(
		'IMARGUMENT',
		1,
		realResult((z) => (z.re === 0 && z.im === 0 ? ERRORS.DIV0 : argument(z))),
	),
	value(
		'IMCONJUGATE',
		1,
		complexResult(1, ([z]) => cx(z.re, -z.im)),
	),
	value(
		'IMSUB',
		2,
		complexResult(2, ([a, b]) => sub(a, b)),
	),
	value(
		'IMDIV',
		2,
		complexResult(2, ([a, b]) => (b.re === 0 && b.im === 0 ? ERRORS.NUM : div(a, b))),
	),
	value('IMPOWER', 2, powerImpl),
	value(
		'IMSQRT',
		1,
		complexResult(1, ([z]) => sqrt(z)),
	),
	variadic('IMSUM', aggregate(cx(0), add)),
	variadic('IMPRODUCT', aggregate(cx(1), mul)),
];
