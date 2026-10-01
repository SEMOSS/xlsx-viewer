import { ERRORS, isCellError } from '../../../../values/errors';
import { checkNumber } from '../../../coerce';
import type { ArrayValue, FunctionDefinition, Scalar } from '../../../types';
import { makeArray } from '../../../types';
import { coerceArgs, define } from '../args';
import { buildDesign } from './design';
import { type LeastSquaresFit, leastSquares } from './least-squares';

function num(v: number): Scalar {
	return checkNumber(v);
}

/**
 * The LINEST / LOGEST result block: coefficients in reverse column order
 * followed by the constant, then (with stats) the standard errors, r2 and
 * sey, F and df, and the regression and residual sums of squares; unused
 * cells hold #N/A.
 */
export function regressionBlock(
	fit: LeastSquaresFit,
	withConstant: boolean,
	stats: boolean,
	transform: (v: number) => number = (v) => v,
): ArrayValue {
	const k = fit.coefficients.length;
	const width = k + 1;
	const top: Scalar[] = [];
	for (let j = k - 1; j >= 0; j--) {
		top.push(num(transform(fit.coefficients[j])));
	}
	top.push(num(transform(fit.intercept)));
	if (!stats) {
		return makeArray([top]);
	}
	const { df, ssreg, ssresid, rank } = fit;
	const sey = df > 0 ? Math.sqrt(ssresid / df) : NaN;
	const errors: Scalar[] = [];
	for (let j = k - 1; j >= 0; j--) {
		const factor = fit.coefficientFactors[j];
		errors.push(factor === 0 ? 0 : num(sey * Math.sqrt(factor)));
	}
	errors.push(withConstant ? num(sey * Math.sqrt(fit.interceptFactor)) : ERRORS.NA);
	const f = df > 0 && rank > 0 && ssresid > 0 ? ssreg / rank / (ssresid / df) : NaN;
	const pad = (row: Scalar[]): Scalar[] => {
		while (row.length < width) {
			row.push(ERRORS.NA);
		}
		return row;
	};
	return makeArray([
		top,
		errors,
		pad([num(ssreg / (ssreg + ssresid)), num(sey)]),
		pad([num(f), df]),
		pad([num(ssreg), num(ssresid)]),
	]);
}

function regression(name: string, logarithmic: boolean): FunctionDefinition {
	return define({
		name,
		minArgs: 1,
		maxArgs: 4,
		params: ['any', 'any', 'value', 'value'],
		returnsArray: true,
		impl: (args, ctx) => {
			const flags = coerceArgs([{ b: true }, { b: false }], args, ctx, 2);
			if (isCellError(flags)) {
				return flags;
			}
			const design = buildDesign(args[0], args[1], ctx);
			if (isCellError(design)) {
				return design;
			}
			let y = design.y;
			if (logarithmic) {
				if (y.some((v) => v <= 0)) {
					return ERRORS.NUM;
				}
				y = y.map(Math.log);
			}
			const withConstant = flags[0] === 1;
			const fit = leastSquares(design.x, y, withConstant);
			return regressionBlock(fit, withConstant, flags[1] === 1, logarithmic ? Math.exp : undefined);
		},
	});
}

export const LINEST_FUNCTIONS: FunctionDefinition[] = [
	regression('LINEST', false),
	regression('LOGEST', true),
];
