import { ERRORS, isCellError } from '../../../../values/errors';
import { checkNumber } from '../../../coerce';
import type { FunctionDefinition, Scalar } from '../../../types';
import { makeArray } from '../../../types';
import { coerceArgs, define } from '../args';
import { buildDesign, newPoints } from './design';
import { leastSquares } from './least-squares';

/** TREND (linear) and GROWTH (exponential) predictions at new_x's. */
function prediction(name: string, exponential: boolean): FunctionDefinition {
	return define({
		name,
		minArgs: 1,
		maxArgs: 4,
		params: ['any', 'any', 'any', 'value'],
		returnsArray: true,
		impl: (args, ctx) => {
			const flags = coerceArgs([{ b: true }], args, ctx, 3);
			if (isCellError(flags)) {
				return flags;
			}
			const design = buildDesign(args[0], args[1], ctx);
			if (isCellError(design)) {
				return design;
			}
			let y = design.y;
			if (exponential) {
				if (y.some((v) => v <= 0)) {
					return ERRORS.NUM;
				}
				y = y.map(Math.log);
			}
			const points = newPoints(design, args[2], ctx);
			if (isCellError(points)) {
				return points;
			}
			const fit = leastSquares(design.x, y, flags[0] === 1);
			const values: Scalar[] = points.x.map((row) => {
				let v = fit.intercept;
				for (let j = 0; j < row.length; j++) {
					v += fit.coefficients[j] * row[j];
				}
				return checkNumber(exponential ? Math.exp(v) : v);
			});
			const out: Scalar[][] = [];
			for (let r = 0; r < points.rows; r++) {
				out.push(values.slice(r * points.cols, (r + 1) * points.cols));
			}
			return makeArray(out);
		},
	});
}

export const TREND_FUNCTIONS: FunctionDefinition[] = [
	prediction('TREND', false),
	prediction('GROWTH', true),
];
