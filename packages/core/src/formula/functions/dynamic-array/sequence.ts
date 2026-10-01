import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { optionalBoolean, optionalInteger, optionalNumber } from '../lookup/args';
import { buildArray, checkDimensions } from './shared';

/**
 * SEQUENCE(rows, [columns], [start], [step])
 *
 * Row-major run of numbers `start + k * step`. Omitted arguments default to
 * 1; sizes are truncated. Zero rows or columns is `#CALC!` (empty array),
 * negative sizes `#VALUE!`.
 */
export const SEQUENCE: FunctionDefinition = {
	name: 'SEQUENCE',
	minArgs: 1,
	maxArgs: 4,
	params: ['value'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => {
		const rows = optionalInteger(args, 0, 1, ctx.date1904);
		if (isCellError(rows)) {
			return rows;
		}
		const cols = optionalInteger(args, 1, 1, ctx.date1904);
		if (isCellError(cols)) {
			return cols;
		}
		const start = optionalNumber(args, 2, 1, ctx.date1904);
		if (isCellError(start)) {
			return start;
		}
		const step = optionalNumber(args, 3, 1, ctx.date1904);
		if (isCellError(step)) {
			return step;
		}
		const bad = checkDimensions(rows, cols);
		if (bad) {
			return bad;
		}
		return buildArray(rows, cols, (r, c) => start + (r * cols + c) * step);
	},
};

/**
 * RANDARRAY([rows], [columns], [min], [max], [integer])
 *
 * Random numbers in [min, max) (defaults 0 and 1), or whole numbers in
 * [min, max] when integer is TRUE (min and max must then be whole numbers).
 * min greater than max is `#VALUE!`. Volatile.
 */
export const RANDARRAY: FunctionDefinition = {
	name: 'RANDARRAY',
	minArgs: 0,
	maxArgs: 5,
	params: ['value'],
	volatile: true,
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => {
		const rows = optionalInteger(args, 0, 1, ctx.date1904);
		if (isCellError(rows)) {
			return rows;
		}
		const cols = optionalInteger(args, 1, 1, ctx.date1904);
		if (isCellError(cols)) {
			return cols;
		}
		const min = optionalNumber(args, 2, 0, ctx.date1904);
		if (isCellError(min)) {
			return min;
		}
		const max = optionalNumber(args, 3, 1, ctx.date1904);
		if (isCellError(max)) {
			return max;
		}
		const integer = optionalBoolean(args, 4, false);
		if (isCellError(integer)) {
			return integer;
		}
		const bad = checkDimensions(rows, cols);
		if (bad) {
			return bad;
		}
		if (min > max) {
			return ERRORS.VALUE;
		}
		if (integer) {
			if (!Number.isInteger(min) || !Number.isInteger(max)) {
				return ERRORS.VALUE;
			}
			const span = max - min + 1;
			return buildArray(
				rows,
				cols,
				() => min + Math.min(span - 1, Math.floor(ctx.random() * span)),
			);
		}
		const span = max - min;
		return buildArray(rows, cols, () => min + ctx.random() * span);
	},
};
