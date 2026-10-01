import { ERRORS, isCellError } from '../../../values/errors';
import type { FormulaValue, FunctionArg, FunctionDefinition, Scalar } from '../../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../../types';
import { isMissingNode } from '../logical/lift';
import { clampToUsed, singleArea } from './shared';

/**
 * ROW / COLUMN return the position of the formula cell or of a reference
 * (a vertical / horizontal array for multi-row / multi-column references;
 * whole columns and rows are clamped to the used range). ROWS / COLUMNS
 * measure references and arrays; AREAS counts the areas of a reference.
 */

function position(name: string, vertical: boolean): FunctionDefinition {
	return {
		name,
		minArgs: 0,
		maxArgs: 1,
		params: ['ref'],
		impl: (args, ctx) => {
			if (args.length === 0 || isMissingNode(args[0])) {
				return (vertical ? ctx.row : ctx.col) + 1;
			}
			const ref = singleArea(args[0]);
			if (isCellError(ref)) {
				return ref;
			}
			const r = clampToUsed(ref, ctx);
			const first = vertical ? r.top : r.left;
			const last = vertical ? r.bottom : r.right;
			if (first === last) {
				return first + 1;
			}
			const numbers: Scalar[] = [];
			for (let i = first; i <= last; i++) {
				numbers.push(i + 1);
			}
			return vertical
				? { kind: 'array', rows: numbers.length, cols: 1, values: numbers.map((n) => [n]) }
				: { kind: 'array', rows: 1, cols: numbers.length, values: [numbers] };
		},
	};
}

function extent(arg: FunctionArg, vertical: boolean): FormulaValue {
	if (isRefValue(arg) || isMultiRefValue(arg)) {
		const ref = singleArea(arg);
		if (isCellError(ref)) {
			return ref;
		}
		const r = ref.range;
		return vertical ? r.bottom - r.top + 1 : r.right - r.left + 1;
	}
	if (isArrayValue(arg)) {
		return vertical ? arg.rows : arg.cols;
	}
	if (isCellError(arg)) {
		return arg;
	}
	if (isLambdaValue(arg) || isMissingNode(arg)) {
		return ERRORS.VALUE;
	}
	return 1;
}

function size(name: string, vertical: boolean): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 1,
		params: ['any'],
		impl: (args) => extent(args[0], vertical),
	};
}

const AREAS: FunctionDefinition = {
	name: 'AREAS',
	minArgs: 1,
	maxArgs: 1,
	params: ['ref'],
	impl: (args) => {
		const arg = args[0];
		if (isRefValue(arg)) {
			return 1;
		}
		if (isMultiRefValue(arg)) {
			return arg.refs.length;
		}
		return isCellError(arg) ? arg : ERRORS.VALUE;
	},
};

export const POSITION_FUNCTIONS: FunctionDefinition[] = [
	position('ROW', true),
	position('COLUMN', false),
	size('ROWS', true),
	size('COLUMNS', false),
	AREAS,
];
