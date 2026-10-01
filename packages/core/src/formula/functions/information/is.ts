import { ERRORS, errorTypeNumber, isCellError } from '../../../values/errors';
import { toNumber } from '../../coerce';
import type { FormulaValue, FunctionArg, FunctionDefinition, Scalar } from '../../types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue, makeArray } from '../../types';
import { isMissingArg } from '../math/args';
import { XLFN } from '../math/define';

/** Normalise an explicitly empty argument to blank. */
function scalarOf(arg: FunctionArg | undefined): Scalar {
	return arg === undefined || isMissingArg(arg) ? null : (arg as Scalar);
}

/**
 * An IS* predicate. Errors are inspected rather than propagated, and ranges
 * or arrays lift the test element-wise.
 */
function predicate(
	name: string,
	test: (v: Scalar) => boolean,
	filePrefix?: typeof XLFN,
): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 1,
		params: ['value'],
		propagateErrors: false,
		...(filePrefix ? { filePrefix } : {}),
		impl: (args) => test(scalarOf(args[0])),
	};
}

/**
 * ISEVEN / ISODD (Analysis ToolPak heritage): errors are returned, booleans
 * and non-numeric text are #VALUE!, blanks count as 0.
 */
function parity(name: string, odd: boolean): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 1,
		params: ['value'],
		propagateErrors: false,
		impl: (args, ctx) => {
			const v = scalarOf(args[0]);
			if (typeof v === 'boolean') {
				return ERRORS.VALUE;
			}
			const n = toNumber(v, ctx.date1904);
			if (typeof n !== 'number') {
				return n;
			}
			return Math.abs(Math.trunc(n) % 2) === (odd ? 1 : 0);
		},
	};
}

/** ISFORMULA(reference): one result per cell of the reference. */
const ISFORMULA: FunctionDefinition = {
	name: 'ISFORMULA',
	minArgs: 1,
	maxArgs: 1,
	params: ['ref'],
	propagateErrors: false,
	filePrefix: XLFN,
	impl: (args, ctx) => {
		let ref = args[0];
		if (isMultiRefValue(ref)) {
			if (ref.refs.length !== 1) {
				return ERRORS.VALUE;
			}
			ref = ref.refs[0];
		}
		if (!isRefValue(ref)) {
			return isCellError(ref) ? ref : ERRORS.VALUE;
		}
		const r = ref.range;
		const has = (row: number, col: number) => ctx.getFormula(ref.sheet, row, col) !== undefined;
		if (r.top === r.bottom && r.left === r.right) {
			return has(r.top, r.left);
		}
		const resolved = ctx.resolveRef(ref);
		return makeArray(
			resolved.values.map((line, i) => line.map((_, j) => has(r.top + i, r.left + j))),
		);
	},
};

/** TYPE: 1 number (and blank), 2 text, 4 logical, 16 error, 64 array, 128 compound (LAMBDA). */
function typeCode(value: FormulaValue): number {
	if (isLambdaValue(value)) {
		return 128;
	}
	if (isArrayValue(value) || isMultiRefValue(value)) {
		return 64;
	}
	if (typeof value === 'string') {
		return 2;
	}
	if (typeof value === 'boolean') {
		return 4;
	}
	if (isCellError(value)) {
		return 16;
	}
	return 1;
}

const TYPE: FunctionDefinition = {
	name: 'TYPE',
	minArgs: 1,
	maxArgs: 1,
	params: ['any'],
	impl: (args, ctx) => {
		const arg = args[0];
		if (isMissingArg(arg)) {
			return 1;
		}
		if (isRefValue(arg)) {
			const r = arg.range;
			if (r.top === r.bottom && r.left === r.right) {
				return typeCode(ctx.getCellValue(arg.sheet, r.top, r.left));
			}
			return 64;
		}
		return typeCode(arg as FormulaValue);
	},
};

const ERROR_TYPE: FunctionDefinition = {
	name: 'ERROR.TYPE',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	propagateErrors: false,
	impl: (args) => {
		const v = scalarOf(args[0]);
		return isCellError(v) ? errorTypeNumber(v.code) : ERRORS.NA;
	},
};

/** N: numbers pass, TRUE is 1, everything else but errors is 0. */
const N: FunctionDefinition = {
	name: 'N',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	impl: (args) => {
		const v = scalarOf(args[0]);
		if (typeof v === 'number') {
			return v;
		}
		if (typeof v === 'boolean') {
			return v ? 1 : 0;
		}
		return isCellError(v) ? v : 0;
	},
};

const NA: FunctionDefinition = {
	name: 'NA',
	minArgs: 0,
	maxArgs: 0,
	params: [],
	impl: () => ERRORS.NA,
};

export const IS_FUNCTIONS: FunctionDefinition[] = [
	predicate('ISBLANK', (v) => v === null),
	predicate('ISERR', (v) => isCellError(v) && v.code !== '#N/A'),
	predicate('ISERROR', (v) => isCellError(v)),
	predicate('ISNA', (v) => isCellError(v) && v.code === '#N/A'),
	predicate('ISLOGICAL', (v) => typeof v === 'boolean'),
	predicate('ISNUMBER', (v) => typeof v === 'number'),
	predicate('ISTEXT', (v) => typeof v === 'string'),
	predicate('ISNONTEXT', (v) => typeof v !== 'string'),
	parity('ISEVEN', false),
	parity('ISODD', true),
	ISFORMULA,
	TYPE,
	ERROR_TYPE,
	N,
	NA,
];
