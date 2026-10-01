import type { CellError } from '../../../types/cell';
import type { FunctionDefinition } from '../../types';
import { finite, numberArg } from './args';

type NumericResult = number | CellError;

/** Excel 2010+ functions are stored with the `_xlfn.` prefix. */
export const XLFN = '_xlfn.' as const;

/**
 * A one-argument numeric function. The argument is coerced with the usual
 * rules; ranges and arrays are lifted element-wise by the evaluator.
 * Non-finite results become `#NUM!`.
 */
export function unaryMath(
	name: string,
	f: (x: number) => NumericResult,
	filePrefix?: typeof XLFN,
): FunctionDefinition {
	return {
		name,
		minArgs: 1,
		maxArgs: 1,
		params: ['value'],
		...(filePrefix ? { filePrefix } : {}),
		impl: (args, ctx) => {
			const x = numberArg(args[0], ctx);
			return typeof x === 'number' ? finite(f(x)) : x;
		},
	};
}

/** A two-argument numeric function with both arguments required. */
export function binaryMath(
	name: string,
	f: (a: number, b: number) => NumericResult,
	filePrefix?: typeof XLFN,
): FunctionDefinition {
	return {
		name,
		minArgs: 2,
		maxArgs: 2,
		params: ['value', 'value'],
		...(filePrefix ? { filePrefix } : {}),
		impl: (args, ctx) => {
			const a = numberArg(args[0], ctx);
			if (typeof a !== 'number') {
				return a;
			}
			const b = numberArg(args[1], ctx);
			return typeof b === 'number' ? finite(f(a, b)) : b;
		},
	};
}
