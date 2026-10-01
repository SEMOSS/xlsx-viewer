/**
 * Bitwise functions (Excel 2013): BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT.
 * Operands are integers in 0 .. 2^48 - 1.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { engNum } from './args';

const LIMIT = 2 ** 48;

function operand(arg: FunctionArg, ctx: FunctionContext): number | CellError {
	const n = engNum(arg, ctx);
	if (isCellError(n)) {
		return n;
	}
	return n < 0 || n >= LIMIT || !Number.isInteger(n) ? ERRORS.NUM : n;
}

function logical(op: (a: bigint, b: bigint) => bigint) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const a = operand(args[0], ctx);
		const b = operand(args[1], ctx);
		if (isCellError(a)) {
			return a;
		}
		if (isCellError(b)) {
			return b;
		}
		return Number(op(BigInt(a), BigInt(b)));
	};
}

function shift(direction: 1 | -1) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const n = operand(args[0], ctx);
		const amount = engNum(args[1], ctx);
		if (isCellError(n)) {
			return n;
		}
		if (isCellError(amount)) {
			return amount;
		}
		const k = Math.trunc(amount) * direction;
		if (Math.abs(k) > 53) {
			return ERRORS.NUM;
		}
		const result = k >= 0 ? n * 2 ** k : Math.floor(n / 2 ** -k);
		return result >= LIMIT ? ERRORS.NUM : result;
	};
}

const def = (name: string, impl: FunctionDefinition['impl']): FunctionDefinition => ({
	name,
	minArgs: 2,
	maxArgs: 2,
	params: ['value'],
	filePrefix: '_xlfn.',
	impl,
});

const and = (a: bigint, b: bigint) => a & b;
const or = (a: bigint, b: bigint) => a | b;
const xor = (a: bigint, b: bigint) => a ^ b;

export const BITWISE_FUNCTIONS: FunctionDefinition[] = [
	def('BITAND', logical(and)),
	def('BITOR', logical(or)),
	def('BITXOR', logical(xor)),
	def('BITLSHIFT', shift(1)),
	def('BITRSHIFT', shift(-1)),
];
