/**
 * Engineering functions, part 1: Bessel and error functions, number base
 * conversion, bitwise operations, unit conversion and comparisons. The
 * complex number functions live in `engineering-complex.ts`.
 */
import { arg, categoryBuilder, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Engineering');

const BESSEL = [
	arg('x', 'The value at which to evaluate the function.'),
	arg('n', 'Order of the function; decimals are truncated.'),
];
const PLACES = opt(
	'places',
	'Minimum number of characters of the result, padded with leading zeros.',
);
const BINARY = arg('number', "The binary number: at most 10 bits, negatives in two's complement.");
const OCTAL = arg('number', "The octal number: at most 10 digits, negatives in two's complement.");
const HEX = arg(
	'number',
	"The hexadecimal number: at most 10 digits, negatives in two's complement.",
);
const BIT_PAIR = [
	arg('number1', 'A non-negative integer below 2^48.'),
	arg('number2', 'A non-negative integer below 2^48.'),
];
const BIT_NUMBER = arg('number', 'A non-negative integer below 2^48.');

export const ENGINEERING_INFO: readonly FunctionInfo[] = [
	fn('BESSELI', 'Evaluates the modified Bessel function of the first kind, In(x).', BESSEL),
	fn('BESSELJ', 'Evaluates the Bessel function of the first kind, Jn(x).', BESSEL),
	fn('BESSELK', 'Evaluates the modified Bessel function of the second kind, Kn(x).', BESSEL),
	fn('BESSELY', 'Evaluates the Bessel function of the second kind, Yn(x).', BESSEL),
	fn('BIN2DEC', 'Rewrites a binary (base-2) number as a decimal number.', [BINARY]),
	fn('BIN2HEX', 'Rewrites a binary (base-2) number as a hexadecimal (base-16) number.', [
		BINARY,
		PLACES,
	]),
	fn('BIN2OCT', 'Rewrites a binary (base-2) number as an octal (base-8) number.', [BINARY, PLACES]),
	fn(
		'BITAND',
		'Combines two integers bit by bit, keeping the bits set in both (bitwise AND).',
		BIT_PAIR,
	),
	fn('BITLSHIFT', 'Shifts the bits of an integer to the left, multiplying it by a power of 2.', [
		BIT_NUMBER,
		arg('shift_amount', 'Bits to shift; negative values shift right.'),
	]),
	fn(
		'BITOR',
		'Combines two integers bit by bit, keeping the bits set in either (bitwise OR).',
		BIT_PAIR,
	),
	fn('BITRSHIFT', 'Shifts the bits of an integer to the right, dividing it by a power of 2.', [
		BIT_NUMBER,
		arg('shift_amount', 'Bits to shift; negative values shift left.'),
	]),
	fn(
		'BITXOR',
		'Combines two integers bit by bit, keeping the bits set in exactly one (bitwise XOR).',
		BIT_PAIR,
	),
	fn('COMPLEX', 'Builds a complex number from its real and imaginary coefficients.', [
		arg('real_num', 'The real coefficient.'),
		arg('i_num', 'The imaginary coefficient.'),
		opt('suffix', 'Suffix of the imaginary part: "i" (the default) or "j".'),
	]),
	fn('CONVERT', 'Changes a measurement from one unit to another, such as miles to kilometers.', [
		arg('number', 'The value to convert.'),
		arg('from_unit', 'The unit of number, such as "mi" or "C".'),
		arg('to_unit', 'The unit of the result, such as "km" or "F".'),
	]),
	fn('DEC2BIN', 'Rewrites a decimal number as a binary (base-2) number.', [
		arg('number', 'The decimal integer to convert, from -512 to 511.'),
		PLACES,
	]),
	fn('DEC2HEX', 'Rewrites a decimal number as a hexadecimal (base-16) number.', [
		arg('number', 'The decimal integer to convert, from -2^39 to 2^39 - 1.'),
		PLACES,
	]),
	fn('DEC2OCT', 'Rewrites a decimal number as an octal (base-8) number.', [
		arg('number', 'The decimal integer to convert, from -2^29 to 2^29 - 1.'),
		PLACES,
	]),
	fn('DELTA', 'Returns 1 when two numbers are equal and 0 when they differ (Kronecker delta).', [
		arg('number1', 'The first number.'),
		opt('number2', 'The second number; 0 if omitted.'),
	]),
	fn('ERF', 'Evaluates the error function between two limits, or from 0 to a single limit.', [
		arg(
			'lower_limit',
			'Lower bound of the integral; without upper_limit, the integral runs from 0 to it.',
		),
		opt('upper_limit', 'Upper bound of the integral.'),
	]),
	fn('ERF.PRECISE', 'Evaluates the error function from 0 to a value.', [
		arg('x', 'The upper bound of the integral.'),
	]),
	fn('ERFC', 'Evaluates the complementary error function, 1 - ERF(x).', [
		arg('x', 'The lower bound of the integral.'),
	]),
	fn('ERFC.PRECISE', 'Evaluates the complementary error function, 1 - ERF(x), like ERFC.', [
		arg('x', 'The lower bound of the integral.'),
	]),
	fn(
		'GESTEP',
		'Returns 1 when a number is greater than or equal to a threshold, and 0 otherwise.',
		[arg('number', 'The value to test.'), opt('step', 'The threshold; 0 if omitted.')],
	),
	fn('HEX2BIN', 'Rewrites a hexadecimal (base-16) number as a binary (base-2) number.', [
		HEX,
		PLACES,
	]),
	fn('HEX2DEC', 'Rewrites a hexadecimal (base-16) number as a decimal number.', [HEX]),
	fn('HEX2OCT', 'Rewrites a hexadecimal (base-16) number as an octal (base-8) number.', [
		HEX,
		PLACES,
	]),
	fn('OCT2BIN', 'Rewrites an octal (base-8) number as a binary (base-2) number.', [OCTAL, PLACES]),
	fn('OCT2DEC', 'Rewrites an octal (base-8) number as a decimal number.', [OCTAL]),
	fn('OCT2HEX', 'Rewrites an octal (base-8) number as a hexadecimal (base-16) number.', [
		OCTAL,
		PLACES,
	]),
];
