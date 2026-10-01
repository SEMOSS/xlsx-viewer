/**
 * Math & Trig functions, part 1: arithmetic, rounding, logarithms, sums and
 * aggregates. Trigonometry, matrices, radix conversion and random numbers
 * live in `math-trig.ts`.
 */
import { CRITERIA, CRITERIA_PAIRS } from './common-args';
import { arg, categoryBuilder, more, opt, series } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Math & Trig');

const VALUE_TO_ROUND = arg('number', 'The value to round.');
const NUM_DIGITS = arg(
	'num_digits',
	'Decimal places to keep; negative values round left of the decimal point.',
);
const SIGNIFICANCE = opt('significance', 'The multiple to round to; 1 if omitted.');
const INTEGERS = series(
	'number',
	'The first value; decimals are truncated.',
	'More values (up to 255 in total).',
);
const POSITIVE = arg('number', 'The positive real number to take the logarithm of.');
const ARRAY_X = arg('array_x', 'The first array or range of values.');
const ARRAY_Y = arg('array_y', 'The second array or range of values.');
const SUM_ITEMS = series(
	'number',
	'The first number, range or array.',
	'More numbers, ranges or arrays (up to 255 in total).',
);
const CHOOSE_ARGS = [
	arg('number', 'Total number of items.'),
	arg('number_chosen', 'Number of items in each combination.'),
];

export const MATH_INFO: readonly FunctionInfo[] = [
	fn('ABS', 'Returns the magnitude of a number, dropping any minus sign.', [
		arg('number', 'The real number to take the absolute value of.'),
	]),
	fn(
		'AGGREGATE',
		'Computes one of 19 aggregates (SUM, LARGE, ...), optionally skipping hidden rows and errors.',
		[
			arg(
				'function_num',
				'Which aggregate: 1 AVERAGE to 13 MODE.SNGL, 14 LARGE to 19 QUARTILE.EXC.',
			),
			arg('options', 'What to skip: 0 to 7, for example 5 hidden rows, 6 errors, 7 both.'),
			arg('ref1', 'The first range or array to aggregate.'),
			more('ref2', 'More ranges; for functions 14 to 19 this is the k value.'),
		],
	),
	fn(
		'CEILING.MATH',
		'Rounds a value up to the next higher multiple of a step, with a mode for negative numbers.',
		[
			VALUE_TO_ROUND,
			SIGNIFICANCE,
			opt('mode', 'For negative numbers, a nonzero value rounds away from zero.'),
		],
	),
	fn(
		'CEILING.PRECISE',
		'Rounds a value up to the next higher multiple of a step (1 by default), regardless of sign.',
		[VALUE_TO_ROUND, SIGNIFICANCE],
	),
	fn(
		'COMBIN',
		'Returns the number of ways to choose a group of items from a set, ignoring order.',
		CHOOSE_ARGS,
	),
	fn(
		'COMBINA',
		'Returns the number of combinations of a group of items when items may repeat.',
		CHOOSE_ARGS,
	),
	fn('EVEN', 'Rounds a number away from zero to the nearest even integer.', [VALUE_TO_ROUND]),
	fn('EXP', 'Raises the constant e (about 2.71828) to a given power.', [
		arg('number', 'The exponent applied to the base e.'),
	]),
	fn('FACT', 'Multiplies all whole numbers from 1 up to a given number (the factorial).', [
		arg('number', 'A non-negative number; decimals are truncated.'),
	]),
	fn(
		'FACTDOUBLE',
		'Multiplies every second whole number down from a given number (the double factorial).',
		[arg('number', 'A non-negative number; decimals are truncated.')],
	),
	fn(
		'FLOOR.MATH',
		'Rounds a value down to the next lower multiple of a step, with a mode for negative numbers.',
		[
			VALUE_TO_ROUND,
			SIGNIFICANCE,
			opt('mode', 'For negative numbers, a nonzero value rounds toward zero.'),
		],
	),
	fn(
		'FLOOR.PRECISE',
		'Rounds a value down to the next lower multiple of a step (1 by default), regardless of sign.',
		[VALUE_TO_ROUND, SIGNIFICANCE],
	),
	fn(
		'GCD',
		'Finds the largest integer that divides all the given integers without a remainder.',
		INTEGERS,
	),
	fn('INT', 'Drops the fractional part of a number by rounding down to the nearest integer.', [
		arg('number', 'The real number to round down.'),
	]),
	fn(
		'ISO.CEILING',
		'Same as CEILING.PRECISE: rounds a value up to the next multiple of a step, whatever its sign.',
		[VALUE_TO_ROUND, SIGNIFICANCE],
	),
	fn(
		'LCM',
		'Finds the smallest positive integer that is a multiple of all the given integers.',
		INTEGERS,
	),
	fn('LN', 'Gives the logarithm of a number to base e.', [POSITIVE]),
	fn('LOG', 'Gives the logarithm of a number to a chosen base (10 by default).', [
		POSITIVE,
		opt('base', 'Base of the logarithm; 10 if omitted.'),
	]),
	fn('LOG10', 'Gives the common (base-10) logarithm of a number.', [POSITIVE]),
	fn('MOD', 'Returns the remainder of a division; the result has the same sign as the divisor.', [
		arg('number', 'The number to divide.'),
		arg('divisor', 'The number to divide by.'),
	]),
	fn('MROUND', 'Rounds a number to the nearest multiple of another number.', [
		VALUE_TO_ROUND,
		arg('multiple', 'The multiple to round to.'),
	]),
	fn(
		'MULTINOMIAL',
		'Returns the factorial of a sum of values divided by the product of their factorials.',
		[...series('number', 'The first value.', 'More values (up to 255 in total).')],
	),
	fn('ODD', 'Rounds a number away from zero to the nearest odd integer.', [VALUE_TO_ROUND]),
	fn(
		'PERCENTOF',
		'Returns the share that a subset of values makes up of the total of all values.',
		[
			arg('data_subset', 'The values in the subset.'),
			arg('data_all', 'All the values, including the subset.'),
		],
	),
	fn('PI', 'Gives the mathematical constant pi to 15 significant digits.'),
	fn('POWER', 'Raises a number to a given exponent.', [
		arg('number', 'The base number.'),
		arg('power', 'The exponent to raise the base to.'),
	]),
	fn('PRODUCT', 'Returns the product of all the numbers in its arguments.', SUM_ITEMS),
	fn('QUOTIENT', 'Divides two numbers and keeps only the whole-number part of the result.', [
		arg('numerator', 'The dividend.'),
		arg('denominator', 'The divisor.'),
	]),
	fn('ROUND', 'Rounds a value half away from zero to a chosen number of digits.', [
		VALUE_TO_ROUND,
		NUM_DIGITS,
	]),
	fn('ROUNDDOWN', 'Rounds a value toward zero to a chosen number of digits.', [
		VALUE_TO_ROUND,
		NUM_DIGITS,
	]),
	fn('ROUNDUP', 'Rounds a value away from zero to a chosen number of digits.', [
		VALUE_TO_ROUND,
		NUM_DIGITS,
	]),
	fn('SIGN', 'Tells the sign of a number: 1 when positive, 0 when zero, -1 when negative.', [
		arg('number', 'Any real number.'),
	]),
	fn('SQRT', 'Gives the positive square root of a number.', [
		arg('number', 'The number to take the square root of.'),
	]),
	fn('SQRTPI', 'Gives the square root of the product of a number and pi.', [
		arg('number', 'The number to multiply by pi.'),
	]),
	fn(
		'SUBTOTAL',
		'Returns an aggregate of a list, skipping other subtotals and optionally hidden rows.',
		[
			arg(
				'function_num',
				'Which aggregate: 1 to 11 include hidden rows, 101 to 111 skip them (9 is SUM).',
			),
			arg('ref1', 'The first range to aggregate.'),
			more('ref2', 'More ranges to aggregate (up to 254 in total).'),
		],
	),
	fn(
		'SUM',
		'Adds up the numbers in its arguments, which can be values, ranges or arrays.',
		SUM_ITEMS,
	),
	fn('SUMIF', 'Totals the cells of a range that satisfy a single condition.', [
		arg('range', 'The cells to test against the condition.'),
		CRITERIA,
		opt('sum_range', 'The cells to add; the tested range itself if omitted.'),
	]),
	fn('SUMIFS', 'Totals the cells that satisfy every one of several conditions.', [
		arg('sum_range', 'The cells to add.'),
		...CRITERIA_PAIRS,
	]),
	fn('SUMPRODUCT', 'Multiplies matching items of arrays and returns the sum of those products.', [
		arg('array1', 'The first array or range.'),
		more('array2', 'More arrays of the same size (up to 255 in total).'),
	]),
	fn('SUMSQ', 'Adds up the squares of the numbers in its arguments.', SUM_ITEMS),
	fn('SUMX2MY2', 'Adds up x^2 - y^2 over matching pairs of values from two arrays.', [
		ARRAY_X,
		ARRAY_Y,
	]),
	fn('SUMX2PY2', 'Adds up x^2 + y^2 over matching pairs of values from two arrays.', [
		ARRAY_X,
		ARRAY_Y,
	]),
	fn('SUMXMY2', 'Adds up (x - y)^2 over matching pairs of values from two arrays.', [
		ARRAY_X,
		ARRAY_Y,
	]),
	fn('TRUNC', 'Cuts off the digits of a number beyond a given precision, without rounding.', [
		arg('number', 'The number to truncate.'),
		opt('num_digits', 'Precision of the truncation; 0 if omitted.'),
	]),
];
