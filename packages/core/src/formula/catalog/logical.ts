/** Logical functions, including LET, LAMBDA and the LAMBDA helper functions. */
import { arg, categoryBuilder, more, opt, series } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Logical');

const LOGICALS = series(
	'logical',
	'The first condition, which evaluates to TRUE or FALSE.',
	'More conditions (up to 255 in total).',
);
const INITIAL_VALUE = opt('initial_value', 'Starting value of the accumulator.');
const ACCUMULATOR_LAMBDA = arg(
	'lambda',
	'A LAMBDA with accumulator and value parameters that returns the new accumulator.',
);

export const LOGICAL_INFO: readonly FunctionInfo[] = [
	fn('AND', 'Returns TRUE only when every condition given evaluates to TRUE.', LOGICALS),
	fn(
		'BYCOL',
		'Reduces each column of an array to one value with a LAMBDA, returning a row of results.',
		[
			arg('array', 'The array to process column by column.'),
			arg('lambda', 'A LAMBDA that takes one column and returns a single value.'),
		],
	),
	fn(
		'BYROW',
		'Reduces each row of an array to one value with a LAMBDA, returning a column of results.',
		[
			arg('array', 'The array to process row by row.'),
			arg('lambda', 'A LAMBDA that takes one row and returns a single value.'),
		],
	),
	fn('FALSE', 'Gives the Boolean constant FALSE.'),
	fn('IF', 'Returns one value when a condition is TRUE and another when it is FALSE.', [
		arg('logical_test', 'A value or expression that evaluates to TRUE or FALSE.'),
		opt('value_if_true', 'The result when logical_test is TRUE.'),
		opt('value_if_false', 'The result when logical_test is FALSE; FALSE if omitted.'),
	]),
	fn(
		'IFERROR',
		'Returns the result of an expression, or a fallback value when that result is an error.',
		[
			arg('value', 'The value or expression to check.'),
			arg('value_if_error', 'The result when value is an error.'),
		],
	),
	fn('IFNA', 'Returns the result of an expression, or a fallback value when that result is #N/A.', [
		arg('value', 'The value or expression to check.'),
		arg('value_if_na', 'The result when value is #N/A.'),
	]),
	fn(
		'IFS',
		'Tests conditions in order and returns the result paired with the first one that holds.',
		[
			arg('logical_test1', 'The first condition to test.'),
			arg('value_if_true1', 'The result when logical_test1 is TRUE.'),
			more('logical_test2', 'Another condition, tested when the previous ones are FALSE.'),
			more('value_if_true2', 'The result when that condition is TRUE.'),
		],
	),
	fn(
		'LAMBDA',
		'Defines a custom function from parameter names and a formula; name it to reuse it.',
		[
			more('parameter1', 'A name for a value passed to the function (up to 253 parameters).'),
			arg('calculation', 'The formula that computes the result; always the last argument.'),
		],
	),
	fn('LET', 'Names intermediate values inside a formula so they are calculated once and reused.', [
		arg('name1', 'The first name to assign; it must start with a letter.'),
		arg('name_value1', 'The value assigned to name1.'),
		arg(
			'calculation_or_name2',
			'Either the final calculation that uses the names, or another name.',
		),
		more('name_value2', 'The value assigned to the preceding name.'),
		more('calculation_or_name3', 'Either the final calculation, or another name to assign.'),
	]),
	fn(
		'MAKEARRAY',
		'Builds an array of a chosen size, computing each value with a LAMBDA of row and column.',
		[
			arg('rows', 'Number of rows of the array.'),
			arg('cols', 'Number of columns of the array.'),
			arg('lambda', 'A LAMBDA with row and column parameters that calculates each value.'),
		],
	),
	fn(
		'MAP',
		'Transforms each value of one or more arrays with a LAMBDA and returns the new array.',
		[
			arg('array1', 'The array to map.'),
			more('array2', 'More arrays, mapped in step with the first.'),
			arg('lambda', 'A LAMBDA with one parameter per array; always the last argument.'),
		],
	),
	fn('NOT', 'Reverses a logical value: TRUE becomes FALSE and FALSE becomes TRUE.', [
		arg('logical', 'A value or expression that evaluates to TRUE or FALSE.'),
	]),
	fn('OR', 'Returns TRUE when at least one condition given evaluates to TRUE.', LOGICALS),
	fn(
		'REDUCE',
		'Folds an array into a single result by feeding each value and a running total to a LAMBDA.',
		[INITIAL_VALUE, arg('array', 'The array to reduce.'), ACCUMULATOR_LAMBDA],
	),
	fn('SCAN', 'Like REDUCE, but returns the running total after each value as an array.', [
		INITIAL_VALUE,
		arg('array', 'The array to scan.'),
		ACCUMULATOR_LAMBDA,
	]),
	fn(
		'SWITCH',
		'Matches a value against a list of cases and returns the result of the first matching case.',
		[
			arg('expression', 'The value to compare.'),
			arg('value1', 'The first value to compare against.'),
			arg('result1', 'The result when expression matches value1.'),
			more('default_or_value2', 'Another value to compare against, or a final default result.'),
			more('result2', 'The result when expression matches that value.'),
		],
	),
	fn('TRUE', 'Gives the Boolean constant TRUE.'),
	fn('XOR', 'Returns TRUE if an odd number of its arguments are TRUE (exclusive or).', LOGICALS),
];
