/**
 * Math & Trig functions, part 2: trigonometry, matrices, number bases,
 * sequences and random numbers.
 */
import { arg, categoryBuilder, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Math & Trig');

const ANGLE = arg('number', 'The angle in radians.');
const REAL = arg('number', 'Any real number.');
const SQUARE = arg('array', 'A numeric array with equal numbers of rows and columns.');

export const MATH_TRIG_INFO: readonly FunctionInfo[] = [
	fn('ACOS', 'Gives the angle in radians, from 0 to pi, whose cosine is a given number.', [
		arg('number', 'The cosine of the angle, from -1 to 1.'),
	]),
	fn('ACOSH', 'Gives the inverse hyperbolic cosine of a number of 1 or more.', [
		arg('number', 'A real number equal to or greater than 1.'),
	]),
	fn('ACOT', 'Gives the angle in radians, from 0 to pi, whose cotangent is a given number.', [
		arg('number', 'The cotangent of the angle.'),
	]),
	fn('ACOTH', 'Gives the inverse hyperbolic cotangent of a number outside the range -1 to 1.', [
		arg('number', 'A real number whose absolute value is greater than 1.'),
	]),
	fn('ARABIC', 'Reads a Roman numeral, such as "MMXXIV", and returns its value as a number.', [
		arg('text', 'The Roman numeral, such as "MMXXIV".'),
	]),
	fn('ASIN', 'Gives the angle in radians, from -pi/2 to pi/2, whose sine is a given number.', [
		arg('number', 'The sine of the angle, from -1 to 1.'),
	]),
	fn('ASINH', 'Gives the inverse hyperbolic sine of a real number.', [REAL]),
	fn('ATAN', 'Gives the angle in radians, from -pi/2 to pi/2, whose tangent is a given number.', [
		arg('number', 'The tangent of the angle.'),
	]),
	fn(
		'ATAN2',
		'Gives the angle in radians between the x-axis and the line from the origin to a point (x, y).',
		[
			arg('x_num', 'The x-coordinate of the point.'),
			arg('y_num', 'The y-coordinate of the point.'),
		],
	),
	fn('ATANH', 'Gives the inverse hyperbolic tangent of a number between -1 and 1.', [
		arg('number', 'A real number strictly between -1 and 1.'),
	]),
	fn('BASE', 'Writes a number as text in another base, such as binary (2) or hexadecimal (16).', [
		arg('number', 'The non-negative integer to convert.'),
		arg('radix', 'The base to convert to, from 2 to 36.'),
		opt('min_length', 'Minimum length of the result, padded with leading zeros.'),
	]),
	fn('COS', 'Gives the cosine of an angle in radians.', [ANGLE]),
	fn('COSH', 'Gives the hyperbolic cosine of a real number.', [REAL]),
	fn('COT', 'Gives the cotangent of an angle in radians.', [ANGLE]),
	fn('COTH', 'Gives the hyperbolic cotangent of a real number.', [REAL]),
	fn('CSC', 'Gives the cosecant of an angle in radians.', [ANGLE]),
	fn('CSCH', 'Gives the hyperbolic cosecant of a real number.', [REAL]),
	fn(
		'DECIMAL',
		'Returns the value of a number written as text in another base, such as "FF" in base 16.',
		[
			arg('text', 'The text to convert, at most 255 characters.'),
			arg('radix', 'The base of the text, from 2 to 36.'),
		],
	),
	fn('DEGREES', 'Changes an angle measured in radians into degrees.', [
		arg('angle', 'The angle in radians.'),
	]),
	fn('MDETERM', 'Calculates the determinant of a square matrix.', [SQUARE]),
	fn('MINVERSE', 'Calculates the inverse of a square matrix.', [SQUARE]),
	fn('MMULT', 'Multiplies two matrices and returns the resulting array.', [
		arg('array1', 'The first matrix; its column count must equal the row count of array2.'),
		arg('array2', 'The second matrix.'),
	]),
	fn('MUNIT', 'Returns the identity matrix of a given size.', [
		arg('dimension', 'Number of rows and columns of the matrix.'),
	]),
	fn('RADIANS', 'Changes an angle measured in degrees into radians.', [
		arg('angle', 'The angle in degrees.'),
	]),
	fn('RAND', 'Generates a uniformly distributed random number in [0, 1) on every recalculation.'),
	fn('RANDARRAY', 'Returns an array of random numbers, optionally whole numbers within a range.', [
		opt('rows', 'Number of rows to return; 1 if omitted.'),
		opt('columns', 'Number of columns to return; 1 if omitted.'),
		opt('min', 'Smallest number to return; 0 if omitted.'),
		opt('max', 'Largest number to return; 1 if omitted.'),
		opt('whole_number', 'TRUE for whole numbers, FALSE or omitted for decimals.'),
	]),
	fn(
		'RANDBETWEEN',
		'Returns a random whole number between two bounds, new on each recalculation.',
		[
			arg('bottom', 'The smallest integer it can return.'),
			arg('top', 'The largest integer it can return.'),
		],
	),
	fn('ROMAN', 'Writes a number as a Roman numeral, such as 2024 as "MMXXIV".', [
		arg('number', 'The number to convert, from 0 to 3999.'),
		opt('form', 'Style: 0 or omitted classic, 1 to 4 increasingly concise.'),
	]),
	fn('SEC', 'Gives the secant of an angle in radians.', [ANGLE]),
	fn('SECH', 'Gives the hyperbolic secant of a real number.', [REAL]),
	fn('SEQUENCE', 'Generates an array of numbers that grow by a fixed step, such as 10, 20, 30.', [
		arg('rows', 'Number of rows to return.'),
		opt('columns', 'Number of columns to return; 1 if omitted.'),
		opt('start', 'First number of the sequence; 1 if omitted.'),
		opt('step', 'Amount added for each following value; 1 if omitted.'),
	]),
	fn('SERIESSUM', 'Evaluates a power series with given coefficients at a value x.', [
		arg('x', 'The input value of the series.'),
		arg('n', 'The initial power of x.'),
		arg('m', 'The step by which the power grows for each term.'),
		arg('coefficients', 'The coefficients that multiply each successive power of x.'),
	]),
	fn('SIN', 'Gives the sine of an angle in radians.', [ANGLE]),
	fn('SINH', 'Gives the hyperbolic sine of a real number.', [REAL]),
	fn('TAN', 'Gives the tangent of an angle in radians.', [ANGLE]),
	fn('TANH', 'Gives the hyperbolic tangent of a real number.', [REAL]),
];
