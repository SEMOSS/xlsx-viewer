/** Engineering functions, part 2: complex numbers written as text such as "3+4i". */
import { arg, categoryBuilder, more } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Engineering');

const INUMBER = [arg('inumber', 'A complex number, as text in the form x+yi or x+yj.')];

export const COMPLEX_INFO: readonly FunctionInfo[] = [
	fn('IMABS', 'Gives the modulus (distance from zero) of a complex number.', INUMBER),
	fn('IMAGINARY', 'Extracts the imaginary coefficient of a complex number.', INUMBER),
	fn(
		'IMARGUMENT',
		'Gives the angle, in radians, of a complex number in the complex plane.',
		INUMBER,
	),
	fn(
		'IMCONJUGATE',
		'Flips the sign of the imaginary part of a complex number (its conjugate).',
		INUMBER,
	),
	fn('IMCOS', 'Calculates the cosine of a complex number.', INUMBER),
	fn('IMCOSH', 'Calculates the hyperbolic cosine of a complex number.', INUMBER),
	fn('IMCOT', 'Calculates the cotangent of a complex number.', INUMBER),
	fn('IMCSC', 'Calculates the cosecant of a complex number.', INUMBER),
	fn('IMCSCH', 'Calculates the hyperbolic cosecant of a complex number.', INUMBER),
	fn('IMDIV', 'Divides one complex number by another.', [
		arg('inumber1', 'The complex numerator (dividend).'),
		arg('inumber2', 'The complex denominator (divisor).'),
	]),
	fn('IMEXP', 'Raises e to the power of a complex number.', INUMBER),
	fn('IMLN', 'Calculates the natural (base-e) logarithm of a complex number.', INUMBER),
	fn('IMLOG10', 'Calculates the base-10 logarithm of a complex number.', INUMBER),
	fn('IMLOG2', 'Calculates the base-2 logarithm of a complex number.', INUMBER),
	fn('IMPOWER', 'Raises a complex number to a power.', [
		arg('inumber', 'The complex number to raise.'),
		arg('number', 'The power: an integer, a fraction or a negative number.'),
	]),
	fn('IMPRODUCT', 'Multiplies complex numbers together.', [
		arg('inumber1', 'The first complex number to multiply.'),
		more('inumber2', 'More complex numbers to multiply (up to 255 in total).'),
	]),
	fn('IMREAL', 'Extracts the real coefficient of a complex number.', INUMBER),
	fn('IMSEC', 'Calculates the secant of a complex number.', INUMBER),
	fn('IMSECH', 'Calculates the hyperbolic secant of a complex number.', INUMBER),
	fn('IMSIN', 'Calculates the sine of a complex number.', INUMBER),
	fn('IMSINH', 'Calculates the hyperbolic sine of a complex number.', INUMBER),
	fn('IMSQRT', 'Calculates the square root of a complex number.', INUMBER),
	fn('IMSUB', 'Subtracts one complex number from another.', [
		arg('inumber1', 'The complex number to subtract from.'),
		arg('inumber2', 'The complex number to subtract.'),
	]),
	fn('IMSUM', 'Adds complex numbers together.', [
		arg('inumber1', 'The first complex number to add.'),
		more('inumber2', 'More complex numbers to add (up to 255 in total).'),
	]),
	fn('IMTAN', 'Calculates the tangent of a complex number.', INUMBER),
];
