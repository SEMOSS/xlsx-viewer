/**
 * Statistical functions, part 1: averages, counts, extremes, spread and
 * shape. Distributions and analysis tools live in the other
 * `statistical-*.ts` files.
 */
import { CRITERIA, CRITERIA_PAIRS, NUMBERS, VALUES } from './common-args';
import { arg, categoryBuilder, opt, series } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Statistical');

const COUNT_VALUES = series(
	'value',
	'The first value, reference or range to count.',
	'More values, references or ranges (up to 255 in total).',
);

export const STATISTICAL_INFO: readonly FunctionInfo[] = [
	fn(
		'AVEDEV',
		'Measures spread as the mean absolute distance of data points from their average.',
		NUMBERS,
	),
	fn(
		'AVERAGE',
		'Calculates the arithmetic mean (average) of the numbers in its arguments.',
		NUMBERS,
	),
	fn('AVERAGEA', 'Averages its arguments, treating text and FALSE as 0 and TRUE as 1.', VALUES),
	fn('AVERAGEIF', 'Averages the cells of a range that satisfy a single condition.', [
		arg('range', 'The cells to test against the condition.'),
		CRITERIA,
		opt('average_range', 'The cells to average; the tested range itself if omitted.'),
	]),
	fn('AVERAGEIFS', 'Averages the cells that satisfy every one of several conditions.', [
		arg('average_range', 'The cells to average.'),
		...CRITERIA_PAIRS,
	]),
	fn('COUNT', 'Counts how many cells or values in the arguments are numbers.', COUNT_VALUES),
	fn('COUNTA', 'Counts the nonblank cells and values in its arguments.', COUNT_VALUES),
	fn('COUNTBLANK', 'Counts how many cells of a range are empty.', [
		arg('range', 'The range in which to count blank cells.'),
	]),
	fn('COUNTIF', 'Counts the cells of a range that satisfy a single condition.', [
		arg('range', 'The cells to test and count.'),
		CRITERIA,
	]),
	fn('COUNTIFS', 'Counts the cells that satisfy every one of several conditions.', CRITERIA_PAIRS),
	fn('DEVSQ', 'Adds up the squared distances of data points from their mean.', NUMBERS),
	fn(
		'GEOMEAN',
		'Calculates the geometric mean: the n-th root of the product of n positive numbers.',
		NUMBERS,
	),
	fn(
		'HARMEAN',
		'Calculates the harmonic mean: n divided by the sum of the reciprocals of n positive numbers.',
		NUMBERS,
	),
	fn(
		'KURT',
		'Measures the kurtosis (tailedness) of a data set compared with a normal distribution.',
		NUMBERS,
	),
	fn('LARGE', 'Picks the k-th highest value of a data set, such as the third largest sale.', [
		arg('array', 'The array or range of data.'),
		arg('k', 'Position counted from the largest; 1 returns the maximum.'),
	]),
	fn(
		'MAX',
		'Finds the highest number among its arguments; text and logical values in ranges are skipped.',
		NUMBERS,
	),
	fn(
		'MAXA',
		'Finds the highest value among its arguments, treating text and FALSE as 0 and TRUE as 1.',
		VALUES,
	),
	fn(
		'MAXIFS',
		'Finds the highest value among the cells that satisfy every one of several conditions.',
		[arg('max_range', 'The cells to take the maximum of.'), ...CRITERIA_PAIRS],
	),
	fn('MEDIAN', 'Finds the middle value of a set of numbers once they are sorted.', NUMBERS),
	fn(
		'MIN',
		'Finds the lowest number among its arguments; text and logical values in ranges are skipped.',
		NUMBERS,
	),
	fn(
		'MINA',
		'Finds the lowest value among its arguments, treating text and FALSE as 0 and TRUE as 1.',
		VALUES,
	),
	fn(
		'MINIFS',
		'Finds the lowest value among the cells that satisfy every one of several conditions.',
		[arg('min_range', 'The cells to take the minimum of.'), ...CRITERIA_PAIRS],
	),
	fn(
		'MODE.MULT',
		'Lists, in a vertical array, every value tied for most frequent in a data set.',
		NUMBERS,
	),
	fn('MODE.SNGL', 'Finds the value that occurs most often in a data set.', NUMBERS),
	fn(
		'SKEW',
		'Measures how lopsided a sample distribution is around its mean (sample skewness).',
		NUMBERS,
	),
	fn(
		'SKEW.P',
		'Measures the asymmetry of a distribution around its mean, for a whole population.',
		NUMBERS,
	),
	fn('SMALL', 'Picks the k-th lowest value of a data set, such as the second smallest time.', [
		arg('array', 'The array or range of data.'),
		arg('k', 'Position counted from the smallest; 1 returns the minimum.'),
	]),
	fn(
		'STDEV.P',
		'Calculates the population standard deviation (n in the denominator) of numeric values.',
		NUMBERS,
	),
	fn(
		'STDEV.S',
		'Calculates the sample standard deviation (n - 1 in the denominator) of numeric values.',
		NUMBERS,
	),
	fn(
		'STDEVA',
		'Calculates the sample standard deviation, treating text and FALSE as 0 and TRUE as 1.',
		VALUES,
	),
	fn(
		'STDEVPA',
		'Calculates the population standard deviation, treating text and FALSE as 0 and TRUE as 1.',
		VALUES,
	),
	fn(
		'TRIMMEAN',
		'Returns the mean of a data set after excluding a share of points from both ends.',
		[
			arg('array', 'The values to trim and average.'),
			arg('percent', 'Fraction of points to exclude, split between both ends (0.2 drops 20%).'),
		],
	),
	fn(
		'VAR.P',
		'Calculates the population variance (n in the denominator) of numeric values.',
		NUMBERS,
	),
	fn(
		'VAR.S',
		'Calculates the sample variance (n - 1 in the denominator) of numeric values.',
		NUMBERS,
	),
	fn('VARA', 'Calculates the sample variance, treating text and FALSE as 0 and TRUE as 1.', VALUES),
	fn(
		'VARPA',
		'Calculates the population variance, treating text and FALSE as 0 and TRUE as 1.',
		VALUES,
	),
];
