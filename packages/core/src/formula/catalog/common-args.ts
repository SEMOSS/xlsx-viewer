/**
 * Argument help shared by several catalog files: day count basis, number and
 * value lists, criteria pairs and probability distribution parameters.
 */
import { arg, more, opt, series } from './define';
import type { FunctionArgInfo } from './types';

/** Day count basis of the securities functions and YEARFRAC. */
export const BASIS = opt(
	'basis',
	'Day count: 0 or omitted US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 EU 30/360.',
);

/** `number1, [number2], ...` of the numeric aggregates (text and logical values are skipped). */
export const NUMBERS = series(
	'number',
	'The first number, range or array.',
	'More numbers, ranges or arrays (up to 255 in total).',
);

/** `value1, [value2], ...` of the A-suffixed aggregates (text counts as 0, TRUE as 1). */
export const VALUES = series(
	'value',
	'The first value, range or array.',
	'More values, ranges or arrays (up to 255 in total).',
);

/** Single condition of SUMIF, COUNTIF and AVERAGEIF. */
export const CRITERIA = arg('criteria', 'The condition, such as 32, ">32", "apples" or B4.');

/** `criteria_range1, criteria1, [criteria_range2, criteria2], ...` of SUMIFS and friends. */
export const CRITERIA_PAIRS: readonly FunctionArgInfo[] = [
	arg('criteria_range1', 'The first range to test.'),
	arg('criteria1', 'The condition for criteria_range1, such as ">32" or "apples".'),
	more('criteria_range2', 'Another range to test.'),
	more('criteria2', 'The condition for that range.'),
];

/** `cumulative` flag of continuous distributions. */
export const CUMULATIVE = arg(
	'cumulative',
	'TRUE for the cumulative distribution function, FALSE for the density function.',
);

/** `cumulative` flag of discrete distributions. */
export const CUMULATIVE_MASS = arg(
	'cumulative',
	'TRUE for the cumulative probability, FALSE for the probability mass function.',
);

export const PROBABILITY = arg('probability', 'A probability for the distribution, from 0 to 1.');
export const MEAN = arg('mean', 'Arithmetic mean of the distribution.');
export const STANDARD_DEV = arg(
	'standard_dev',
	'Standard deviation of the distribution (positive).',
);
export const DEG_FREEDOM = arg('deg_freedom', 'Number of degrees of freedom.');
export const DEG_FREEDOM1 = arg('deg_freedom1', 'Numerator degrees of freedom.');
export const DEG_FREEDOM2 = arg('deg_freedom2', 'Denominator degrees of freedom.');
export const X_VALUE = arg('x', 'The value at which to evaluate the distribution.');
export const X_NON_NEGATIVE = arg(
	'x',
	'The value at which to evaluate the distribution (0 or more).',
);
export const ALPHA = arg('alpha', 'Shape parameter of the distribution (positive).');
export const BETA = arg('beta', 'Scale parameter of the distribution (positive).');

/** Interval bounds of the beta distribution. */
export const BETA_BOUNDS: readonly FunctionArgInfo[] = [
	opt('A', 'Lower bound of the interval of x; 0 if omitted.'),
	opt('B', 'Upper bound of the interval of x; 1 if omitted.'),
];

/** Arguments of the binomial-family functions. */
export const TRIALS = arg('trials', 'Number of independent trials.');
export const PROBABILITY_S = arg('probability_s', 'Probability of success on each trial.');
export const NUMBER_S = arg('number_s', 'Number of successes.');
export const NUMBER_F = arg('number_f', 'Number of failures.');
export const CRITERION_ALPHA = arg('alpha', 'The criterion probability, from 0 to 1.');

/** Arguments of the hypergeometric distribution. */
export const HYPGEOM_ARGS: readonly FunctionArgInfo[] = [
	arg('sample_s', 'Number of successes in the sample.'),
	arg('number_sample', 'Size of the sample.'),
	arg('population_s', 'Number of successes in the population.'),
	arg('number_pop', 'Size of the population.'),
];

/** Paired data sets of the correlation and covariance functions. */
export const ARRAY_PAIR: readonly FunctionArgInfo[] = [
	arg('array1', 'The first range or array of values.'),
	arg('array2', 'The second range or array of values, the same size as array1.'),
];

/** Known data of the linear regression functions. */
export const KNOWN_YX: readonly FunctionArgInfo[] = [
	arg("known_y's", 'The dependent data points.'),
	arg("known_x's", 'The independent data points.'),
];

/** Arguments of the t-test functions. */
export const T_TEST_ARGS: readonly FunctionArgInfo[] = [
	arg('array1', 'The first data set.'),
	arg('array2', 'The second data set.'),
	arg('tails', 'Number of distribution tails: 1 or 2.'),
	arg('type', 'Kind of test: 1 paired, 2 equal variance, 3 unequal variance.'),
];

/** Arguments of the z-test functions. */
export const Z_TEST_ARGS: readonly FunctionArgInfo[] = [
	arg('array', 'The data to test against x.'),
	arg('x', 'The value to test.'),
	opt('sigma', 'Known population standard deviation; the sample standard deviation if omitted.'),
];

/** Arguments of the confidence interval functions. */
export const CONFIDENCE_ARGS: readonly FunctionArgInfo[] = [
	arg('alpha', 'Significance level; 0.05 gives a 95% confidence level.'),
	arg('standard_dev', 'Population standard deviation of the data.'),
	arg('size', 'Sample size.'),
];

/** `array, k` of the percentile functions. */
export const PERCENTILE_INC_ARGS: readonly FunctionArgInfo[] = [
	arg('array', 'The data.'),
	arg('k', 'The percentile, from 0 to 1 inclusive.'),
];

/** `array, x, [significance]` of the percent rank functions. */
export const PERCENTRANK_ARGS: readonly FunctionArgInfo[] = [
	arg('array', 'The data.'),
	arg('x', 'The value to rank.'),
	opt('significance', 'Number of significant digits in the result; 3 if omitted.'),
];

/** `number, ref, [order]` of the rank functions. */
export const RANK_ARGS: readonly FunctionArgInfo[] = [
	arg('number', 'The number to rank.'),
	arg('ref', 'The list of numbers to rank against.'),
	opt('order', '0 or omitted ranks in descending order; any other value ranks ascending.'),
];
