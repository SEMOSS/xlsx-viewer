/**
 * Compatibility functions: the pre-2010 names kept for older workbooks. Each
 * one has a newer, more precisely named replacement.
 */
import {
	ALPHA,
	ARRAY_PAIR,
	BETA,
	BETA_BOUNDS,
	CONFIDENCE_ARGS,
	CRITERION_ALPHA,
	CUMULATIVE,
	CUMULATIVE_MASS,
	DEG_FREEDOM,
	DEG_FREEDOM1,
	DEG_FREEDOM2,
	HYPGEOM_ARGS,
	KNOWN_YX,
	MEAN,
	NUMBER_F,
	NUMBER_S,
	NUMBERS,
	PERCENTILE_INC_ARGS,
	PERCENTRANK_ARGS,
	PROBABILITY,
	PROBABILITY_S,
	RANK_ARGS,
	STANDARD_DEV,
	T_TEST_ARGS,
	TRIALS,
	X_NON_NEGATIVE,
	X_VALUE,
	Z_TEST_ARGS,
} from './common-args';
import { arg, categoryBuilder, series } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Compatibility');

const BETA_SHAPES = [
	arg('alpha', 'First shape parameter of the distribution (positive).'),
	arg('beta', 'Second shape parameter of the distribution (positive).'),
];
const F_ARGS = [DEG_FREEDOM1, DEG_FREEDOM2];
const ROUNDING = [
	arg('number', 'The value to round.'),
	arg('significance', 'The multiple to round to.'),
];
const LN_ARGS = [
	arg('mean', 'Mean of ln(x).'),
	arg('standard_dev', 'Standard deviation of ln(x).'),
];

export const COMPATIBILITY_INFO: readonly FunctionInfo[] = [
	fn('BETADIST', 'Older form of BETA.DIST that gives the cumulative beta distribution.', [
		arg('x', 'The value between A and B at which to evaluate the distribution.'),
		...BETA_SHAPES,
		...BETA_BOUNDS,
	]),
	fn('BETAINV', 'Older form of BETA.INV that inverts the cumulative beta distribution.', [
		PROBABILITY,
		...BETA_SHAPES,
		...BETA_BOUNDS,
	]),
	fn('BINOMDIST', 'Older form of BINOM.DIST for binomial distribution probabilities.', [
		NUMBER_S,
		TRIALS,
		PROBABILITY_S,
		arg('cumulative', 'TRUE for at most number_s successes, FALSE for exactly number_s.'),
	]),
	fn(
		'CEILING',
		'Older form of CEILING.MATH that rounds a number up to a multiple of significance.',
		ROUNDING,
	),
	fn('CHIDIST', 'Older form of CHISQ.DIST.RT for the right-tail chi-squared probability.', [
		X_NON_NEGATIVE,
		DEG_FREEDOM,
	]),
	fn('CHIINV', 'Older form of CHISQ.INV.RT that inverts the right-tail chi-squared probability.', [
		PROBABILITY,
		DEG_FREEDOM,
	]),
	fn('CHITEST', 'Older form of CHISQ.TEST for the chi-squared test of independence.', [
		arg('actual_range', 'The observed values.'),
		arg('expected_range', 'The expected values.'),
	]),
	fn('CONCATENATE', 'Older form of CONCAT that joins text items into one string.', [
		...series(
			'text',
			'The first item to join: text, a number or a cell.',
			'More items to join (up to 255 in total).',
		),
	]),
	fn(
		'CONFIDENCE',
		'Older form of CONFIDENCE.NORM for the margin of a confidence interval for a mean.',
		CONFIDENCE_ARGS,
	),
	fn(
		'COVAR',
		'Older form of COVARIANCE.P for the population covariance of two data sets.',
		ARRAY_PAIR,
	),
	fn('CRITBINOM', 'Older form of BINOM.INV that inverts the cumulative binomial distribution.', [
		TRIALS,
		PROBABILITY_S,
		CRITERION_ALPHA,
	]),
	fn('EXPONDIST', 'Older form of EXPON.DIST for the exponential distribution.', [
		X_NON_NEGATIVE,
		arg('lambda', 'The rate parameter of the distribution (positive).'),
		CUMULATIVE,
	]),
	fn('FDIST', 'Older form of F.DIST.RT for the right-tail F probability.', [
		X_NON_NEGATIVE,
		...F_ARGS,
	]),
	fn('FINV', 'Older form of F.INV.RT that inverts the right-tail F probability.', [
		PROBABILITY,
		...F_ARGS,
	]),
	fn(
		'FLOOR',
		'Older form of FLOOR.MATH that rounds a number down to a multiple of significance.',
		ROUNDING,
	),
	fn('FORECAST', 'Older form of FORECAST.LINEAR that predicts a value on a linear trend.', [
		arg('x', 'The x value to predict a y value for.'),
		...KNOWN_YX,
	]),
	fn('FTEST', 'Older form of F.TEST that compares the variances of two data sets.', [
		arg('array1', 'The first data set.'),
		arg('array2', 'The second data set.'),
	]),
	fn('GAMMADIST', 'Older form of GAMMA.DIST for the gamma distribution.', [
		X_NON_NEGATIVE,
		ALPHA,
		BETA,
		CUMULATIVE,
	]),
	fn('GAMMAINV', 'Older form of GAMMA.INV that inverts the cumulative gamma distribution.', [
		PROBABILITY,
		ALPHA,
		BETA,
	]),
	fn(
		'HYPGEOMDIST',
		'Older form of HYPGEOM.DIST for the hypergeometric probability of a sample.',
		HYPGEOM_ARGS,
	),
	fn('LOGINV', 'Older form of LOGNORM.INV that inverts the cumulative lognormal distribution.', [
		PROBABILITY,
		...LN_ARGS,
	]),
	fn('LOGNORMDIST', 'Older form of LOGNORM.DIST for the cumulative lognormal distribution.', [
		arg('x', 'The positive value at which to evaluate the distribution.'),
		...LN_ARGS,
	]),
	fn('MODE', 'Older form of MODE.SNGL that finds the most frequent value.', NUMBERS),
	fn('NEGBINOMDIST', 'Older form of NEGBINOM.DIST for the negative binomial probability.', [
		NUMBER_F,
		arg('number_s', 'Threshold number of successes.'),
		PROBABILITY_S,
	]),
	fn('NORMDIST', 'Older form of NORM.DIST for the normal distribution.', [
		X_VALUE,
		MEAN,
		STANDARD_DEV,
		CUMULATIVE,
	]),
	fn('NORMINV', 'Older form of NORM.INV that inverts the cumulative normal distribution.', [
		PROBABILITY,
		MEAN,
		STANDARD_DEV,
	]),
	fn('NORMSDIST', 'Older form of NORM.S.DIST for the cumulative standard normal distribution.', [
		arg('z', 'The value at which to evaluate the distribution.'),
	]),
	fn(
		'NORMSINV',
		'Older form of NORM.S.INV that inverts the cumulative standard normal distribution.',
		[PROBABILITY],
	),
	fn(
		'PERCENTILE',
		'Older form of PERCENTILE.INC that finds the k-th percentile of a data set.',
		PERCENTILE_INC_ARGS,
	),
	fn(
		'PERCENTRANK',
		'Older form of PERCENTRANK.INC for the percentage rank of a value in a data set.',
		PERCENTRANK_ARGS,
	),
	fn('POISSON', 'Older form of POISSON.DIST for the Poisson distribution.', [
		arg('x', 'Number of events.'),
		arg('mean', 'Expected number of events.'),
		CUMULATIVE_MASS,
	]),
	fn('QUARTILE', 'Older form of QUARTILE.INC that finds a quartile of a data set.', [
		arg('array', 'The data.'),
		arg(
			'quart',
			'Which value: 0 minimum, 1 first quartile, 2 median, 3 third quartile, 4 maximum.',
		),
	]),
	fn('RANK', 'Older form of RANK.EQ that ranks a number within a list of numbers.', RANK_ARGS),
	fn('STDEV', 'Older form of STDEV.S for the sample standard deviation.', NUMBERS),
	fn('STDEVP', 'Older form of STDEV.P for the population standard deviation.', NUMBERS),
	fn(
		'TDIST',
		"Older form of T.DIST.2T and T.DIST.RT for tail probabilities of Student's t-distribution.",
		[X_NON_NEGATIVE, DEG_FREEDOM, arg('tails', 'Number of distribution tails: 1 or 2.')],
	),
	fn('TINV', "Older form of T.INV.2T that inverts the two-tailed Student's t-distribution.", [
		PROBABILITY,
		DEG_FREEDOM,
	]),
	fn('TTEST', "Older form of T.TEST that returns the p-value of Student's t-test.", T_TEST_ARGS),
	fn('VAR', 'Older form of VAR.S for the sample variance.', NUMBERS),
	fn('VARP', 'Older form of VAR.P for the population variance.', NUMBERS),
	fn('WEIBULL', 'Older form of WEIBULL.DIST for the Weibull distribution.', [
		X_NON_NEGATIVE,
		ALPHA,
		BETA,
		CUMULATIVE,
	]),
	fn('ZTEST', 'Older form of Z.TEST that returns the one-tailed p-value of a z-test.', Z_TEST_ARGS),
];
