/** Statistical functions, part 2: probability distributions and their inverses. */
import {
	ALPHA,
	BETA,
	BETA_BOUNDS,
	CRITERION_ALPHA,
	CUMULATIVE,
	CUMULATIVE_MASS,
	DEG_FREEDOM,
	DEG_FREEDOM1,
	DEG_FREEDOM2,
	HYPGEOM_ARGS,
	MEAN,
	NUMBER_F,
	NUMBER_S,
	PROBABILITY,
	PROBABILITY_S,
	STANDARD_DEV,
	TRIALS,
	X_NON_NEGATIVE,
	X_VALUE,
} from './common-args';
import { arg, categoryBuilder, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Statistical');

const BETA_ALPHA = arg('alpha', 'First shape parameter of the distribution (positive).');
const BETA_BETA = arg('beta', 'Second shape parameter of the distribution (positive).');
const LN_MEAN = arg('mean', 'Mean of ln(x).');
const LN_STANDARD_DEV = arg('standard_dev', 'Standard deviation of ln(x).');
const F_ARGS = [DEG_FREEDOM1, DEG_FREEDOM2];

export const DISTRIBUTION_INFO: readonly FunctionInfo[] = [
	fn(
		'BETA.DIST',
		'Returns the beta distribution, often used to model proportions and percentages.',
		[
			arg('x', 'The value between A and B at which to evaluate the distribution.'),
			BETA_ALPHA,
			BETA_BETA,
			CUMULATIVE,
			...BETA_BOUNDS,
		],
	),
	fn(
		'BETA.INV',
		'Finds the value x at which the cumulative beta distribution reaches a probability.',
		[PROBABILITY, BETA_ALPHA, BETA_BETA, ...BETA_BOUNDS],
	),
	fn(
		'BINOM.DIST',
		'Returns the probability of a number of successes in a fixed number of independent trials.',
		[
			NUMBER_S,
			TRIALS,
			PROBABILITY_S,
			arg('cumulative', 'TRUE for at most number_s successes, FALSE for exactly number_s.'),
		],
	),
	fn(
		'BINOM.DIST.RANGE',
		'Returns the probability that the number of successes in trials falls within a range.',
		[
			TRIALS,
			PROBABILITY_S,
			arg('number_s', 'Number of successes, or the lower bound when number_s2 is given.'),
			opt('number_s2', 'Upper bound on the number of successes.'),
		],
	),
	fn(
		'BINOM.INV',
		'Returns the smallest success count whose cumulative binomial probability reaches alpha.',
		[TRIALS, PROBABILITY_S, CRITERION_ALPHA],
	),
	fn(
		'CHISQ.DIST',
		'Evaluates the chi-squared distribution (density or left-tailed cumulative probability).',
		[X_NON_NEGATIVE, DEG_FREEDOM, CUMULATIVE],
	),
	fn('CHISQ.DIST.RT', 'Calculates the area of the chi-squared distribution to the right of x.', [
		X_NON_NEGATIVE,
		DEG_FREEDOM,
	]),
	fn(
		'CHISQ.INV',
		'Finds the value whose left-tail chi-squared probability equals a given probability.',
		[PROBABILITY, DEG_FREEDOM],
	),
	fn(
		'CHISQ.INV.RT',
		'Finds the value whose right-tail chi-squared probability equals a given probability.',
		[PROBABILITY, DEG_FREEDOM],
	),
	fn(
		'EXPON.DIST',
		'Evaluates the exponential distribution, often used to model the time between events.',
		[
			X_NON_NEGATIVE,
			arg('lambda', 'The rate parameter of the distribution (positive).'),
			CUMULATIVE,
		],
	),
	fn('F.DIST', 'Evaluates the F distribution (density or left-tailed cumulative probability).', [
		X_NON_NEGATIVE,
		...F_ARGS,
		CUMULATIVE,
	]),
	fn(
		'F.DIST.RT',
		'Calculates the area of the F distribution to the right of x, used to compare variances.',
		[X_NON_NEGATIVE, ...F_ARGS],
	),
	fn('F.INV', 'Finds the value whose left-tail F probability equals a given probability.', [
		PROBABILITY,
		...F_ARGS,
	]),
	fn('F.INV.RT', 'Finds the value whose right-tail F probability equals a given probability.', [
		PROBABILITY,
		...F_ARGS,
	]),
	fn('GAMMA', 'Evaluates the gamma function, which extends the factorial to real numbers.', [
		arg('number', 'The value to evaluate; not 0 or a negative integer.'),
	]),
	fn(
		'GAMMA.DIST',
		'Evaluates the gamma distribution, used for skewed data such as waiting times.',
		[
			X_NON_NEGATIVE,
			ALPHA,
			arg('beta', 'Scale parameter of the distribution; 1 gives the standard gamma distribution.'),
			CUMULATIVE,
		],
	),
	fn(
		'GAMMA.INV',
		'Finds the value at which the cumulative gamma distribution reaches a probability.',
		[PROBABILITY, ALPHA, BETA],
	),
	fn('GAMMALN', 'Calculates ln(GAMMA(x)), the natural logarithm of the gamma function.', [
		arg('x', 'The positive value to evaluate.'),
	]),
	fn('GAMMALN.PRECISE', 'Calculates ln(GAMMA(x)) with full precision; equivalent to GAMMALN.', [
		arg('x', 'The positive value to evaluate.'),
	]),
	fn('GAUSS', 'Returns the probability that a standard normal value lies between the mean and z.', [
		arg('z', 'The number of standard deviations from the mean.'),
	]),
	fn(
		'HYPGEOM.DIST',
		'Returns the probability of a number of sample successes when drawing without replacement.',
		[
			...HYPGEOM_ARGS,
			arg(
				'cumulative',
				'TRUE for the cumulative probability, FALSE for exactly sample_s successes.',
			),
		],
	),
	fn(
		'LOGNORM.DIST',
		'Evaluates the lognormal distribution, for values whose logarithm is normally distributed.',
		[
			arg('x', 'The positive value at which to evaluate the distribution.'),
			LN_MEAN,
			LN_STANDARD_DEV,
			CUMULATIVE,
		],
	),
	fn(
		'LOGNORM.INV',
		'Finds the value at which the cumulative lognormal distribution reaches a probability.',
		[PROBABILITY, LN_MEAN, LN_STANDARD_DEV],
	),
	fn(
		'NEGBINOM.DIST',
		'Gives the chance of a number of failures occurring before a set number of successes.',
		[NUMBER_F, arg('number_s', 'Threshold number of successes.'), PROBABILITY_S, CUMULATIVE_MASS],
	),
	fn(
		'NORM.DIST',
		'Evaluates the normal distribution (density or cumulative) for a mean and standard deviation.',
		[X_VALUE, MEAN, STANDARD_DEV, CUMULATIVE],
	),
	fn(
		'NORM.INV',
		'Finds the value at which the cumulative normal distribution reaches a probability.',
		[PROBABILITY, MEAN, STANDARD_DEV],
	),
	fn(
		'NORM.S.DIST',
		'Evaluates the normal distribution with mean 0 and standard deviation 1 at a z-score.',
		[arg('z', 'The value at which to evaluate the distribution.'), CUMULATIVE],
	),
	fn(
		'NORM.S.INV',
		'Finds the z-score at which the cumulative standard normal distribution reaches a probability.',
		[PROBABILITY],
	),
	fn('PHI', 'Evaluates the bell curve (standard normal density) at a value.', [
		arg('x', 'The value at which to evaluate the density.'),
	]),
	fn(
		'POISSON.DIST',
		'Evaluates the Poisson distribution, used to model how many events happen in an interval.',
		[arg('x', 'Number of events.'), arg('mean', 'Expected number of events.'), CUMULATIVE_MASS],
	),
	fn(
		'T.DIST',
		"Evaluates Student's t-distribution (density or left-tailed cumulative probability).",
		[X_VALUE, DEG_FREEDOM, CUMULATIVE],
	),
	fn('T.DIST.2T', "Calculates the two-tailed probability of Student's t-distribution beyond x.", [
		X_NON_NEGATIVE,
		DEG_FREEDOM,
	]),
	fn('T.DIST.RT', "Calculates the right-tail probability of Student's t-distribution beyond x.", [
		X_VALUE,
		DEG_FREEDOM,
	]),
	fn('T.INV', 'Finds the t-value whose left-tail probability equals a given probability.', [
		PROBABILITY,
		DEG_FREEDOM,
	]),
	fn('T.INV.2T', 'Finds the t-value whose two-tailed probability equals a given probability.', [
		PROBABILITY,
		DEG_FREEDOM,
	]),
	fn(
		'WEIBULL.DIST',
		'Evaluates the Weibull distribution, often used in reliability and failure analysis.',
		[X_NON_NEGATIVE, ALPHA, BETA, CUMULATIVE],
	),
];
