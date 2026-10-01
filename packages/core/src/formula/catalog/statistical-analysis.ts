/**
 * Statistical functions, part 3: correlation, regression, forecasting,
 * percentiles, ranks and hypothesis tests.
 */
import {
	ARRAY_PAIR,
	CONFIDENCE_ARGS,
	KNOWN_YX,
	PERCENTILE_INC_ARGS,
	PERCENTRANK_ARGS,
	RANK_ARGS,
	T_TEST_ARGS,
	Z_TEST_ARGS,
} from './common-args';
import { arg, categoryBuilder, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Statistical');

const TARGET_DATE = arg('target_date', 'The date or period to predict a value for.');
const ETS_VALUES = arg('values', 'The historical values to forecast from.');
const TIMELINE = arg('timeline', 'Dates or periods matching the values, at a constant step.');
const SEASONALITY = opt(
	'seasonality',
	'Season length: 1 or omitted detects it, 0 means no seasonality.',
);
const DATA_COMPLETION = opt(
	'data_completion',
	'1 or omitted fills missing points by interpolation, 0 uses zeros.',
);
const AGGREGATION = opt(
	'aggregation',
	'How values on the same date combine: 1 or omitted AVERAGE, up to 7 SUM.',
);
const KNOWN_X_OPT = opt("known_x's", 'Known x values; 1, 2, 3, ... if omitted.');
const NEW_X_OPT = opt("new_x's", "X values to predict y values for; known_x's if omitted.");
const STATS = opt(
	'stats',
	'TRUE also returns regression statistics; FALSE or omitted only the coefficients.',
);
const CONST_LINEAR = opt('const', 'TRUE or omitted calculates b normally; FALSE forces b to 0.');
const CONST_EXPONENTIAL = opt(
	'const',
	'TRUE or omitted calculates b normally; FALSE forces b to 1.',
);
const LINEAR_Y = arg("known_y's", 'Known y values of the relationship y = mx + b.');
const EXPONENTIAL_Y = arg("known_y's", 'Known y values of the relationship y = b * m^x.');
const CHOSEN = [
	arg('number', 'Total number of objects.'),
	arg('number_chosen', 'Number of objects in each arrangement.'),
];

export const ANALYSIS_INFO: readonly FunctionInfo[] = [
	fn(
		'CHISQ.TEST',
		'Runs a chi-squared independence test on observed and expected counts and returns the p-value.',
		[
			arg('actual_range', 'The observed values.'),
			arg(
				'expected_range',
				'The expected values, such as row total times column total over grand total.',
			),
		],
	),
	fn(
		'CONFIDENCE.NORM',
		'Returns the margin of error of a confidence interval for a mean, using a normal distribution.',
		CONFIDENCE_ARGS,
	),
	fn(
		'CONFIDENCE.T',
		'Returns the margin of error of a confidence interval for a mean, using a t-distribution.',
		CONFIDENCE_ARGS,
	),
	fn('CORREL', 'Measures the linear correlation (from -1 to 1) between two data sets.', ARRAY_PAIR),
	fn(
		'COVARIANCE.P',
		'Calculates the population covariance of two data sets (dividing by n).',
		ARRAY_PAIR,
	),
	fn(
		'COVARIANCE.S',
		'Calculates the sample covariance of two data sets (dividing by n - 1).',
		ARRAY_PAIR,
	),
	fn(
		'F.TEST',
		'Runs an F-test and returns the two-tailed p-value that two data sets have equal variances.',
		[arg('array1', 'The first data set.'), arg('array2', 'The second data set.')],
	),
	fn('FISHER', 'Applies the Fisher z-transformation, 0.5 * ln((1 + x) / (1 - x)), to a value.', [
		arg('x', 'The value to transform, strictly between -1 and 1.'),
	]),
	fn(
		'FISHERINV',
		'Undoes the Fisher transformation, turning a z value back into a correlation coefficient.',
		[arg('y', 'The value to transform back.')],
	),
	fn(
		'FORECAST.ETS',
		'Predicts a future value from a time series using exponential triple smoothing (ETS).',
		[TARGET_DATE, ETS_VALUES, TIMELINE, SEASONALITY, DATA_COMPLETION, AGGREGATION],
	),
	fn(
		'FORECAST.ETS.CONFINT',
		'Gives the margin of a confidence interval around a FORECAST.ETS prediction.',
		[
			TARGET_DATE,
			ETS_VALUES,
			TIMELINE,
			opt('confidence_level', 'Confidence level between 0 and 1; 0.95 if omitted.'),
			SEASONALITY,
			DATA_COMPLETION,
			AGGREGATION,
		],
	),
	fn(
		'FORECAST.ETS.SEASONALITY',
		'Detects the season length (in timeline steps) of a time series.',
		[ETS_VALUES, TIMELINE, DATA_COMPLETION, AGGREGATION],
	),
	fn(
		'FORECAST.ETS.STAT',
		'Returns a statistic, such as a smoothing parameter or an error measure, of an ETS forecast.',
		[
			ETS_VALUES,
			TIMELINE,
			arg(
				'statistic_type',
				'Which statistic: 1 alpha, 2 beta, 3 gamma, 4 MASE, 5 SMAPE, 6 MAE, 7 RMSE, 8 step.',
			),
			SEASONALITY,
			DATA_COMPLETION,
			AGGREGATION,
		],
	),
	fn('FORECAST.LINEAR', 'Predicts a value along the linear trend fitted to known x and y values.', [
		arg('x', 'The x value to predict a y value for.'),
		...KNOWN_YX,
	]),
	fn(
		'FREQUENCY',
		'Counts how many values fall into each interval and returns the counts as a vertical array.',
		[
			arg('data_array', 'The values to count.'),
			arg('bins_array', 'Upper bounds of the intervals to group the values into.'),
		],
	),
	fn(
		'GROWTH',
		'Fits an exponential curve to known data and returns predicted y values for new x values.',
		[EXPONENTIAL_Y, KNOWN_X_OPT, NEW_X_OPT, CONST_EXPONENTIAL],
	),
	fn(
		'INTERCEPT',
		'Returns where the linear regression line through known x and y values crosses the y-axis.',
		KNOWN_YX,
	),
	fn(
		'LINEST',
		'Fits a straight line to data by least squares and returns its coefficients and statistics.',
		[LINEAR_Y, KNOWN_X_OPT, CONST_LINEAR, STATS],
	),
	fn('LOGEST', 'Fits an exponential curve to data and returns its coefficients and statistics.', [
		EXPONENTIAL_Y,
		KNOWN_X_OPT,
		CONST_EXPONENTIAL,
		STATS,
	]),
	fn('PEARSON', "Calculates Pearson's r, the product-moment correlation of two data sets.", [
		arg('array1', 'The independent values.'),
		arg('array2', 'The dependent values.'),
	]),
	fn(
		'PERCENTILE.EXC',
		'Finds the value below which a fraction k of the data falls, for k strictly between 0 and 1.',
		[arg('array', 'The data.'), arg('k', 'The percentile, strictly between 0 and 1.')],
	),
	fn(
		'PERCENTILE.INC',
		'Finds the value below which a fraction k of the data falls, for k from 0 to 1 inclusive.',
		PERCENTILE_INC_ARGS,
	),
	fn(
		'PERCENTRANK.EXC',
		'Gives the relative standing of a value in a data set, as a fraction strictly between 0 and 1.',
		PERCENTRANK_ARGS,
	),
	fn(
		'PERCENTRANK.INC',
		'Gives the relative standing of a value in a data set, as a fraction from 0 to 1 inclusive.',
		PERCENTRANK_ARGS,
	),
	fn('PERMUT', 'Returns the number of ordered arrangements of objects chosen from a set.', CHOSEN),
	fn(
		'PERMUTATIONA',
		'Returns the number of ordered arrangements of objects chosen from a set, with repetition.',
		CHOSEN,
	),
	fn(
		'PROB',
		'Adds up the probabilities of the x values that lie between a lower and an upper limit.',
		[
			arg('x_range', 'Values of x that have probabilities attached.'),
			arg('prob_range', 'Probabilities of the values in x_range; they must add up to 1.'),
			arg('lower_limit', 'Lower bound of the values to find the probability of.'),
			opt('upper_limit', 'Upper bound; if omitted, the probability of exactly lower_limit.'),
		],
	),
	fn(
		'QUARTILE.EXC',
		'Finds a quartile of a data set using the exclusive percentile method (PERCENTILE.EXC).',
		[arg('array', 'The data.'), arg('quart', 'Which quartile: 1, 2 or 3.')],
	),
	fn(
		'QUARTILE.INC',
		'Finds a quartile of a data set using the inclusive percentile method (PERCENTILE.INC).',
		[
			arg('array', 'The data.'),
			arg(
				'quart',
				'Which value: 0 minimum, 1 first quartile, 2 median, 3 third quartile, 4 maximum.',
			),
		],
	),
	fn(
		'RANK.AVG',
		'Ranks a number within a list of numbers, giving ties the average of the positions they share.',
		RANK_ARGS,
	),
	fn(
		'RANK.EQ',
		'Ranks a number within a list of numbers, giving ties the best position they share.',
		RANK_ARGS,
	),
	fn(
		'RSQ',
		'Calculates r squared, the coefficient of determination of a linear fit to known data.',
		KNOWN_YX,
	),
	fn(
		'SLOPE',
		'Calculates the gradient of the least-squares line fitted to known x and y values.',
		KNOWN_YX,
	),
	fn(
		'STANDARDIZE',
		'Returns the z-score of a value for a distribution with a given mean and standard deviation.',
		[
			arg('x', 'The value to normalize.'),
			arg('mean', 'Arithmetic mean of the distribution.'),
			arg('standard_dev', 'Standard deviation of the distribution.'),
		],
	),
	fn(
		'STEYX',
		'Measures the typical error of y values predicted by a linear regression on known data.',
		KNOWN_YX,
	),
	fn(
		'T.TEST',
		"Runs Student's t-test on two samples and returns the p-value for equal means.",
		T_TEST_ARGS,
	),
	fn('TREND', 'Returns y values predicted by a least-squares straight line fitted to known data.', [
		LINEAR_Y,
		KNOWN_X_OPT,
		NEW_X_OPT,
		CONST_LINEAR,
	]),
	fn(
		'Z.TEST',
		'Runs a z-test of a sample against a hypothesized mean and returns the one-tailed p-value.',
		Z_TEST_ARGS,
	),
];
