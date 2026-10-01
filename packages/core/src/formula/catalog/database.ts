/** Database functions: aggregates over the records of a list that match a criteria range. */
import { arg, categoryBuilder } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Database');

const DATABASE_ARGS = [
	arg('database', 'The list of records, with column labels in its first row.'),
	arg(
		'field',
		'The column to use: its label in quotes, such as "Age", or its position (1 = first).',
	),
	arg('criteria', 'The range of conditions: a column label with at least one condition below it.'),
];

export const DATABASE_INFO: readonly FunctionInfo[] = [
	fn(
		'DAVERAGE',
		'Averages one field over the database records that satisfy a criteria range.',
		DATABASE_ARGS,
	),
	fn(
		'DCOUNT',
		'Counts the numeric cells of one field over the records that satisfy a criteria range.',
		DATABASE_ARGS,
	),
	fn(
		'DCOUNTA',
		'Counts the nonblank cells of one field over the records that satisfy a criteria range.',
		DATABASE_ARGS,
	),
	fn(
		'DGET',
		'Returns one field of the single database record that satisfies a criteria range.',
		DATABASE_ARGS,
	),
	fn(
		'DMAX',
		'Finds the highest value of one field over the records that satisfy a criteria range.',
		DATABASE_ARGS,
	),
	fn(
		'DMIN',
		'Finds the lowest value of one field over the records that satisfy a criteria range.',
		DATABASE_ARGS,
	),
	fn(
		'DPRODUCT',
		'Multiplies one field over the database records that satisfy a criteria range.',
		DATABASE_ARGS,
	),
	fn(
		'DSTDEV',
		'Calculates the sample standard deviation of one field over the matching records.',
		DATABASE_ARGS,
	),
	fn(
		'DSTDEVP',
		'Calculates the population standard deviation of one field over the matching records.',
		DATABASE_ARGS,
	),
	fn(
		'DSUM',
		'Totals one field over the database records that satisfy a criteria range.',
		DATABASE_ARGS,
	),
	fn(
		'DVAR',
		'Calculates the sample variance of one field over the matching records.',
		DATABASE_ARGS,
	),
	fn(
		'DVARP',
		'Calculates the population variance of one field over the matching records.',
		DATABASE_ARGS,
	),
];
