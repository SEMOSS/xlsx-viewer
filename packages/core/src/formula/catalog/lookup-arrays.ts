/**
 * Lookup & Reference functions, part 2: dynamic-array functions that filter,
 * sort, reshape, combine and summarize arrays.
 */
import { arg, categoryBuilder, more, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Lookup & Reference');

const STACKED = [
	arg('array1', 'The first array to append.'),
	more('array2', 'More arrays to append (up to 254 in total).'),
];
const TO_VECTOR = [
	arg('array', 'The array or reference to flatten.'),
	opt('ignore', 'Values to skip: 0 or omitted none, 1 blanks, 2 errors, 3 blanks and errors.'),
	opt('scan_by_column', 'FALSE or omitted reads the array by row; TRUE reads it by column.'),
];
const FIELD_HEADERS = opt(
	'field_headers',
	'Headers: 0 none, 1 present but hidden, 2 none but generate them, 3 present and shown.',
);
const AGGREGATE_FUNCTION = arg('function', 'How to aggregate the values, such as SUM or a LAMBDA.');
const FILTER_ARRAY = opt('filter_array', 'A Boolean column marking which rows to include.');

export const LOOKUP_ARRAY_INFO: readonly FunctionInfo[] = [
	fn('CHOOSECOLS', 'Picks columns of an array by their numbers, in any order.', [
		arg('array', 'The array to take columns from.'),
		arg('col_num1', 'Number of the first column to return; negative numbers count from the end.'),
		more('col_num2', 'More column numbers to return.'),
	]),
	fn('CHOOSEROWS', 'Picks rows of an array by their numbers, in any order.', [
		arg('array', 'The array to take rows from.'),
		arg('row_num1', 'Number of the first row to return; negative numbers count from the end.'),
		more('row_num2', 'More row numbers to return.'),
	]),
	fn('DROP', 'Returns an array without its first or last few rows or columns.', [
		arg('array', 'The array to drop rows or columns from.'),
		arg('rows', 'Rows to drop; negative values drop from the end.'),
		opt('columns', 'Columns to drop; negative values drop from the end.'),
	]),
	fn('EXPAND', 'Grows an array to a given number of rows and columns, padding the new cells.', [
		arg('array', 'The array to expand.'),
		arg('rows', 'Number of rows of the result; leave it empty to keep the current count.'),
		opt('columns', 'Number of columns of the result; the current count if omitted.'),
		opt('pad_with', 'Value for the new cells; #N/A if omitted.'),
	]),
	fn('FILTER', 'Keeps the rows or columns of a range or array that meet conditions.', [
		arg('array', 'The array or range to filter.'),
		arg('include', 'A Boolean array, as tall or as wide as array, marking what to keep.'),
		opt('if_empty', 'Value to return when nothing is kept.'),
	]),
	fn(
		'GROUPBY',
		'Groups rows by the values of one or more fields and aggregates the related values.',
		[
			arg('row_fields', 'Column or columns whose values define the groups.'),
			arg('values', 'Column or columns of values to aggregate.'),
			AGGREGATE_FUNCTION,
			FIELD_HEADERS,
			opt(
				'total_depth',
				'Totals: 0 none, 1 or omitted grand total, 2 with subtotals; negative puts them on top.',
			),
			opt('sort_order', 'Index of the column to sort by; negative values sort descending.'),
			FILTER_ARRAY,
			opt('field_relationship', 'With several row fields: 0 or omitted hierarchy, 1 table.'),
		],
	),
	fn('HSTACK', 'Places arrays side by side, left to right, to form one wider array.', STACKED),
	fn(
		'PIVOTBY',
		'Groups rows and columns by field values and aggregates the related values in a cross-tab.',
		[
			arg('row_fields', 'Column or columns whose values define the row groups.'),
			arg('col_fields', 'Column or columns whose values define the column groups.'),
			arg('values', 'Column or columns of values to aggregate.'),
			AGGREGATE_FUNCTION,
			FIELD_HEADERS,
			opt('row_total_depth', 'Row totals: 0 none, 1 or omitted grand total, 2 with subtotals.'),
			opt('row_sort_order', 'Index of the row field to sort by; negative values sort descending.'),
			opt('col_total_depth', 'Column totals: 0 none, 1 or omitted grand total, 2 with subtotals.'),
			opt(
				'col_sort_order',
				'Index of the column field to sort by; negative values sort descending.',
			),
			FILTER_ARRAY,
			opt(
				'relative_to',
				'Basis for percentages: 0 or omitted column totals, 1 row totals, 2 grand total.',
			),
		],
	),
	fn('SORT', 'Returns a range or array in sorted order, by one of its rows or columns.', [
		arg('array', 'The range or array to sort.'),
		opt('sort_index', 'Row or column number to sort by; 1 if omitted.'),
		opt('sort_order', '1 or omitted ascending, -1 descending.'),
		opt('by_col', 'FALSE or omitted sorts rows; TRUE sorts columns.'),
	]),
	fn(
		'SORTBY',
		'Orders a range or array using the values of one or more other ranges or arrays as keys.',
		[
			arg('array', 'The range or array to sort.'),
			arg('by_array1', 'The range or array to sort by.'),
			opt('sort_order1', '1 or omitted ascending, -1 descending.'),
			more('by_array2', 'Another range or array to sort by.'),
			more('sort_order2', 'Order for that range or array: 1 ascending, -1 descending.'),
		],
	),
	fn('TAKE', 'Keeps only the first or last few rows or columns of an array.', [
		arg('array', 'The array to take rows or columns from.'),
		arg('rows', 'Rows to take; negative values take from the end.'),
		opt('columns', 'Columns to take; negative values take from the end.'),
	]),
	fn(
		'TOCOL',
		'Flattens an array into one column, optionally skipping blanks and errors.',
		TO_VECTOR,
	),
	fn('TOROW', 'Flattens an array into one row, optionally skipping blanks and errors.', TO_VECTOR),
	fn('TRIMRANGE', 'Removes empty rows and columns from the outer edges of a range or array.', [
		arg('range', 'The range or array to trim.'),
		opt('trim_rows', 'Rows to trim: 0 none, 1 leading, 2 trailing, 3 or omitted both.'),
		opt('trim_cols', 'Columns to trim: 0 none, 1 leading, 2 trailing, 3 or omitted both.'),
	]),
	fn(
		'UNIQUE',
		'Removes duplicates from a range or array, keeping each distinct row or column once.',
		[
			arg('array', 'The range or array to get unique rows or columns from.'),
			opt('by_col', 'FALSE or omitted compares rows; TRUE compares columns.'),
			opt(
				'exactly_once',
				'TRUE returns only values that occur once; FALSE or omitted every distinct value.',
			),
		],
	),
	fn('VSTACK', 'Places arrays one below the other to form one taller array.', STACKED),
	fn('WRAPCOLS', 'Reshapes a single row or column into a 2-D array, filling it column by column.', [
		arg('vector', 'The row or column to wrap.'),
		arg('wrap_count', 'Maximum number of values in each column.'),
		opt('pad_with', 'Value used to pad the last column; #N/A if omitted.'),
	]),
	fn('WRAPROWS', 'Reshapes a single row or column into a 2-D array, filling it row by row.', [
		arg('vector', 'The row or column to wrap.'),
		arg('wrap_count', 'Maximum number of values in each row.'),
		opt('pad_with', 'Value used to pad the last row; #N/A if omitted.'),
	]),
];
