/**
 * Lookup & Reference functions, part 1: lookups, references, links and
 * PivotTable data. Dynamic-array shaping functions live in `lookup-arrays.ts`.
 */
import { arg, categoryBuilder, more, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Lookup & Reference');

const LOOKUP_VALUE = arg('lookup_value', 'The value to search for.');
const XMATCH_MODE = opt(
	'match_mode',
	'0 or omitted exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcard, 3 regex.',
);
const XSEARCH_MODE = opt(
	'search_mode',
	'1 or omitted first to last, -1 last to first, 2 binary search ascending, -2 descending.',
);

export const LOOKUP_INFO: readonly FunctionInfo[] = [
	fn(
		'ADDRESS',
		'Returns the text address, such as "$B$3", of the cell at a given row and column number.',
		[
			arg('row_num', 'Row number of the reference.'),
			arg('column_num', 'Column number of the reference.'),
			opt(
				'abs_num',
				'Reference type: 1 or omitted absolute, 2 absolute row, 3 absolute column, 4 relative.',
			),
			opt('a1', 'TRUE or omitted for A1 style, FALSE for R1C1 style.'),
			opt('sheet_text', 'Name of a worksheet to include in the reference.'),
		],
	),
	fn(
		'AREAS',
		'Counts the separate areas (contiguous ranges or single cells) that make up a reference.',
		[
			arg(
				'reference',
				'A reference to cells; put several areas in parentheses, such as (A1:B2,D4).',
			),
		],
	),
	fn('CHOOSE', 'Picks a value from a list of values by its position.', [
		arg('index_num', 'Which value to return: 1 picks value1, 2 picks value2, and so on.'),
		arg('value1', 'The first value to choose from.'),
		more('value2', 'More values to choose from (up to 254 in total).'),
	]),
	fn('COLUMN', 'Gives the column number of a cell reference, or of the formula cell itself.', [
		opt('reference', 'The cell or range to get the column number of; the formula cell if omitted.'),
	]),
	fn('COLUMNS', 'Counts the columns of a reference or array.', [
		arg('array', 'An array, array formula or range reference.'),
	]),
	fn('FIELDVALUE', 'Gets one field, such as Price or Population, from a linked data type value.', [
		arg('value', 'A cell or value holding a linked data type.'),
		arg('field_name', 'Name of the field to retrieve, in quotes.'),
	]),
	fn('FORMULATEXT', 'Shows the formula held in a cell as text.', [
		arg('reference', 'A reference to a cell or range of cells.'),
	]),
	fn(
		'GETPIVOTDATA',
		'Retrieves a summary value from a PivotTable by naming its data field and items.',
		[
			arg('data_field', 'Name of the PivotTable field holding the data, in quotes.'),
			arg('pivot_table', 'A reference to any cell in the PivotTable.'),
			more('field1', 'A field name to filter by.'),
			more('item1', 'The item of that field to filter by.'),
		],
	),
	fn(
		'HLOOKUP',
		'Searches the top row of a table for a value and returns an item from the same column.',
		[
			arg('lookup_value', 'What to search for in the top row of table_array.'),
			arg('table_array', 'The range to look in; its top row holds the values to match.'),
			arg('row_index_num', 'Which row of table_array to return, counting its first row as 1.'),
			opt(
				'range_lookup',
				'FALSE requires an exact match; TRUE or omitted allows a close match in a sorted row.',
			),
		],
	),
	fn(
		'HYPERLINK',
		'Creates a link that opens a document, web page or workbook location when clicked.',
		[
			arg('link_location', 'Path or URL of the target, as text.'),
			opt('friendly_name', 'Text or value shown in the cell; link_location if omitted.'),
		],
	),
	fn('IMAGE', 'Places a picture from a web address into a cell, with alternative text.', [
		arg('source', 'The https URL of the image file.'),
		opt('alt_text', 'Alternative text describing the image, for accessibility.'),
		opt(
			'sizing',
			'Sizing: 0 or omitted fit the cell, 1 fill the cell, 2 original size, 3 custom size.',
		),
		opt('height', 'Custom height in pixels, with sizing 3.'),
		opt('width', 'Custom width in pixels, with sizing 3.'),
	]),
	fn('INDEX', 'Picks the cell or value at a given row and column position of a range or array.', [
		arg('array', 'A range or array; several ranges in parentheses give the reference form.'),
		arg('row_num', 'Row to return; 0 returns the whole column.'),
		opt('column_num', 'Column to return; 0 returns the whole row.'),
		opt('area_num', 'Which of several ranges to use; 1 if omitted.'),
	]),
	fn('INDIRECT', 'Converts a text string, such as "Sheet2!B4", into the reference it names.', [
		arg(
			'ref_text',
			'Text holding a reference (A1 or R1C1 style), a defined name or a table reference.',
		),
		opt('a1', 'TRUE or omitted for A1 style, FALSE for R1C1 style.'),
	]),
	fn(
		'LOOKUP',
		'Searches a sorted row or column and returns the matching item from another row or column.',
		[
			LOOKUP_VALUE,
			arg('lookup_vector', 'A single row or column, sorted in ascending order, to search.'),
			opt('result_vector', 'A row or column of the same size to return a value from.'),
		],
	),
	fn('MATCH', 'Finds a value in a single row or column and returns its position.', [
		LOOKUP_VALUE,
		arg('lookup_array', 'The single row or column to search.'),
		opt(
			'match_type',
			'1 or omitted largest value <= lookup_value, 0 exact match, -1 smallest value >=.',
		),
	]),
	fn(
		'OFFSET',
		'Shifts a reference by some rows and columns, optionally resizing it, and returns the result.',
		[
			arg('reference', 'The starting cell or range.'),
			arg('rows', 'Rows to move: negative values move up, positive values down.'),
			arg('cols', 'Columns to move: negative values move left, positive values right.'),
			opt('height', 'Height of the result in rows; the height of reference if omitted.'),
			opt('width', 'Width of the result in columns; the width of reference if omitted.'),
		],
	),
	fn('ROW', 'Gives the row number of a cell reference, or of the formula cell itself.', [
		opt('reference', 'The cell or range to get the row number of; the formula cell if omitted.'),
	]),
	fn('ROWS', 'Counts the rows of a reference or array.', [
		arg('array', 'An array, array formula or range reference.'),
	]),
	fn('RTD', 'Gets continuously updated data from a COM automation server, such as a market feed.', [
		arg('progID', 'Registered ProgID of the automation add-in, in quotes.'),
		arg('server', 'Name of the server running the add-in; "" for this computer.'),
		arg('topic1', 'The first parameter that identifies the data to retrieve.'),
		more('topic2', 'More parameters (up to 253 in total).'),
	]),
	fn('TRANSPOSE', 'Swaps the rows and columns of a range or array.', [
		arg('array', 'The range or array to transpose.'),
	]),
	fn(
		'VLOOKUP',
		'Searches the first column of a table for a value and returns an item from the same row.',
		[
			arg('lookup_value', 'What to search for in the leftmost column of table_array.'),
			arg('table_array', 'The range to look in; its leftmost column holds the values to match.'),
			arg(
				'col_index_num',
				'Which column of table_array to return, counting its first column as 1.',
			),
			opt(
				'range_lookup',
				'FALSE requires an exact match; TRUE or omitted allows a close match in sorted data.',
			),
		],
	),
	fn(
		'XLOOKUP',
		'Finds a value in one range or array and returns the item at the same position in another.',
		[
			LOOKUP_VALUE,
			arg('lookup_array', 'The array or range to search.'),
			arg('return_array', 'The array or range to return values from.'),
			opt('if_not_found', 'Value to return when nothing matches; #N/A if omitted.'),
			XMATCH_MODE,
			XSEARCH_MODE,
		],
	),
	fn(
		'XMATCH',
		'Finds a value in an array or range and returns its position, with flexible match modes.',
		[LOOKUP_VALUE, arg('lookup_array', 'The array or range to search.'), XMATCH_MODE, XSEARCH_MODE],
	),
];
