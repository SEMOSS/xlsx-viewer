/** Information functions: value tests (IS*), cell and environment details, error codes. */
import { arg, categoryBuilder, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Information');

const VALUE = arg('value', 'The value to test: a cell, formula, name or constant.');
const INTEGER = arg('number', 'The value to test; decimals are truncated.');

export const INFORMATION_INFO: readonly FunctionInfo[] = [
	fn('CELL', 'Reports a detail of a cell, such as its address, number format, contents or type.', [
		arg(
			'info_type',
			'The kind of information, such as "address", "col", "contents", "format" or "type".',
		),
		opt('reference', 'The cell to describe; the last changed cell if omitted.'),
	]),
	fn(
		'ERROR.TYPE',
		'Gives the code of an error value, such as 2 for #DIV/0!, or #N/A for non-errors.',
		[arg('error_val', 'The error value, or a reference to the cell to check.')],
	),
	fn(
		'INFO',
		'Reports a detail of the environment, such as the operating system or calculation mode.',
		[
			arg(
				'type_text',
				'The kind of information, such as "directory", "osversion", "recalc" or "system".',
			),
		],
	),
	fn('ISBLANK', 'Tests whether a reference points to an empty cell.', [VALUE]),
	fn('ISERR', 'Tests whether a value is an error other than #N/A.', [VALUE]),
	fn('ISERROR', 'Tests whether a value is any error, #N/A included.', [VALUE]),
	fn('ISEVEN', 'Tests whether a number is even.', [INTEGER]),
	fn('ISFORMULA', 'Tests whether a referenced cell contains a formula.', [
		arg('reference', 'The cell to test.'),
	]),
	fn('ISLOGICAL', 'Tests whether a value is TRUE or FALSE.', [VALUE]),
	fn('ISNA', 'Tests whether a value is the #N/A error.', [VALUE]),
	fn('ISNONTEXT', 'Tests whether a value is anything other than text, including a blank cell.', [
		VALUE,
	]),
	fn('ISNUMBER', 'Tests whether a value is a number.', [VALUE]),
	fn('ISODD', 'Tests whether a number is odd.', [INTEGER]),
	fn('ISOMITTED', 'Tests whether a LAMBDA parameter was left out of the call.', [
		arg('argument', 'The LAMBDA parameter to check.'),
	]),
	fn('ISREF', 'Tests whether a value is a reference.', [VALUE]),
	fn('ISTEXT', 'Tests whether a value is text.', [VALUE]),
	fn(
		'N',
		'Turns a value into a number: numbers and dates pass through, TRUE gives 1, anything else 0.',
		[arg('value', 'The value to convert.')],
	),
	fn('NA', 'Produces the #N/A error, to mark a value as not available.'),
	fn('SHEET', 'Gives the position number of a sheet in the workbook.', [
		opt('value', 'A sheet name or a reference; the sheet holding the formula if omitted.'),
	]),
	fn('SHEETS', 'Counts the sheets covered by a reference, or all sheets of the workbook.', [
		opt(
			'reference',
			'The reference to count the sheets of; all sheets of the workbook if omitted.',
		),
	]),
	fn(
		'TYPE',
		'Gives a code for the kind of value: 1 number, 2 text, 4 logical, 16 error, 64 array.',
		[arg('value', 'Any value.')],
	),
];
