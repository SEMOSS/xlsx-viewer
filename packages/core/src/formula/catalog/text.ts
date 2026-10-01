/**
 * Text functions, part 1: the classic string functions. Joining, splitting,
 * regular expressions and language functions live in `text-modern.ts`.
 */
import { arg, categoryBuilder, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Text');

const TEXT = arg('text', 'The text, or a reference to a cell holding it.');
const SOURCE = arg('text', 'The text to extract from.');
const NUM_CHARS = opt('num_chars', 'Number of characters to extract; 1 if omitted.');
const NUM_BYTES = opt('num_bytes', 'Number of bytes to extract; 1 if omitted.');
const START_NUM = opt('start_num', 'Position to start searching from; 1 if omitted.');
const WITHIN_TEXT = arg('within_text', 'The text to search in.');
const DECIMALS = opt('decimals', 'Digits to the right of the decimal point; 2 if omitted.');

export const TEXT_INFO: readonly FunctionInfo[] = [
	fn(
		'ASC',
		'Turns full-width (double-byte) characters into their half-width (single-byte) forms.',
		[TEXT],
	),
	fn('BAHTTEXT', 'Converts a number to Thai text followed by the word for baht.', [
		arg('number', 'The number to convert.'),
	]),
	fn(
		'CHAR',
		'Gives the character that a code from 1 to 255 stands for in the system character set.',
		[arg('number', 'A code number from 1 to 255.')],
	),
	fn('CLEAN', 'Strips control characters and other non-printing characters from text.', [
		arg('text', 'The text to clean.'),
	]),
	fn('CODE', 'Gives the character-set code of the first character in a text.', [
		arg('text', 'The text whose first character to examine.'),
	]),
	fn(
		'DBCS',
		'Turns half-width (single-byte) characters into their full-width (double-byte) forms.',
		[TEXT],
	),
	fn('DOLLAR', 'Formats a number as currency text, rounded to a given number of decimals.', [
		arg('number', 'The number to format.'),
		DECIMALS,
	]),
	fn(
		'EXACT',
		'Compares two texts, including letter case, and returns TRUE only when they are identical.',
		[arg('text1', 'The first text.'), arg('text2', 'The second text.')],
	),
	fn(
		'FIND',
		'Locates one text inside another and gives its starting position; case matters, no wildcards.',
		[arg('find_text', 'The text to find.'), WITHIN_TEXT, START_NUM],
	),
	fn(
		'FINDB',
		'Locates one text inside another and gives its starting byte position; case matters.',
		[arg('find_text', 'The text to find.'), WITHIN_TEXT, START_NUM],
	),
	fn(
		'FIXED',
		'Formats a number as text with a fixed number of decimals and optional thousands separators.',
		[
			arg('number', 'The number to round and convert.'),
			DECIMALS,
			opt('no_commas', 'TRUE leaves out the thousands separators.'),
		],
	),
	fn('LEFT', 'Takes a number of characters from the start of a text.', [SOURCE, NUM_CHARS]),
	fn('LEFTB', 'Takes characters from the start of a text, counting by bytes.', [SOURCE, NUM_BYTES]),
	fn('LEN', 'Counts the characters of a text, spaces included.', [
		arg('text', 'The text to measure; spaces count as characters.'),
	]),
	fn('LENB', 'Counts the bytes taken by the characters of a text.', [
		arg('text', 'The text to measure.'),
	]),
	fn('LOWER', 'Changes every letter of a text to lowercase.', [
		arg('text', 'The text to convert.'),
	]),
	fn('MID', 'Takes a run of characters out of a text, given a start position and a length.', [
		SOURCE,
		arg('start_num', 'Position of the first character to extract; 1 is the first character.'),
		arg('num_chars', 'Number of characters to extract.'),
	]),
	fn(
		'MIDB',
		'Takes a run of characters out of a text, with the start and length counted in bytes.',
		[
			SOURCE,
			arg('start_num', 'Byte position of the first character to extract.'),
			arg('num_bytes', 'Number of bytes to extract.'),
		],
	),
	fn(
		'NUMBERVALUE',
		'Reads a number from text using explicitly given decimal and group separators.',
		[
			arg('text', 'The text to convert.'),
			opt('decimal_separator', 'Character used as the decimal separator.'),
			opt('group_separator', 'Character used to separate groups of digits, such as thousands.'),
		],
	),
	fn('PHONETIC', 'Returns the furigana (phonetic guide) stored with Japanese text in a cell.', [
		arg('reference', 'A cell or range holding the text with furigana.'),
	]),
	fn(
		'PROPER',
		'Changes a text to title case: a capital at the start of each word, lowercase elsewhere.',
		[arg('text', 'The text to convert.')],
	),
	fn(
		'REPLACE',
		'Overwrites the characters at a given position and length of a text with new text.',
		[
			arg('old_text', 'The text to change.'),
			arg('start_num', 'Position of the first character to replace.'),
			arg('num_chars', 'Number of characters to replace.'),
			arg('new_text', 'The text to insert.'),
		],
	),
	fn('REPLACEB', 'Overwrites part of a text, located by byte position and length, with new text.', [
		arg('old_text', 'The text to change.'),
		arg('start_num', 'Byte position of the first character to replace.'),
		arg('num_bytes', 'Number of bytes to replace.'),
		arg('new_text', 'The text to insert.'),
	]),
	fn('REPT', 'Builds a text by repeating another text several times.', [
		arg('text', 'The text to repeat.'),
		arg('number_times', 'How many times to repeat it.'),
	]),
	fn('RIGHT', 'Takes a number of characters from the end of a text.', [SOURCE, NUM_CHARS]),
	fn('RIGHTB', 'Takes characters from the end of a text, counting by bytes.', [SOURCE, NUM_BYTES]),
	fn(
		'SEARCH',
		'Locates one text inside another and gives its position; case is ignored and wildcards work.',
		[
			arg('find_text', 'The text to find; ? matches one character and * any sequence.'),
			WITHIN_TEXT,
			START_NUM,
		],
	),
	fn('SEARCHB', 'Locates one text inside another and gives its byte position; case is ignored.', [
		arg('find_text', 'The text to find.'),
		WITHIN_TEXT,
		START_NUM,
	]),
	fn('SUBSTITUTE', 'Swaps every occurrence, or one chosen occurrence, of a text for new text.', [
		arg('text', 'The text, or the cell, to change.'),
		arg('old_text', 'The text to replace.'),
		arg('new_text', 'The text to put in its place.'),
		opt('instance_num', 'Which occurrence to replace; every occurrence if omitted.'),
	]),
	fn('T', 'Passes text values through unchanged and turns any other value into empty text.', [
		arg('value', 'The value to test.'),
	]),
	fn('TEXT', 'Formats a value with a number format code and returns the result as text.', [
		arg('value', 'The number, formula or reference to format.'),
		arg('format_text', 'A number format in quotes, such as "0.00" or "dd/mm/yyyy".'),
	]),
	fn(
		'TRIM',
		'Cleans up spacing: drops leading and trailing spaces and collapses runs of spaces to one.',
		[arg('text', 'The text to trim.')],
	),
	fn('UNICHAR', 'Gives the character with a given Unicode code point.', [
		arg('number', 'The Unicode code point.'),
	]),
	fn('UNICODE', 'Gives the Unicode code point of the first character in a text.', [
		arg('text', 'The text whose first character to examine.'),
	]),
	fn('UPPER', 'Changes every letter of a text to uppercase.', [
		arg('text', 'The text to convert.'),
	]),
	fn('VALUE', 'Turns text that looks like a number, date or time into an actual number.', [
		arg('text', 'The text to convert, in quotes or as a cell reference.'),
	]),
];
