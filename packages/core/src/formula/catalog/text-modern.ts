/**
 * Text functions, part 2: joining and splitting, regular expressions,
 * value-to-text conversion and language services.
 */
import { arg, categoryBuilder, more, opt, series } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Text');

const CASE_SENSITIVITY = opt(
	'case_sensitivity',
	'0 or omitted case-sensitive, 1 case-insensitive.',
);
const MATCH_MODE = opt('match_mode', '0 or omitted case-sensitive, 1 case-insensitive.');
const PATTERN = arg('pattern', 'The regular expression to match.');

/** Arguments shared by TEXTBEFORE and TEXTAFTER. */
const DELIMITED = [
	arg('text', 'The text to search.'),
	arg('delimiter', 'The text that marks the point to extract at.'),
	opt(
		'instance_num',
		'Which occurrence of the delimiter to use; negative values count from the end.',
	),
	MATCH_MODE,
	opt('match_end', '1 treats the end of the text as a delimiter; 0 or omitted does not.'),
	opt('if_not_found', 'Value to return when there is no match; #N/A if omitted.'),
];

export const TEXT_MODERN_INFO: readonly FunctionInfo[] = [
	fn('ARRAYTOTEXT', 'Turns the contents of an array or range into a single text string.', [
		arg('array', 'The array to convert.'),
		opt('format', '0 or omitted concise format, 1 strict format with quotes and row separators.'),
	]),
	fn(
		'CONCAT',
		'Joins text from several values, ranges or arrays into one string, with no delimiter.',
		[
			...series(
				'text',
				'The first item, range or array to join.',
				'More items to join (up to 253 in total).',
			),
		],
	),
	fn('DETECTLANGUAGE', 'Returns the language code of a text, using an online language service.', [
		arg('text', 'The text to analyze.'),
	]),
	fn('REGEXEXTRACT', 'Extracts the parts of a text that match a regular expression.', [
		arg('text', 'The text to search.'),
		PATTERN,
		opt(
			'return_mode',
			'0 or omitted first match, 1 all matches, 2 capture groups of the first match.',
		),
		CASE_SENSITIVITY,
	]),
	fn(
		'REGEXREPLACE',
		'Substitutes new text for every part of a text that a regular expression matches.',
		[
			arg('text', 'The text to change.'),
			PATTERN,
			arg('replacement', 'The text to put in place of each match; $1 inserts a capture group.'),
			opt(
				'occurrence',
				'Which match to replace: 0 or omitted all of them; negative values count from the end.',
			),
			CASE_SENSITIVITY,
		],
	),
	fn('REGEXTEST', 'Returns TRUE when a regular expression matches somewhere in a text.', [
		arg('text', 'The text to test.'),
		PATTERN,
		CASE_SENSITIVITY,
	]),
	fn('TEXTAFTER', 'Extracts the part of a text that follows a delimiter.', DELIMITED),
	fn('TEXTBEFORE', 'Extracts the part of a text that precedes a delimiter.', DELIMITED),
	fn(
		'TEXTJOIN',
		'Joins text from several values, ranges or arrays, putting a delimiter between items.',
		[
			arg('delimiter', 'Text to insert between the items; it can be empty.'),
			arg('ignore_empty', 'TRUE skips empty cells.'),
			arg('text1', 'The first item, range or array to join.'),
			more('text2', 'More items to join (up to 252 in total).'),
		],
	),
	fn(
		'TEXTSPLIT',
		'Breaks a text into an array of pieces, splitting at column and row delimiters.',
		[
			arg('text', 'The text to split.'),
			arg('col_delimiter', 'Text that marks where to split across columns.'),
			opt('row_delimiter', 'Text that marks where to split down rows.'),
			opt(
				'ignore_empty',
				'TRUE skips empty values between consecutive delimiters; FALSE or omitted keeps them.',
			),
			MATCH_MODE,
			opt('pad_with', 'Value used to pad missing entries; #N/A if omitted.'),
		],
	),
	fn('TRANSLATE', 'Converts text into another language using an online translation service.', [
		arg('text', 'The text to translate.'),
		opt('source_language', 'Language code of the text, such as "en"; detected if omitted.'),
		opt('target_language', 'Language code to translate into; the system language if omitted.'),
	]),
	fn('VALUETOTEXT', 'Turns a value of any type into text.', [
		arg('value', 'The value to convert.'),
		opt('format', '0 or omitted concise format, 1 strict format with quotes around text.'),
	]),
];
