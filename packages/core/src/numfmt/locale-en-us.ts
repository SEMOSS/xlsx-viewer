/**
 * en-US locale tables used by the number format engine.
 */

export const MONTH_NAMES: readonly string[] = [
	'January',
	'February',
	'March',
	'April',
	'May',
	'June',
	'July',
	'August',
	'September',
	'October',
	'November',
	'December',
];

export const MONTH_ABBREVIATIONS: readonly string[] = MONTH_NAMES.map((m) => m.slice(0, 3));

export const DAY_NAMES: readonly string[] = [
	'Sunday',
	'Monday',
	'Tuesday',
	'Wednesday',
	'Thursday',
	'Friday',
	'Saturday',
];

export const DAY_ABBREVIATIONS: readonly string[] = DAY_NAMES.map((d) => d.slice(0, 3));

/** The eight color names Excel accepts in number formats (`[Red]`), lower-cased. */
export const NAMED_COLORS: ReadonlyMap<string, string> = new Map([
	['black', '#000000'],
	['blue', '#0000FF'],
	['cyan', '#00FFFF'],
	['green', '#00FF00'],
	['magenta', '#FF00FF'],
	['red', '#FF0000'],
	['white', '#FFFFFF'],
	['yellow', '#FFFF00'],
]);

/** Windows en-US long date pattern, used for `[$-F800]` / `[$-x-sysdate]`. */
export const SYSTEM_LONG_DATE_FORMAT = 'dddd, mmmm d, yyyy';

/** Windows en-US time pattern, used for `[$-F400]` / `[$-x-systime]`. */
export const SYSTEM_TIME_FORMAT = 'h:mm:ss AM/PM';

/** Symbols that mark a number format as a currency format. */
export const CURRENCY_SYMBOLS: ReadonlySet<string> = new Set([
	'$',
	'€',
	'£',
	'¥',
	'¢',
	'₹',
	'₩',
	'₽',
	'₺',
	'₪',
	'₫',
	'₱',
	'₦',
	'₴',
	'฿',
	'R$',
	'kr',
	'kr.',
	'zł',
	'Kč',
	'Ft',
	'CHF',
	'Fr.',
	'R',
	'RM',
	'Rp',
	'lei',
	'лв',
	'руб.',
]);
