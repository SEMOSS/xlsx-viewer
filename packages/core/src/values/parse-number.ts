import { dateToSerial, daysInMonth, timeToSerial } from './dates';

/** Result of recognising numeric text, with the format Excel would auto-apply on entry. */
export interface ParsedNumber {
	value: number;
	/** Number format implied by how the text was written (`#,##0`, `0%`, `m/d/yyyy`, ...). */
	numFmt?: string;
	kind:
		| 'number'
		| 'percent'
		| 'currency'
		| 'date'
		| 'time'
		| 'datetime'
		| 'fraction'
		| 'scientific';
}

export interface ParseNumberOptions {
	date1904?: boolean;
	/** Year used when a date omits it (defaults to the current year). */
	defaultYear?: number;
}

const MONTHS = [
	'january',
	'february',
	'march',
	'april',
	'may',
	'june',
	'july',
	'august',
	'september',
	'october',
	'november',
	'december',
];

/** Month number (1-12) for a full or 3+ letter abbreviated English month name. */
function monthFromName(name: string): number | null {
	const n = name.toLowerCase().replace(/\.$/, '');
	if (n.length < 3) {
		return null;
	}
	for (let i = 0; i < 12; i++) {
		if (MONTHS[i] === n || MONTHS[i].startsWith(n)) {
			if (n.length >= 3) {
				return i + 1;
			}
		}
	}
	if (n === 'sept') {
		return 9;
	}
	return null;
}

function expandYear(y: number, digits: number): number {
	if (digits <= 2) {
		return y < 30 ? 2000 + y : 1900 + y;
	}
	return y;
}

function validDate(y: number, m: number, d: number): boolean {
	if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1) {
		return false;
	}
	if (y === 1900 && m === 2 && d === 29) {
		return true; // Excel's phantom leap day
	}
	return d <= daysInMonth(y, m);
}

function serialFor(y: number, m: number, d: number, opts: ParseNumberOptions): number {
	if (!opts.date1904 && y === 1900 && m === 2 && d === 29) {
		return 60;
	}
	return dateToSerial(y, m, d, opts.date1904);
}

interface TimeParse {
	value: number;
	numFmt: string;
}

/** `13:45`, `1:45 PM`, `1:45:30`, `3 pm`, `25:30`, `12:30:15.5`. */
function parseTime(text: string): TimeParse | null {
	const m = /^(\d{1,4})(?::(\d{1,2}))?(?::(\d{1,2}(?:\.\d+)?))?\s*([AaPp][Mm]?)?$/.exec(text);
	if (!m) {
		return null;
	}
	const hasMinutes = m[2] !== undefined;
	const ampm = m[4];
	if (!hasMinutes && !ampm) {
		return null;
	}
	let hour = Number(m[1]);
	const minute = hasMinutes ? Number(m[2]) : 0;
	const secondText = m[3];
	const second = secondText !== undefined ? Number(secondText) : 0;
	if (minute > 59 || second >= 60) {
		return null;
	}
	if (ampm) {
		if (hour < 1 || hour > 12) {
			return null;
		}
		const pm = ampm[0].toLowerCase() === 'p';
		if (hour === 12) {
			hour = pm ? 12 : 0;
		} else if (pm) {
			hour += 12;
		}
	}
	const value = timeToSerial(hour, minute, 0) + second / 86_400;
	let numFmt: string;
	if (secondText !== undefined && secondText.includes('.')) {
		numFmt = 'mm:ss.0';
	} else if (ampm) {
		numFmt = secondText !== undefined ? 'h:mm:ss AM/PM' : 'h:mm AM/PM';
	} else if (hour >= 24) {
		numFmt = '[h]:mm:ss';
	} else {
		numFmt = secondText !== undefined ? 'h:mm:ss' : 'h:mm';
	}
	return { value, numFmt };
}

interface DateParse {
	value: number;
	numFmt: string;
}

function parseDate(text: string, opts: ParseNumberOptions): DateParse | null {
	const year = opts.defaultYear ?? new Date().getFullYear();
	let m: RegExpExecArray | null;

	// m/d/yyyy, m-d-yyyy, m/d/yy, m/d
	if ((m = /^(\d{1,2})([/-])(\d{1,2})(?:\2(\d{1,4}))?$/.exec(text))) {
		const mon = Number(m[1]);
		const day = Number(m[3]);
		const y = m[4] !== undefined ? expandYear(Number(m[4]), m[4].length) : year;
		if (validDate(y, mon, day)) {
			return {
				value: serialFor(y, mon, day, opts),
				numFmt: m[4] !== undefined ? 'm/d/yyyy' : 'd-mmm',
			};
		}
		// m/yyyy -> first of month
		if (m[4] === undefined && Number(m[3]) > 31) {
			return null;
		}
		return null;
	}
	// yyyy-m-d, yyyy/m/d
	if ((m = /^(\d{4})([/-])(\d{1,2})\2(\d{1,2})$/.exec(text))) {
		const y = Number(m[1]);
		const mon = Number(m[3]);
		const day = Number(m[4]);
		if (validDate(y, mon, day)) {
			return { value: serialFor(y, mon, day, opts), numFmt: 'm/d/yyyy' };
		}
		return null;
	}
	// d-mmm-yyyy, d mmm yyyy, d-mmm, d mmm
	if ((m = /^(\d{1,2})[\s-]+([A-Za-z]{3,9}\.?)(?:[\s-]+(\d{2,4}))?$/.exec(text))) {
		const day = Number(m[1]);
		const mon = monthFromName(m[2]);
		if (mon === null) {
			return null;
		}
		const y = m[3] !== undefined ? expandYear(Number(m[3]), m[3].length) : year;
		if (validDate(y, mon, day)) {
			return {
				value: serialFor(y, mon, day, opts),
				numFmt: m[3] !== undefined ? 'd-mmm-yy' : 'd-mmm',
			};
		}
		return null;
	}
	// mmm d, yyyy / mmmm d yyyy / mmm d
	if ((m = /^([A-Za-z]{3,9}\.?)\s+(\d{1,2})(?:(?:,\s*|\s+)(\d{4}))?$/.exec(text))) {
		const mon = monthFromName(m[1]);
		if (mon === null) {
			return null;
		}
		const day = Number(m[2]);
		const y = m[3] !== undefined ? Number(m[3]) : year;
		if (validDate(y, mon, day)) {
			return {
				value: serialFor(y, mon, day, opts),
				numFmt: m[3] !== undefined ? 'd-mmm-yy' : 'd-mmm',
			};
		}
		return null;
	}
	// mmm-yy, mmm yyyy, mmm-yyyy -> first day of the month
	if ((m = /^([A-Za-z]{3,9}\.?)[\s-]+(\d{2}|\d{4})$/.exec(text))) {
		const mon = monthFromName(m[1]);
		if (mon === null) {
			return null;
		}
		const y = expandYear(Number(m[2]), m[2].length);
		if (validDate(y, mon, 1)) {
			return { value: serialFor(y, mon, 1, opts), numFmt: 'mmm-yy' };
		}
	}
	return null;
}

/**
 * Recognise numeric text the way Excel does when you type into a cell (en-US
 * conventions): plain and scientific numbers, thousands separators, percent,
 * dollar currency, accounting-style negatives, fractions, dates and times.
 * Returns `null` when the text is not numeric.
 */
export function parseNumericText(
	input: string,
	opts: ParseNumberOptions = {},
): ParsedNumber | null {
	const text = input.trim();
	if (text === '') {
		return null;
	}
	let m: RegExpExecArray | null;

	// Plain / scientific number.
	if ((m = /^([+-]?)(\d+\.?\d*|\.\d+)(?:[eE]([+-]?\d+))?$/.exec(text))) {
		const value = Number(text);
		if (!Number.isFinite(value)) {
			return null;
		}
		return m[3] !== undefined
			? { value, numFmt: '0.00E+00', kind: 'scientific' }
			: { value, kind: 'number' };
	}

	// Signs, currency, parentheses, thousands separators and percent.
	m =
		/^([+-]?)(\(?)\s*([+-]?)(\$?)\s*([+-]?)((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)\s*(\)?)\s*(%?)$/.exec(
			text,
		);
	if (m) {
		const [, sign1, open, sign2, dollar, sign3, digits, close, pct] = m;
		if ((open && !close) || (!open && close)) {
			return null;
		}
		const signs = [sign1, sign2, sign3].filter(Boolean);
		if (signs.length > 1) {
			return null;
		}
		if (dollar && pct) {
			return null;
		}
		let value = Number(digits.replace(/,/g, ''));
		if (!Number.isFinite(value)) {
			return null;
		}
		if (signs[0] === '-' || open) {
			value = -value;
		}
		const decimals = digits.includes('.') ? digits.split('.')[1].length : 0;
		const grouped = digits.includes(',');
		if (pct) {
			return { value: value / 100, numFmt: decimals > 0 ? '0.00%' : '0%', kind: 'percent' };
		}
		if (dollar) {
			return {
				value,
				numFmt: decimals > 0 ? '"$"#,##0.00' : '"$"#,##0',
				kind: 'currency',
			};
		}
		if (grouped) {
			return { value, numFmt: decimals > 0 ? '#,##0.00' : '#,##0', kind: 'number' };
		}
		return { value, kind: 'number' };
	}

	// Mixed fraction "1 1/2" or "0 3/4".
	if ((m = /^([+-]?)(\d+)\s+(\d+)\/(\d+)$/.exec(text))) {
		const whole = Number(m[2]);
		const num = Number(m[3]);
		const den = Number(m[4]);
		if (den === 0 || num >= den * 10) {
			return null;
		}
		const v = whole + num / den;
		return {
			value: m[1] === '-' ? -v : v,
			numFmt: m[4].length > 1 ? '# ??/??' : '# ?/?',
			kind: 'fraction',
		};
	}

	// Date and/or time.
	const dt =
		/^(.*?)\s+(\d{1,4}(?::\d{1,2}){1,2}(?:\.\d+)?\s*(?:[AaPp][Mm]?)?|\d{1,2}\s*[AaPp][Mm]?)$/.exec(
			text,
		);
	if (dt) {
		const d = parseDate(dt[1].trim(), opts);
		const t = parseTime(dt[2].trim());
		if (d && t) {
			return { value: d.value + t.value, numFmt: 'm/d/yyyy h:mm', kind: 'datetime' };
		}
	}
	const date = parseDate(text, opts);
	if (date) {
		return { value: date.value, numFmt: date.numFmt, kind: 'date' };
	}
	const time = parseTime(text);
	if (time) {
		return { value: time.value, numFmt: time.numFmt, kind: 'time' };
	}
	return null;
}

/** Convenience: numeric value of text, or `null`. */
export function textToNumber(text: string, opts: ParseNumberOptions = {}): number | null {
	return parseNumericText(text, opts)?.value ?? null;
}
