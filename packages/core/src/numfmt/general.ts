/**
 * Excel's General number format.
 *
 * General shows as much precision as fits a character budget (11 for the
 * classic width): integers plainly while they fit, other values rounded to
 * the available decimals, and scientific notation (6 significant digits at
 * width 11, `1.23457E+11`) when the integer part does not fit or a tiny
 * value would lose its significant digits (`1E-10`). Between 1E-4 and 1 the
 * value is rounded to fit; below 1E-4 it is shown in fixed notation only if
 * it fits exactly. The minus sign of negative numbers is not counted.
 */
import { decimalText, roundDecimal, toDecimal, type Decimal } from './decimal';

/** Default General width in characters. */
export const GENERAL_WIDTH = 11;

function fixedGeneral(d: Decimal, width: number): string | null {
	const e = d.exp;
	if (e >= 0) {
		if (e + 1 > width) {
			return null;
		}
		const s = decimalText(roundDecimal(d, Math.max(0, width - e - 2)));
		return s.length > width ? null : s;
	}
	if (e >= -4) {
		const s = decimalText(roundDecimal(d, Math.max(0, width - 2)));
		return s === '0' || s.length > width ? null : s;
	}
	const s = decimalText(roundDecimal(d, width + 1));
	return s === '0' || s.length > width ? null : s;
}

function scientificGeneral(d: Decimal, width: number): string {
	let exp = d.exp;
	const expDigits = Math.abs(exp) >= 100 ? 3 : 2;
	const m = roundDecimal({ digits: d.digits, exp: 0 }, Math.max(0, width - 4 - expDigits));
	exp += m.exp;
	const mantissa = m.digits.length > 1 ? `${m.digits.charAt(0)}.${m.digits.slice(1)}` : m.digits;
	const abs = Math.abs(exp);
	return `${mantissa}E${exp < 0 ? '-' : '+'}${abs < 10 ? '0' : ''}${abs}`;
}

/**
 * Format a number with Excel's General format in a budget of `maxChars`
 * characters (minus sign excluded). `TEXT(x, "General")` uses 11.
 */
export function formatGeneral(value: number, maxChars: number = GENERAL_WIDTH): string {
	if (!Number.isFinite(value)) {
		return '#NUM!';
	}
	if (value === 0) {
		return '0';
	}
	const width = maxChars >= 1 ? Math.floor(maxChars) : GENERAL_WIDTH;
	const d = toDecimal(value);
	const text = fixedGeneral(d, width) ?? scientificGeneral(d, width);
	return value < 0 ? '-' + text : text;
}
