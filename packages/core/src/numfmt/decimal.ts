/**
 * Exact decimal digit arithmetic for display rounding.
 *
 * Excel first rounds a double to 15 significant digits and then rounds half
 * away from zero to the displayed precision. Doing both steps on the decimal
 * digit string (instead of with binary floating point) gives Excel's results
 * for values such as 1.005 (-> 1.01) and 2.675 (-> 2.68), and makes `%` and
 * scaling commas exact shifts of the decimal exponent.
 */

/** A non-negative decimal `d0.d1d2... x 10^exp`; `digits` has no trailing zeros, '' is zero. */
export interface Decimal {
	readonly digits: string;
	readonly exp: number;
}

export const ZERO: Decimal = { digits: '', exp: 0 };

/** Decimal digits of `|x|` rounded to 15 significant digits. */
export function toDecimal(x: number): Decimal {
	if (x === 0 || !Number.isFinite(x)) {
		return ZERO;
	}
	const s = Math.abs(x).toExponential(14);
	const e = s.indexOf('e');
	let end = e;
	while (end > 2 && s.charCodeAt(end - 1) === 48) {
		end--;
	}
	const digits = end > 2 ? s.charAt(0) + s.slice(2, end) : s.charAt(0);
	return { digits, exp: Number(s.slice(e + 1)) };
}

/** Multiply by `10^shift` (exact). */
export function shiftDecimal(d: Decimal, shift: number): Decimal {
	return shift === 0 || d.digits === '' ? d : { digits: d.digits, exp: d.exp + shift };
}

/** Round half away from zero to `decimals` places after the decimal point (may be negative). */
export function roundDecimal(d: Decimal, decimals: number): Decimal {
	const { digits } = d;
	const keep = d.exp + decimals + 1;
	if (digits === '' || keep >= digits.length) {
		return d;
	}
	if (keep < 0) {
		return ZERO;
	}
	let exp = d.exp;
	let kept = digits.slice(0, keep);
	if (digits.charCodeAt(keep) >= 53) {
		let i = keep - 1;
		while (i >= 0 && kept.charCodeAt(i) === 57) {
			i--;
		}
		if (i < 0) {
			kept = '1';
			exp += 1;
		} else {
			kept = kept.slice(0, i) + String.fromCharCode(kept.charCodeAt(i) + 1);
		}
	}
	let end = kept.length;
	while (end > 0 && kept.charCodeAt(end - 1) === 48) {
		end--;
	}
	return end === 0 ? ZERO : { digits: kept.slice(0, end), exp };
}

/**
 * Integer digits ('' when the integer part is zero) and fraction digits of a
 * decimal, the fraction right-padded with zeros to at least `decimals` places.
 */
export function splitDecimal(d: Decimal, decimals: number): { int: string; frac: string } {
	const { digits, exp } = d;
	let int = '';
	let frac = '';
	if (digits !== '') {
		if (exp >= 0) {
			const n = exp + 1;
			if (digits.length > n) {
				int = digits.slice(0, n);
				frac = digits.slice(n);
			} else {
				int = digits.length === n ? digits : digits + '0'.repeat(n - digits.length);
			}
		} else {
			frac = '0'.repeat(-exp - 1) + digits;
		}
	}
	if (frac.length < decimals) {
		frac += '0'.repeat(decimals - frac.length);
	}
	return { int, frac };
}

/** Digits of a non-negative integer ('' for zero), limited to 15 significant digits like Excel. */
export function integerDigits(n: number): string {
	if (n < 1e15) {
		return n === 0 ? '' : String(n);
	}
	return splitDecimal(roundDecimal(toDecimal(n), 0), 0).int;
}

/** Plain decimal text of a decimal value (`'0'` for zero, no trailing zeros). */
export function decimalText(d: Decimal): string {
	const { int, frac } = splitDecimal(d, 0);
	return (int || '0') + (frac ? '.' + frac : '');
}
