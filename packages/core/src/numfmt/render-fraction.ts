/**
 * Fraction rendering (`# ?/?`, `# ??/??`, `?/?`, `# ?/8`, `# ??/100`).
 */
import { integerDigits } from './decimal';
import { intSlot, putFill, type Out } from './output';
import type { NumberFormatSection } from './types';

/**
 * Best rational approximation `[numerator, denominator]` of `x >= 0` with a
 * denominator of at most `maxDen`: continued-fraction convergents plus the
 * best semiconvergent at the limit (the Stern-Brocot best approximation).
 * Ties prefer the smaller denominator.
 */
export function approximateFraction(x: number, maxDen: number): [number, number] {
	const limit = Math.max(1, Math.floor(maxDen));
	if (!(x > 0) || !Number.isFinite(x)) {
		return [0, 1];
	}
	let p0 = 0;
	let q0 = 1;
	let p1 = 1;
	let q1 = 0;
	let r = x;
	for (let iter = 0; iter < 64; iter++) {
		const a = Math.floor(r);
		const p2 = a * p1 + p0;
		const q2 = a * q1 + q0;
		if (q2 > limit) {
			const k = Math.floor((limit - q0) / q1);
			const ps = k * p1 + p0;
			const qs = k * q1 + q0;
			if (k > 0 && Math.abs(x - ps / qs) < Math.abs(x - p1 / q1)) {
				return [ps, qs];
			}
			return [p1, q1];
		}
		p0 = p1;
		q0 = q1;
		p1 = p2;
		q1 = q2;
		const f = r - a;
		if (f < 1e-12 || Math.abs(x - p1 / q1) <= 1e-15 * x) {
			break;
		}
		r = 1 / f;
	}
	return [p1, q1];
}

/** Excel's 15 significant digit precision (removes binary noise such as 25.499999999999993). */
function round15(v: number): number {
	return Number(v.toPrecision(15));
}

/** Denominator placeholder `idx` (left-aligned; unused positions become spaces). */
function denSlot(digits: string, slots: number, idx: number, ph: string): string {
	if (idx < digits.length) {
		return idx === slots - 1 ? digits.slice(idx) : digits.charAt(idx);
	}
	return ph === '#' ? '' : ' ';
}

/**
 * Render a fraction section for a non-negative value. With a whole part,
 * a zero numerator blanks the fraction (spaces keep the alignment) and a
 * zero value shows `0`; without one the value is an improper fraction.
 */
export function renderFraction(section: NumberFormatSection, abs: number, out: Out): void {
	const x = round15(section.scale === 0 ? abs : abs * 10 ** section.scale);
	const hasWhole = section.intSlots > 0;
	let whole = 0;
	let num: number;
	let den: number;
	if (section.denFixed > 0) {
		den = section.denFixed;
		if (hasWhole) {
			whole = Math.floor(x);
			num = Math.round(round15((x - whole) * den));
			if (num >= den) {
				whole += 1;
				num -= den;
			}
		} else {
			num = Math.round(round15(x * den));
		}
	} else {
		const maxDen = 10 ** Math.min(section.denSlots, 15) - 1;
		if (hasWhole) {
			whole = Math.floor(x);
			[num, den] = approximateFraction(x - whole, maxDen);
			if (num >= den) {
				whole += 1;
				num = 0;
				den = 1;
			}
		} else {
			[num, den] = approximateFraction(x, maxDen);
		}
	}
	const blank = hasWhole && num === 0;
	const wholeDigits = whole > 0 ? integerDigits(whole) : blank ? '0' : '';
	const numDigits = String(num);
	const denDigits = String(den);
	for (const item of section.items) {
		switch (item.k) {
			case 'lit':
				out.s += item.text;
				break;
			case 'fill':
				putFill(out, item.char);
				break;
			case 'whole':
				out.s += intSlot(wholeDigits, section.intSlots, item.pos, item.ph, section.grouping);
				break;
			case 'num':
				out.s += blank ? ' ' : intSlot(numDigits, section.numSlots, item.pos, item.ph, false);
				break;
			case 'slash':
				out.s += blank ? ' ' : '/';
				break;
			case 'den':
				out.s += blank ? ' ' : denSlot(denDigits, section.denSlots, item.idx, item.ph);
				break;
			case 'denFixed':
				out.s += blank ? ' '.repeat(item.text.length) : item.text;
				break;
			default:
				break;
		}
	}
}
