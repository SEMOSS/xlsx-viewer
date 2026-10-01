/**
 * Complex numbers as Excel writes them: text such as "3+4i", "-2.5j", "i"
 * or "1E-3-2i" (the suffix is a lowercase `i` or `j`), or a plain real
 * number. Results are formatted with up to 15 significant digits per part.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { numberToText } from '../../coerce';
import type { Scalar } from '../../types';

export type Suffix = 'i' | 'j';

export interface Complex {
	re: number;
	im: number;
	/** Suffix written in the input; `null` for plain real numbers. */
	suffix: Suffix | null;
}

const NUM = '(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][+-]?\\d+)?';
const REAL_ONLY = new RegExp(`^[+-]?${NUM}$`);
const IMAG_ONLY = new RegExp(`^([+-]?)(${NUM})?([ij])$`);
const FULL = new RegExp(`^([+-]?${NUM})([+-])(${NUM})?([ij])$`);

/** Parse a complex-number argument; text that is not a complex number is #NUM!, logicals #VALUE!. */
export function parseComplex(v: Scalar): Complex | CellError {
	if (isCellError(v)) {
		return v;
	}
	if (v === null) {
		return { re: 0, im: 0, suffix: null };
	}
	if (typeof v === 'number') {
		return { re: v, im: 0, suffix: null };
	}
	if (typeof v === 'boolean') {
		return ERRORS.VALUE;
	}
	if (v === '') {
		return { re: 0, im: 0, suffix: null };
	}
	if (REAL_ONLY.test(v)) {
		return { re: Number(v), im: 0, suffix: null };
	}
	let m = IMAG_ONLY.exec(v);
	if (m) {
		const mag = m[2] === undefined ? 1 : Number(m[2]);
		return { re: 0, im: m[1] === '-' ? -mag : mag, suffix: m[3] as Suffix };
	}
	m = FULL.exec(v);
	if (m) {
		const mag = m[3] === undefined ? 1 : Number(m[3]);
		return { re: Number(m[1]), im: m[2] === '-' ? -mag : mag, suffix: m[4] as Suffix };
	}
	return ERRORS.NUM;
}

/** The suffix shared by several operands (`i` when none is written); mixing `i` and `j` is #VALUE!. */
export function commonSuffix(values: readonly Complex[]): Suffix | CellError {
	let suffix: Suffix | null = null;
	for (const z of values) {
		if (z.suffix !== null) {
			if (suffix !== null && suffix !== z.suffix) {
				return ERRORS.VALUE;
			}
			suffix = z.suffix;
		}
	}
	return suffix ?? 'i';
}

/** Excel text for a complex number; non-finite parts are #NUM!. */
export function formatComplex(re: number, im: number, suffix: Suffix): string | CellError {
	if (!Number.isFinite(re) || !Number.isFinite(im)) {
		return ERRORS.NUM;
	}
	const reText = numberToText(re);
	const imText = numberToText(im);
	if (imText === '0') {
		return reText;
	}
	const imPart = imText === '1' ? suffix : imText === '-1' ? `-${suffix}` : `${imText}${suffix}`;
	if (reText === '0') {
		return imPart;
	}
	return `${reText}${imPart.startsWith('-') ? '' : '+'}${imPart}`;
}

export interface Cx {
	re: number;
	im: number;
}

export const cx = (re: number, im = 0): Cx => ({ re, im });

export function add(a: Cx, b: Cx): Cx {
	return cx(a.re + b.re, a.im + b.im);
}

export function sub(a: Cx, b: Cx): Cx {
	return cx(a.re - b.re, a.im - b.im);
}

export function mul(a: Cx, b: Cx): Cx {
	return cx(a.re * b.re - a.im * b.im, a.re * b.im + a.im * b.re);
}

/** Quotient; division by zero yields NaN parts (reported as #NUM!). */
export function div(a: Cx, b: Cx): Cx {
	const d = b.re * b.re + b.im * b.im;
	if (d === 0) {
		return cx(Number.NaN, Number.NaN);
	}
	return cx((a.re * b.re + a.im * b.im) / d, (a.im * b.re - a.re * b.im) / d);
}

export function abs(z: Cx): number {
	return Math.hypot(z.re, z.im);
}

export function arg(z: Cx): number {
	return Math.atan2(z.im, z.re);
}

export function exp(z: Cx): Cx {
	const r = Math.exp(z.re);
	return cx(r * Math.cos(z.im), r * Math.sin(z.im));
}

/** Principal natural logarithm (NaN parts for zero). */
export function ln(z: Cx): Cx {
	if (z.re === 0 && z.im === 0) {
		return cx(Number.NaN, Number.NaN);
	}
	return cx(Math.log(abs(z)), arg(z));
}

/** Principal power `z^n` in polar form, as Excel computes it. */
export function pow(z: Cx, n: number): Cx {
	if (z.re === 0 && z.im === 0) {
		return n > 0 ? cx(0) : cx(Number.NaN, Number.NaN);
	}
	const r = Math.pow(abs(z), n);
	const t = arg(z) * n;
	return cx(r * Math.cos(t), r * Math.sin(t));
}

export function sqrt(z: Cx): Cx {
	const r = Math.sqrt(abs(z));
	const t = arg(z) / 2;
	return cx(r * Math.cos(t), r * Math.sin(t));
}

export function sin(z: Cx): Cx {
	return cx(Math.sin(z.re) * Math.cosh(z.im), Math.cos(z.re) * Math.sinh(z.im));
}

export function cos(z: Cx): Cx {
	return cx(Math.cos(z.re) * Math.cosh(z.im), -Math.sin(z.re) * Math.sinh(z.im));
}

export function sinh(z: Cx): Cx {
	return cx(Math.sinh(z.re) * Math.cos(z.im), Math.cosh(z.re) * Math.sin(z.im));
}

export function cosh(z: Cx): Cx {
	return cx(Math.cosh(z.re) * Math.cos(z.im), Math.sinh(z.re) * Math.sin(z.im));
}

export function inverse(z: Cx): Cx {
	return div(cx(1), z);
}

/*
 * tan, cot, sec and csc use the double-angle forms Excel evaluates (they
 * reproduce Excel's last digits), e.g. tan(a+bi) = (sin 2a + i sinh 2b) /
 * (cos 2a + cosh 2b). A zero denominator (a pole) yields non-finite parts.
 */

export function tan(z: Cx): Cx {
	const d = Math.cos(2 * z.re) + Math.cosh(2 * z.im);
	return cx(Math.sin(2 * z.re) / d, Math.sinh(2 * z.im) / d);
}

export function cot(z: Cx): Cx {
	const d = Math.cosh(2 * z.im) - Math.cos(2 * z.re);
	return cx(Math.sin(2 * z.re) / d, -Math.sinh(2 * z.im) / d);
}

export function sec(z: Cx): Cx {
	const d = Math.cos(2 * z.re) + Math.cosh(2 * z.im);
	return cx((2 * Math.cos(z.re) * Math.cosh(z.im)) / d, (2 * Math.sin(z.re) * Math.sinh(z.im)) / d);
}

export function csc(z: Cx): Cx {
	const d = Math.cosh(2 * z.im) - Math.cos(2 * z.re);
	return cx(
		(2 * Math.sin(z.re) * Math.cosh(z.im)) / d,
		(-2 * Math.cos(z.re) * Math.sinh(z.im)) / d,
	);
}
