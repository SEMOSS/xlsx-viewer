/**
 * Excel's Increase Decimal / Decrease Decimal on an existing format code.
 * Every numeric section is edited in place (so currency, accounting,
 * percent, scientific and conditional formats keep their shape); quoted
 * text, the text section, dates without seconds and fractions are left
 * alone. Times with seconds gain or lose fractional seconds (up to 3).
 */
import { formatGeneral } from './general';
import { parseFormatCode } from './parse';
import type { NumberFormatSection, NumberFormatToken } from './types';

/** Excel's limit on decimal places in a number format. */
export const MAX_FORMAT_DECIMALS = 30;

function generalBase(delta: number, value: number | undefined): string {
	let decimals = 0;
	let scientific = false;
	if (typeof value === 'number' && Number.isFinite(value)) {
		const shown = formatGeneral(value);
		const e = shown.indexOf('E');
		const mantissa = e < 0 ? shown : shown.slice(0, e);
		const dot = mantissa.indexOf('.');
		decimals = dot < 0 ? 0 : mantissa.length - dot - 1;
		scientific = e >= 0;
	}
	const d = Math.max(0, Math.min(MAX_FORMAT_DECIMALS, decimals + delta));
	const core = d > 0 ? '0.' + '0'.repeat(d) : '0';
	return scientific ? core + 'E+00' : core;
}

function isDigit(t: NumberFormatToken, char?: string): boolean {
	return t.type === 'digit' && (char === undefined || t.char === char);
}

/** Accounting zero sections (`"-"??`): the `?` after the dash stand for the decimals. */
function adjustDash(tokens: readonly NumberFormatToken[], raws: string[], delta: number): boolean {
	const dash = tokens.findIndex((t) => t.type === 'literal' && t.text === '-');
	if (dash < 0 || !tokens.some((t) => t.type === 'fill')) {
		return false;
	}
	const marks: number[] = [];
	for (let i = dash + 1; i < tokens.length && isDigit(tokens[i], '?'); i++) {
		marks.push(i);
	}
	if (tokens.some((t, i) => t.type === 'digit' && !marks.includes(i))) {
		return false;
	}
	if (delta > 0) {
		const add = Math.min(delta, MAX_FORMAT_DECIMALS - marks.length);
		if (add <= 0) {
			return false;
		}
		raws[marks.length > 0 ? marks[marks.length - 1] : dash] += '?'.repeat(add);
		return true;
	}
	const remove = Math.min(-delta, marks.length);
	for (let k = 0; k < remove; k++) {
		raws[marks[marks.length - 1 - k]] = '';
	}
	return remove > 0;
}

function adjustNumber(
	tokens: readonly NumberFormatToken[],
	raws: string[],
	delta: number,
): boolean {
	if (adjustDash(tokens, raws, delta)) {
		return true;
	}
	const expIdx = tokens.findIndex((t) => t.type === 'exponent');
	const end = expIdx < 0 ? tokens.length : expIdx;
	let point = -1;
	const digits: number[] = [];
	for (let i = 0; i < end; i++) {
		const t = tokens[i];
		if (t.type === 'point' && point < 0) {
			point = i;
		} else if (t.type === 'digit') {
			digits.push(i);
		}
	}
	if (digits.length === 0) {
		return false;
	}
	const decimals = point < 0 ? [] : digits.filter((i) => i > point);
	if (delta > 0) {
		const add = Math.min(delta, MAX_FORMAT_DECIMALS - decimals.length);
		if (add <= 0) {
			return false;
		}
		const zeros = '0'.repeat(add);
		if (point < 0) {
			raws[digits[digits.length - 1]] += '.' + zeros;
		} else {
			raws[decimals.length > 0 ? decimals[decimals.length - 1] : point] += zeros;
		}
		return true;
	}
	const remove = Math.min(-delta, decimals.length);
	if (remove === 0) {
		return false;
	}
	for (let k = 0; k < remove; k++) {
		raws[decimals[decimals.length - 1 - k]] = '';
	}
	if (remove === decimals.length) {
		raws[point] = '';
	}
	return true;
}

function adjustTime(tokens: readonly NumberFormatToken[], raws: string[], delta: number): boolean {
	let seconds = -1;
	tokens.forEach((t, i) => {
		if ((t.type === 'date' && t.part === 's') || (t.type === 'elapsed' && t.unit === 's')) {
			seconds = i;
		}
	});
	if (seconds < 0) {
		return false;
	}
	const point = tokens[seconds + 1]?.type === 'point' ? seconds + 1 : -1;
	const zeros: number[] = [];
	for (let i = point + 1; point >= 0 && i < tokens.length && isDigit(tokens[i], '0'); i++) {
		zeros.push(i);
	}
	if (delta > 0) {
		const add = Math.min(delta, 3 - zeros.length);
		if (add <= 0) {
			return false;
		}
		if (zeros.length > 0) {
			raws[zeros[zeros.length - 1]] += '0'.repeat(add);
		} else {
			raws[seconds] += '.' + '0'.repeat(add);
		}
		return true;
	}
	const remove = Math.min(-delta, zeros.length);
	for (let k = 0; k < remove; k++) {
		raws[zeros[zeros.length - 1 - k]] = '';
	}
	if (remove > 0 && remove === zeros.length) {
		raws[point] = '';
	}
	return remove > 0;
}

function adjustSection(section: NumberFormatSection, delta: number): string {
	const raws = section.tokens.map((t) => t.raw);
	if (section.kind === 'number' || section.kind === 'scientific') {
		adjustNumber(section.tokens, raws, delta);
	} else if (
		section.kind === 'date' &&
		!section.tokens.some((t) => t.type === 'locale' && t.system !== undefined)
	) {
		// System date / time tags ignore the written pattern, so they are left alone.
		adjustTime(section.tokens, raws, delta);
	}
	return raws.join('');
}

/**
 * Increase (`delta > 0`) or decrease the decimal places of a format code,
 * like Excel's ribbon buttons. For General, `value` (when given) sets the
 * starting point to the decimals General currently shows for it.
 */
export function adjustDecimals(formatCode: string, delta: number, value?: number): string {
	const step = Math.trunc(delta);
	const parsed = parseFormatCode(formatCode);
	if (parsed.isGeneral) {
		return step === 0 ? formatCode : generalBase(step, value);
	}
	if (step === 0) {
		return formatCode;
	}
	return parsed.sections
		.map((s) => (s === parsed.text ? s.tokens.map((t) => t.raw).join('') : adjustSection(s, step)))
		.join(';');
}
