/**
 * Compiler for date / time sections. Resolves `m` / `mm` to minutes when
 * they follow an hour code or precede a seconds code (Excel's rule), turns
 * `.0` / `.00` / `.000` after seconds into fractional seconds and maps every
 * code to the field it prints.
 */
import { ItemList, literalText } from './output';
import type { DateField, DateLetter, NumberFormatToken, RenderItem, TimeUnit } from './types';

export interface DateLayout {
	items: RenderItem[];
	hasDate: boolean;
	hasTime: boolean;
	hour12: boolean;
	subsecDigits: number;
	elapsedUnit: TimeUnit | null;
}

const UNIT_RANK: Readonly<Record<TimeUnit, number>> = { s: 0, m: 1, h: 2 };

function isUnit(t: NumberFormatToken, unit: 'h' | 's'): boolean {
	return (t.type === 'date' && t.part === unit) || (t.type === 'elapsed' && t.unit === unit);
}

/** `m` / `mm` at `i` are minutes: the previous time code is an hour or the next is seconds. */
function isMinute(body: readonly NumberFormatToken[], i: number): boolean {
	for (let p = i - 1; p >= 0; p--) {
		const t = body[p];
		if (t.type === 'date' || t.type === 'elapsed') {
			if (isUnit(t, 'h')) {
				return true;
			}
			break;
		}
	}
	for (let q = i + 1; q < body.length; q++) {
		const t = body[q];
		if (t.type === 'date' || t.type === 'elapsed') {
			return isUnit(t, 's');
		}
	}
	return false;
}

function dateField(part: DateLetter, len: number, minute: boolean): DateField {
	switch (part) {
		case 'y':
			return len <= 2 ? 'year2' : 'year4';
		case 'e':
			return 'year4';
		case 'b':
			return len <= 2 ? 'buddhist2' : 'buddhist4';
		case 'g':
			return 'era';
		case 'd':
			return len === 1 ? 'day' : len === 2 ? 'dayPad' : len === 3 ? 'weekdayShort' : 'weekdayLong';
		case 'h':
			return len === 1 ? 'hour' : 'hourPad';
		case 's':
			return len === 1 ? 'second' : 'secondPad';
		case 'm':
			if (minute) {
				return len === 1 ? 'minute' : 'minutePad';
			}
			if (len <= 2) {
				return len === 1 ? 'month' : 'monthPad';
			}
			return len === 3 ? 'monthShort' : len === 5 ? 'monthLetter' : 'monthLong';
	}
}

const TIME_FIELDS: ReadonlySet<DateField> = new Set<DateField>([
	'hour',
	'hourPad',
	'minute',
	'minutePad',
	'second',
	'secondPad',
]);

export function compileDate(body: readonly NumberFormatToken[]): DateLayout {
	const n = body.length;
	const list = new ItemList();
	let hasDate = false;
	let hasTime = false;
	let subsecDigits = 0;
	let elapsedUnit: TimeUnit | null = null;
	const hour12 = body.some((t) => t.type === 'ampm');
	for (let i = 0; i < n; i++) {
		const t = body[i];
		switch (t.type) {
			case 'date': {
				const minute = t.part === 'm' && t.length <= 2 && isMinute(body, i);
				const field = dateField(t.part, t.length, minute);
				list.push({ k: 'date', field });
				if (TIME_FIELDS.has(field)) {
					hasTime = true;
				} else {
					hasDate = true;
				}
				break;
			}
			case 'elapsed':
				list.push({ k: 'elapsed', unit: t.unit, len: t.length });
				hasTime = true;
				if (elapsedUnit === null || UNIT_RANK[t.unit] > UNIT_RANK[elapsedUnit]) {
					elapsedUnit = t.unit;
				}
				break;
			case 'ampm':
				list.push({ k: 'ampm', am: t.am, pm: t.pm });
				hasTime = true;
				break;
			case 'point': {
				let zeros = 0;
				if (i > 0 && isUnit(body[i - 1], 's')) {
					for (let j = i + 1; j < n; j++) {
						const z = body[j];
						if (z.type !== 'digit' || z.char !== '0') {
							break;
						}
						zeros++;
					}
				}
				if (zeros === 0) {
					list.lit('.');
					break;
				}
				const len = Math.min(zeros, 3);
				list.push({ k: 'subsec', len });
				if (zeros > len) {
					list.lit('0'.repeat(zeros - len));
				}
				subsecDigits = Math.max(subsecDigits, len);
				i += zeros;
				break;
			}
			case 'fill':
				list.push({ k: 'fill', char: t.char });
				break;
			default:
				list.lit(literalText(t));
		}
	}
	return { items: list.items, hasDate, hasTime, hour12, subsecDigits, elapsedUnit };
}
