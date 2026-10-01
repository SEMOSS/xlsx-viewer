/**
 * Date / time rendering. The serial is rounded once to the finest unit the
 * section shows (whole seconds, or 1/10, 1/100, 1/1000 s with fractional
 * seconds) and every field derives from that rounded value, so 23:59:59.7
 * shown as `h:mm:ss` is `0:00:00` of the next day, while larger units are
 * truncated (`h:mm` shows 10:29:40 as `10:29`), like Excel.
 */
import { MAX_DATE_SERIAL, serialToDateParts, type DateParts } from '../values/dates';
import { DAY_ABBREVIATIONS, DAY_NAMES, MONTH_ABBREVIATIONS, MONTH_NAMES } from './locale-en-us';
import { putFill, type Out } from './output';
import type { DateField, NumberFormatSection, TimeUnit } from './types';

/** Days between the 1900 and 1904 epochs. */
const EPOCH_GAP_1904 = 1462;
const TICKS_PER_SECOND = [1, 10, 100, 1000];
const NO_PARTS: DateParts = {
	year: 1900,
	month: 1,
	day: 0,
	hour: 0,
	minute: 0,
	second: 0,
	millisecond: 0,
	weekday: 6,
};

/** Largest serial a date format can show (9999-12-31). */
export function maxDateSerial(date1904: boolean): number {
	return date1904 ? MAX_DATE_SERIAL - EPOCH_GAP_1904 : MAX_DATE_SERIAL;
}

function pad2(n: number): string {
	return n < 10 ? '0' + n : String(n);
}

function padTo(n: number, len: number): string {
	const s = String(n);
	return s.length < len ? '0'.repeat(len - s.length) + s : s;
}

interface Clock {
	hour: number;
	minute: number;
	second: number;
	hour12: boolean;
}

function fieldText(field: DateField, p: DateParts, c: Clock): string {
	switch (field) {
		case 'year2':
			return pad2(p.year % 100);
		case 'year4':
			return String(p.year);
		case 'buddhist2':
			return pad2((p.year + 543) % 100);
		case 'buddhist4':
			return String(p.year + 543);
		case 'era':
			return '';
		case 'month':
			return String(p.month);
		case 'monthPad':
			return pad2(p.month);
		case 'monthShort':
			return MONTH_ABBREVIATIONS[p.month - 1];
		case 'monthLong':
			return MONTH_NAMES[p.month - 1];
		case 'monthLetter':
			return MONTH_NAMES[p.month - 1].charAt(0);
		case 'day':
			return String(p.day);
		case 'dayPad':
			return pad2(p.day);
		case 'weekdayShort':
			return DAY_ABBREVIATIONS[p.weekday];
		case 'weekdayLong':
			return DAY_NAMES[p.weekday];
		case 'hour':
		case 'hourPad': {
			const h = c.hour12 ? c.hour % 12 || 12 : c.hour;
			return field === 'hour' ? String(h) : pad2(h);
		}
		case 'minute':
			return String(c.minute);
		case 'minutePad':
			return pad2(c.minute);
		case 'second':
			return String(c.second);
		case 'secondPad':
			return pad2(c.second);
	}
}

function unitTicks(unit: TimeUnit, perSecond: number): number {
	return unit === 'h' ? 3600 * perSecond : unit === 'm' ? 60 * perSecond : perSecond;
}

/**
 * Render a date section for a non-negative serial. Returns false when the
 * value is beyond 9999-12-31 (Excel shows `#####`).
 */
export function renderDate(
	section: NumberFormatSection,
	value: number,
	date1904: boolean,
	out: Out,
): boolean {
	const perSecond = TICKS_PER_SECOND[section.subsecDigits];
	const perDay = 86_400 * perSecond;
	const ticks = Math.round(value * perDay);
	const days = Math.floor(ticks / perDay);
	if (!(days <= maxDateSerial(date1904))) {
		return false;
	}
	const rem = ticks - days * perDay;
	const clock: Clock = {
		hour: Math.floor(rem / (3600 * perSecond)),
		minute: Math.floor(rem / (60 * perSecond)) % 60,
		second: Math.floor(rem / perSecond) % 60,
		hour12: section.hour12,
	};
	const parts = section.hasDate ? serialToDateParts(days, date1904) : NO_PARTS;
	for (const item of section.items) {
		switch (item.k) {
			case 'lit':
				out.s += item.text;
				break;
			case 'fill':
				putFill(out, item.char);
				break;
			case 'date':
				out.s += fieldText(item.field, parts, clock);
				break;
			case 'elapsed': {
				const total = item.unit === section.elapsedUnit;
				const v = total
					? Math.floor(ticks / unitTicks(item.unit, perSecond))
					: item.unit === 'h'
						? clock.hour
						: item.unit === 'm'
							? clock.minute
							: clock.second;
				out.s += padTo(v, item.len);
				break;
			}
			case 'subsec':
				out.s += '.' + padTo(rem % perSecond, section.subsecDigits).slice(0, item.len);
				break;
			case 'ampm':
				out.s += clock.hour < 12 ? item.am : item.pm;
				break;
			default:
				break;
		}
	}
	return true;
}
