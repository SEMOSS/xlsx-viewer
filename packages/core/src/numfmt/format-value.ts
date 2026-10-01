/**
 * Formatting entry points: section selection, sign and color handling, and
 * the formula-bar (edit) representation of a value.
 */
import { DEFAULT_INDEXED_COLORS } from '../color/indexed-palette';
import { numberToText } from '../formula/coerce';
import type { CellValue } from '../types/cell';
import { serialToDateParts } from '../values/dates';
import { isCellError } from '../values/errors';
import { formatGeneral, GENERAL_WIDTH } from './general';
import { finish, newOut } from './output';
import { parseFormatCode } from './parse';
import { maxDateSerial, renderDate } from './render-date';
import { renderFraction } from './render-fraction';
import { renderGeneralSection, renderNumber, renderTextSection } from './render-number';
import type {
	NumberFormatCondition,
	FormatOptions,
	NumberFormatSection,
	FormattedValue,
	ParsedNumberFormat,
	SectionRule,
} from './types';

/** `selectSection` result: no numeric section applies, use General. */
export const SECTION_GENERAL = -1;
/** `selectSection` result: no condition matched (Excel shows `#####`). */
export const SECTION_NONE = -2;

const INVALID_TEXT = '#####';

function matches(rule: SectionRule, v: number): boolean {
	const c: NumberFormatCondition | undefined = rule.condition;
	if (c === undefined) {
		return true;
	}
	switch (c.op) {
		case '<':
			return v < c.value;
		case '<=':
			return v <= c.value;
		case '>':
			return v > c.value;
		case '>=':
			return v >= c.value;
		case '=':
			return v === c.value;
		case '<>':
			return v !== c.value;
	}
}

/**
 * Index into `parsed.numeric` of the section that formats `v`, or
 * {@link SECTION_GENERAL} / {@link SECTION_NONE}. With conditions, a value
 * matching neither of the first two sections falls to the third; with only
 * two conditional sections Excel shows `#####`, and a single conditional
 * section that does not match leaves the value in General.
 */
export function selectSection(parsed: ParsedNumberFormat, v: number): number {
	const n = parsed.numeric.length;
	if (n === 0) {
		return SECTION_GENERAL;
	}
	if (!parsed.conditional) {
		if (n === 1) {
			return 0;
		}
		if (n === 2) {
			return v < 0 ? 1 : 0;
		}
		return v > 0 ? 0 : v < 0 ? 1 : 2;
	}
	if (matches(parsed.rules[0], v)) {
		return 0;
	}
	if (n === 1) {
		return SECTION_GENERAL;
	}
	if (matches(parsed.rules[1], v)) {
		return 1;
	}
	return n >= 3 ? 2 : SECTION_NONE;
}

function sectionColor(section: NumberFormatSection, options: FormatOptions): string | undefined {
	if (section.color !== undefined) {
		return section.color;
	}
	if (section.colorIndex === undefined) {
		return undefined;
	}
	const i = section.colorIndex + 7;
	const entry = options.indexedColors?.[i] ?? DEFAULT_INDEXED_COLORS[i];
	return entry ? '#' + entry.slice(-6).toUpperCase() : undefined;
}

function formatNumberValue(
	v: number,
	parsed: ParsedNumberFormat,
	options: FormatOptions,
): FormattedValue {
	if (!Number.isFinite(v)) {
		return { text: '#NUM!' };
	}
	const width = options.generalMaxChars ?? GENERAL_WIDTH;
	if (parsed.isGeneral) {
		return { text: formatGeneral(v, width) };
	}
	const index = selectSection(parsed, v);
	if (index === SECTION_GENERAL) {
		return { text: formatGeneral(v, width) };
	}
	if (index === SECTION_NONE) {
		return { text: INVALID_TEXT, invalid: true };
	}
	const section = parsed.numeric[index];
	const minus = v < 0 && !parsed.rules[index].dropSign;
	const abs = Math.abs(v);
	const out = newOut();
	switch (section.kind) {
		case 'empty':
			return { text: '' };
		case 'general':
			renderGeneralSection(section, abs, width, out);
			break;
		case 'number':
		case 'scientific':
			renderNumber(section, abs, out);
			break;
		case 'fraction':
			renderFraction(section, abs, out);
			break;
		case 'date': {
			// Negative date-times only exist in the 1904 system (shown with a minus sign).
			const date1904 = options.date1904 === true;
			if ((minus && !date1904) || !renderDate(section, abs, date1904, out)) {
				return { text: INVALID_TEXT, invalid: true };
			}
			break;
		}
		default:
			break;
	}
	return finish(out, minus ? '-' : '', sectionColor(section, options));
}

function formatTextValue(
	text: string,
	parsed: ParsedNumberFormat,
	options: FormatOptions,
): FormattedValue {
	const section = parsed.text;
	if (section === null) {
		return { text };
	}
	const out = newOut();
	renderTextSection(section, text, out);
	return finish(out, '', sectionColor(section, options));
}

/**
 * Format a cell value with an Excel number format code (en-US output).
 * Numbers go through the positive / negative / zero (or conditional)
 * sections; text and booleans (as TRUE / FALSE) through the text section
 * when there is one and unchanged otherwise; errors show their code.
 */
export function formatValue(
	value: CellValue,
	formatCode: string,
	options: FormatOptions = {},
): FormattedValue {
	if (value === null || value === undefined) {
		return { text: '' };
	}
	if (typeof value === 'number') {
		return formatNumberValue(value, parseFormatCode(formatCode), options);
	}
	if (typeof value === 'string') {
		return formatTextValue(value, parseFormatCode(formatCode), options);
	}
	if (typeof value === 'boolean') {
		return formatTextValue(value ? 'TRUE' : 'FALSE', parseFormatCode(formatCode), options);
	}
	if (isCellError(value)) {
		return { text: value.code };
	}
	return { text: String(value) };
}

function pad2(n: number): string {
	return n < 10 ? '0' + n : String(n);
}

function timeOfDay(seconds: number): string {
	const h = Math.floor(seconds / 3600);
	const m = Math.floor(seconds / 60) % 60;
	const s = seconds % 60;
	return `${h % 12 || 12}:${pad2(m)}:${pad2(s)} ${h < 12 ? 'AM' : 'PM'}`;
}

function editDateText(value: number, timeOnly: boolean, date1904: boolean): string | null {
	if (value < 0) {
		return null;
	}
	const ticks = Math.round(value * 86_400);
	const days = Math.floor(ticks / 86_400);
	if (!(days <= maxDateSerial(date1904))) {
		return null;
	}
	const seconds = ticks - days * 86_400;
	if (timeOnly && days === 0) {
		return timeOfDay(seconds);
	}
	const p = serialToDateParts(days, date1904);
	const date = `${p.month}/${p.day}/${p.year}`;
	return seconds > 0 || timeOnly ? `${date} ${timeOfDay(seconds)}` : date;
}

/**
 * What Excel shows in the formula bar / cell editor: dates as `m/d/yyyy`
 * (plus ` h:mm:ss AM/PM` when there is a time part), times as
 * `h:mm:ss AM/PM`, percentages as `12.5%`, other numbers in full precision
 * without grouping, booleans as TRUE / FALSE.
 */
export function formatValueForEdit(
	value: CellValue,
	formatCode = 'General',
	options: FormatOptions = {},
): string {
	if (value === null || value === undefined) {
		return '';
	}
	if (typeof value === 'boolean') {
		return value ? 'TRUE' : 'FALSE';
	}
	if (typeof value === 'string') {
		return value;
	}
	if (typeof value !== 'number') {
		return isCellError(value) ? value.code : '';
	}
	if (!Number.isFinite(value)) {
		return '#NUM!';
	}
	const parsed = parseFormatCode(formatCode);
	if (parsed.isDate) {
		const text = editDateText(value, parsed.isTimeOnly, options.date1904 === true);
		if (text !== null) {
			return text;
		}
	}
	const first = parsed.numeric[0];
	if (first !== undefined && first.percent && first.kind !== 'date') {
		return numberToText(value * 100) + '%';
	}
	return numberToText(value);
}
