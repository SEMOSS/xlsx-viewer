import type { CellError } from '../types/cell';
import { dateToSerial } from '../values/dates';
import { cellError, parseErrorCode } from '../values/errors';

const ISO_RE =
	/^(?:(-?\d{4,})-(\d{2})-(\d{2}))?(?:T?(\d{2}):(\d{2})(?::(\d{2})(?:[.,](\d+))?)?)?(?:Z|[+-]\d{2}(?::?\d{2})?)?$/;

/**
 * Serial number for an ISO 8601 date / time as stored in `t="d"` cells
 * (Strict OOXML and some producers). Time-zone designators are ignored:
 * spreadsheet dates are calendar values. Returns `undefined` when the text
 * is not an ISO date.
 */
export function isoToSerial(text: string, date1904: boolean): number | undefined {
	const s = text.trim();
	if (!s) {
		return undefined;
	}
	const m = ISO_RE.exec(s);
	if (!m || (m[1] === undefined && m[4] === undefined)) {
		return undefined;
	}
	let serial = 0;
	if (m[1] !== undefined) {
		const year = Number(m[1]);
		const month = Number(m[2]);
		const day = Number(m[3]);
		if (month < 1 || month > 12 || day < 1 || day > 31) {
			return undefined;
		}
		serial = dateToSerial(year, month, day, date1904);
	}
	if (m[4] !== undefined) {
		const seconds =
			Number(m[4]) * 3600 +
			Number(m[5]) * 60 +
			(m[6] !== undefined ? Number(m[6]) : 0) +
			(m[7] !== undefined ? Number(`0.${m[7]}`) : 0);
		serial += seconds / 86_400;
	}
	return serial;
}

/** Error value for an error literal in a `t="e"` cell; unknown literals become `#UNKNOWN!`. */
export function errorValue(text: string): CellError {
	return cellError(parseErrorCode(text) ?? '#UNKNOWN!');
}
