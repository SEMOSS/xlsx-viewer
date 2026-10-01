import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError, parseErrorCode } from '../../../values/errors';
import { parseNumericText } from '../../../values/parse-number';
import { createCriteriaMatcher } from '../../criteria';
import type { FunctionArg, FunctionContext, Scalar } from '../../types';
import { toGrid } from '../lookup/grid';
import type { Database } from './table';
import { labelKey } from './table';

/** Predicate over record indexes (0-based, below the header row). */
export type RecordTest = (record: number) => boolean;

/** Text that is not a number, date, logical or error literal. */
function isPlainText(s: string, date1904: boolean): boolean {
	if (parseNumericText(s, { date1904 }) !== null) {
		return false;
	}
	const upper = s.trim().toUpperCase();
	return upper !== 'TRUE' && upper !== 'FALSE' && parseErrorCode(s) === null;
}

/** `text*`, keeping a trailing lone `~` literal so the added `*` stays a wildcard. */
function beginsWithPattern(text: string): string {
	let tildes = 0;
	for (let i = text.length - 1; i >= 0 && text[i] === '~'; i--) {
		tildes++;
	}
	return `${tildes % 2 === 1 ? `${text}~` : text}*`;
}

/**
 * Matcher for one criteria cell, following Excel's database / Advanced
 * Filter rules. A blank cell is no condition (`null`). Plain text matches
 * values that *begin* with it (`Dav` finds "David" and "Davis"; write `=Dav`
 * for an exact match). Comparisons (`>10`, `<>x`, `=`), numbers, logicals,
 * errors and wildcards behave as in COUNTIF.
 */
export function criterionMatcher(
	value: Scalar,
	date1904: boolean,
): ((v: Scalar) => boolean) | null {
	if (value === null || value === '') {
		return null;
	}
	if (typeof value === 'string') {
		const first = value[0];
		if (first !== '=' && first !== '<' && first !== '>' && isPlainText(value, date1904)) {
			return createCriteriaMatcher(beginsWithPattern(value), date1904);
		}
	}
	return createCriteriaMatcher(value, date1904);
}

interface Condition {
	col: number;
	test: (v: Scalar) => boolean;
}

/**
 * Compile a criteria range: a header row of column labels and one or more
 * condition rows. Conditions in one row must all hold (AND); a record
 * matching any row qualifies (OR). A row without conditions matches every
 * record. A condition under a label that is not a database column (a
 * computed criterion, which is not supported) never matches. Fewer than two
 * rows is `#VALUE!`.
 */
export function readCriteria(
	arg: FunctionArg,
	db: Database,
	ctx: FunctionContext,
): RecordTest | CellError {
	const grid = toGrid(arg, ctx);
	if (isCellError(grid)) {
		return grid;
	}
	if (grid.rows < 2) {
		return ERRORS.VALUE;
	}
	const all: RecordTest = () => true;
	// Rows past the used range are entirely blank, and a blank row matches everything.
	if (grid.rows > grid.dataRows) {
		return all;
	}
	const cols = Math.min(grid.cols, grid.dataCols);
	const targets: Array<number | undefined> = [];
	for (let c = 0; c < cols; c++) {
		const key = labelKey(grid.get(0, c));
		targets.push(key === null ? undefined : db.columns.get(key));
	}
	const alternatives: Condition[][] = [];
	for (let r = 1; r < grid.rows; r++) {
		const conditions: Condition[] = [];
		let impossible = false;
		for (let c = 0; c < cols; c++) {
			const test = criterionMatcher(grid.get(r, c), ctx.date1904);
			if (!test) {
				continue;
			}
			const col = targets[c];
			if (col === undefined) {
				impossible = true;
				break;
			}
			conditions.push({ col, test });
		}
		if (impossible) {
			continue;
		}
		if (conditions.length === 0) {
			return all;
		}
		alternatives.push(conditions);
	}
	const data = db.grid;
	const rowMatches = (conditions: Condition[], record: number): boolean => {
		for (const { col, test } of conditions) {
			if (!test(data.get(record + 1, col))) {
				return false;
			}
		}
		return true;
	};
	if (alternatives.length === 1) {
		const only = alternatives[0];
		return (record) => rowMatches(only, record);
	}
	return (record) => alternatives.some((conditions) => rowMatches(conditions, record));
}
