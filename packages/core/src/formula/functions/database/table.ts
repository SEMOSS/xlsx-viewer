import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toText } from '../../coerce';
import type { FunctionArg, FunctionContext, Scalar } from '../../types';
import { isOmitted, scalarArg } from '../lookup/args';
import type { Grid } from '../lookup/grid';
import { toGrid } from '../lookup/grid';

/**
 * A database range for the D-functions: the first row holds the column
 * labels, every row below it is a record. References stay lazy, so only the
 * columns the criteria and the field touch are read.
 */
export interface Database {
	grid: Grid;
	/** Number of records (rows below the header), limited to the sheet's used range. */
	records: number;
	/** 0-based column by label key (first occurrence wins). */
	columns: Map<string, number>;
}

/** Case-insensitive key of a column label; `null` for a blank label. */
export function labelKey(v: Scalar): string | null {
	if (v === null || v === '') {
		return null;
	}
	const text = toText(v);
	return (isCellError(text) ? text.code : text).toLowerCase();
}

/** Read the database argument (a range or an array with a header row). */
export function readDatabase(arg: FunctionArg, ctx: FunctionContext): Database | CellError {
	const grid = toGrid(arg, ctx);
	if (isCellError(grid)) {
		return grid;
	}
	const columns = new Map<string, number>();
	const labelled = Math.min(grid.cols, grid.dataCols);
	for (let c = 0; c < labelled; c++) {
		const key = labelKey(grid.get(0, c));
		if (key !== null && !columns.has(key)) {
			columns.set(key, c);
		}
	}
	const records = Math.max(0, Math.min(grid.rows, grid.dataRows) - 1);
	return { grid, records, columns };
}

/**
 * Resolve the field argument: a column label (case-insensitive) or a
 * 1-based column number. Returns `null` when the field is omitted and
 * `optional` (DCOUNT / DCOUNTA then count whole records); anything that does
 * not name a column is `#VALUE!`.
 */
export function resolveField(
	db: Database,
	args: readonly FunctionArg[],
	index: number,
	optional: boolean,
): number | null | CellError {
	if (isOmitted(args, index)) {
		return optional ? null : ERRORS.VALUE;
	}
	const field = scalarArg(args, index);
	if (typeof field === 'number') {
		const k = Math.trunc(field);
		return k >= 1 && k <= db.grid.cols ? k - 1 : ERRORS.VALUE;
	}
	if (typeof field === 'string') {
		return db.columns.get(field.toLowerCase()) ?? ERRORS.VALUE;
	}
	return isCellError(field) ? field : ERRORS.VALUE;
}
