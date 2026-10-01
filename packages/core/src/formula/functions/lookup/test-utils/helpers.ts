import { columnToLetters } from '../../../../address/cell-address';
import { MockContext } from '../../../test-utils/mock-context';
import type { FormulaValue, FunctionContext, Scalar } from '../../../types';
import { isArrayValue, isRefValue } from '../../../types';

/**
 * Plain value of a function result for assertions: references are read
 * (a single cell becomes its scalar), arrays become their row lists.
 */
export function plain(v: FormulaValue, ctx: FunctionContext): Scalar | Scalar[][] {
	if (isRefValue(v)) {
		const a = ctx.resolveRef(v);
		return a.rows === 1 && a.cols === 1 ? a.values[0][0] : a.values;
	}
	if (isArrayValue(v)) {
		return v.values;
	}
	return v as Scalar;
}

/** Cells for a column of values starting at `startRow` (1-based), e.g. `column('A', [1, 2])`. */
export function column(col: string, values: Scalar[], startRow = 1): Record<string, Scalar> {
	const out: Record<string, Scalar> = {};
	values.forEach((v, i) => {
		if (v !== null) {
			out[`${col}${startRow + i}`] = v;
		}
	});
	return out;
}

/** Cells for a row of values starting at column `startCol` (0-based) on `row` (1-based). */
export function row(rowNumber: number, values: Scalar[], startCol = 0): Record<string, Scalar> {
	const out: Record<string, Scalar> = {};
	values.forEach((v, i) => {
		if (v !== null) {
			out[`${columnToLetters(startCol + i)}${rowNumber}`] = v;
		}
	});
	return out;
}

/** Cells for a block of rows; the top-left value goes to row `startRow` (1-based), column `startCol` (0-based). */
export function block(rows: Scalar[][], startRow = 1, startCol = 0): Record<string, Scalar> {
	const out: Record<string, Scalar> = {};
	rows.forEach((line, r) => Object.assign(out, row(startRow + r, line, startCol)));
	return out;
}

/** A MockContext that counts cell reads, to check that lookups stay lazy. */
export class CountingContext extends MockContext {
	reads = 0;
	private usedCache: ReturnType<MockContext['usedRange']> | undefined;

	override getCellValue(sheet: number, r: number, c: number): Scalar {
		this.reads++;
		return super.getCellValue(sheet, r, c);
	}

	/** Cached: the base implementation scans every cell, which is slow for 100k-row fixtures. */
	override usedRange(sheet: number): ReturnType<MockContext['usedRange']> {
		if (sheet !== 0) {
			return super.usedRange(sheet);
		}
		this.usedCache ??= super.usedRange(sheet);
		return this.usedCache;
	}
}
