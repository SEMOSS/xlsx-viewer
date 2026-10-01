import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FormulaValue, FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import type { RecordTest } from './criteria';
import { readCriteria } from './criteria';
import type { Database } from './table';
import { readDatabase, resolveField } from './table';

/** A prepared D-function call: the database, the field column (null = whole records) and the criteria. */
interface Query {
	db: Database;
	field: number | null;
	matches: RecordTest;
}

function prepare(
	args: readonly FunctionArg[],
	ctx: FunctionContext,
	fieldOptional: boolean,
): Query | CellError {
	const db = readDatabase(args[0], ctx);
	if (isCellError(db)) {
		return db;
	}
	const field = resolveField(db, args, 1, fieldOptional);
	if (isCellError(field)) {
		return field;
	}
	const matches = readCriteria(args[2], db, ctx);
	if (isCellError(matches)) {
		return matches;
	}
	return { db, field, matches };
}

/**
 * Numbers in the field column of the matching records. Text, logicals and
 * blanks are skipped (as SUM does inside ranges); an error is returned.
 */
function matchingNumbers(q: Query): number[] | CellError {
	const out: number[] = [];
	const col = q.field ?? 0;
	for (let r = 0; r < q.db.records; r++) {
		if (!q.matches(r)) {
			continue;
		}
		const v = q.db.grid.get(r + 1, col);
		if (typeof v === 'number') {
			out.push(v);
		} else if (isCellError(v)) {
			return v;
		}
	}
	return out;
}

function sum(values: readonly number[]): number {
	let s = 0;
	for (const v of values) {
		s += v;
	}
	return s;
}

/** Sum of squared deviations from the mean (two-pass, numerically stable). */
function squaredDeviations(values: readonly number[]): number {
	const mean = sum(values) / values.length;
	let ss = 0;
	for (const v of values) {
		ss += (v - mean) * (v - mean);
	}
	return ss;
}

type Aggregate = (values: number[]) => number | CellError;

/** A D-function that aggregates the numbers of the field column. */
function numeric(name: string, aggregate: Aggregate): FunctionDefinition {
	return {
		name,
		minArgs: 3,
		maxArgs: 3,
		params: ['any', 'value', 'any'],
		impl: (args, ctx): FormulaValue => {
			const q = prepare(args, ctx, false);
			if (isCellError(q)) {
				return q;
			}
			const values = matchingNumbers(q);
			return isCellError(values) ? values : aggregate(values);
		},
	};
}

const variance =
	(sample: boolean) =>
	(values: number[]): number | CellError => {
		const n = values.length;
		if (n < (sample ? 2 : 1)) {
			return ERRORS.DIV0;
		}
		return squaredDeviations(values) / (sample ? n - 1 : n);
	};

const stdev =
	(sample: boolean) =>
	(values: number[]): number | CellError => {
		const v = variance(sample)(values);
		return isCellError(v) ? v : Math.sqrt(v);
	};

/** DSUM(database, field, criteria): sum of the matching numbers. */
export const DSUM = numeric('DSUM', sum);

/** DAVERAGE: mean of the matching numbers; none is `#DIV/0!`. */
export const DAVERAGE = numeric('DAVERAGE', (v) => (v.length ? sum(v) / v.length : ERRORS.DIV0));

/** DMAX: largest matching number, 0 when there is none. */
export const DMAX = numeric('DMAX', (v) => (v.length ? v.reduce((a, b) => (b > a ? b : a)) : 0));

/** DMIN: smallest matching number, 0 when there is none. */
export const DMIN = numeric('DMIN', (v) => (v.length ? v.reduce((a, b) => (b < a ? b : a)) : 0));

/** DPRODUCT: product of the matching numbers, 0 when there is none. */
export const DPRODUCT = numeric('DPRODUCT', (v) => (v.length ? v.reduce((a, b) => a * b, 1) : 0));

/** DSTDEV: sample standard deviation (needs two numbers). */
export const DSTDEV = numeric('DSTDEV', stdev(true));

/** DSTDEVP: population standard deviation. */
export const DSTDEVP = numeric('DSTDEVP', stdev(false));

/** DVAR: sample variance (needs two numbers). */
export const DVAR = numeric('DVAR', variance(true));

/** DVARP: population variance. */
export const DVARP = numeric('DVARP', variance(false));

/**
 * A counting D-function. With the field omitted every matching record is
 * counted; otherwise `counts` decides which field values count.
 */
function counting(name: string, counts: (v: unknown) => boolean): FunctionDefinition {
	return {
		name,
		minArgs: 3,
		maxArgs: 3,
		params: ['any', 'value', 'any'],
		impl: (args, ctx) => {
			const q = prepare(args, ctx, true);
			if (isCellError(q)) {
				return q;
			}
			let n = 0;
			for (let r = 0; r < q.db.records; r++) {
				if (q.matches(r) && (q.field === null || counts(q.db.grid.get(r + 1, q.field)))) {
					n++;
				}
			}
			return n;
		},
	};
}

/** DCOUNT(database, [field], criteria): matching cells holding numbers (or matching records). */
export const DCOUNT = counting('DCOUNT', (v) => typeof v === 'number');

/** DCOUNTA(database, [field], criteria): matching non-blank cells (or matching records). */
export const DCOUNTA = counting('DCOUNTA', (v) => v !== null);

/**
 * DGET(database, field, criteria): the field value of the single matching
 * record. No match is `#VALUE!`, more than one `#NUM!`.
 */
export const DGET: FunctionDefinition = {
	name: 'DGET',
	minArgs: 3,
	maxArgs: 3,
	params: ['any', 'value', 'any'],
	impl: (args, ctx) => {
		const q = prepare(args, ctx, false);
		if (isCellError(q)) {
			return q;
		}
		let found = -1;
		for (let r = 0; r < q.db.records; r++) {
			if (q.matches(r)) {
				if (found >= 0) {
					return ERRORS.NUM;
				}
				found = r;
			}
		}
		return found < 0 ? ERRORS.VALUE : q.db.grid.get(found + 1, q.field ?? 0);
	},
};
