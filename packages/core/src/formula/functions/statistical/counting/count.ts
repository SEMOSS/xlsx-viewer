import { ERRORS, isCellError } from '../../../../values/errors';
import { forEachScalar, toNumber } from '../../../coerce';
import type { FunctionContext, FunctionDefinition, RefValue } from '../../../types';
import { isArrayValue, isMultiRefValue, isRefValue } from '../../../types';
import { define, isMissing } from '../args';

/**
 * COUNT: numbers inside references and arrays; typed arguments also count
 * when they are logical values or text that reads as a number.
 */
const COUNT: FunctionDefinition = define({
	name: 'COUNT',
	minArgs: 1,
	maxArgs: Infinity,
	params: ['any'],
	impl: (args, ctx) => {
		let n = 0;
		for (const arg of args) {
			if (isMissing(arg)) {
				continue;
			}
			forEachScalar(arg, ctx, (v, fromRef) => {
				if (typeof v === 'number') {
					n++;
				} else if (!fromRef) {
					if (typeof v === 'boolean') {
						n++;
					} else if (typeof v === 'string' && typeof toNumber(v, ctx.date1904) === 'number') {
						n++;
					}
				}
			});
		}
		return n;
	},
});

/** COUNTA: every non-empty value, including errors and empty text. */
const COUNTA: FunctionDefinition = define({
	name: 'COUNTA',
	minArgs: 1,
	maxArgs: Infinity,
	params: ['any'],
	impl: (args, ctx) => {
		let n = 0;
		for (const arg of args) {
			if (isMissing(arg)) {
				continue;
			}
			forEachScalar(arg, ctx, (v) => {
				if (v !== null) {
					n++;
				}
			});
		}
		return n;
	},
});

/** Number of cells in a (possibly 3-D) reference. */
export function cellCount(ref: RefValue): number {
	const r = ref.range;
	const sheets = ref.sheetEnd === undefined ? 1 : Math.abs(ref.sheetEnd - ref.sheet) + 1;
	return (r.bottom - r.top + 1) * (r.right - r.left + 1) * sheets;
}

function blankCells(ref: RefValue, ctx: FunctionContext): number {
	let filled = 0;
	ctx.forEachCell(ref, (v) => {
		if (v !== '') {
			filled++;
		}
	});
	return cellCount(ref) - filled;
}

/** COUNTBLANK: empty cells, counting formulas that return empty text as blank. */
const COUNTBLANK: FunctionDefinition = define({
	name: 'COUNTBLANK',
	minArgs: 1,
	maxArgs: 1,
	params: ['ref'],
	impl: (args, ctx) => {
		const arg = args[0];
		if (isRefValue(arg)) {
			return blankCells(arg, ctx);
		}
		if (isMultiRefValue(arg)) {
			return arg.refs.reduce((n, r) => n + blankCells(r, ctx), 0);
		}
		if (isArrayValue(arg)) {
			let n = 0;
			for (const row of arg.values) {
				for (const v of row) {
					if (v === null || v === '') {
						n++;
					}
				}
			}
			return n;
		}
		return isCellError(arg) ? arg : ERRORS.VALUE;
	},
});

export const COUNT_FUNCTIONS: FunctionDefinition[] = [COUNT, COUNTA, COUNTBLANK];
