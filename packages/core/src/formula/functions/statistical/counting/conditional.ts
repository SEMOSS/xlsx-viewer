import type { CellError } from '../../../../types/cell';
import { ERRORS, isCellError } from '../../../../values/errors';
import { createCriteriaMatcher } from '../../../criteria';
import type {
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	ParamKind,
	Scalar,
} from '../../../types';
import { isArrayValue, isRefValue } from '../../../types';
import { define } from '../args';
import { cellCount } from './count';
import { criteriaValue, type Grid, gridAt, resizeRef, sameShape, toGrid } from './grid';

type Matcher = (v: Scalar) => boolean;

/** Up to 127 (range, criteria) pairs after `lead` leading parameters. */
function pairedParams(lead: readonly ParamKind[]): ParamKind[] {
	const out = [...lead];
	for (let i = 0; i < 127; i++) {
		out.push('ref', 'value');
	}
	return out;
}

/** Criteria grids and matchers from (range, criteria) argument pairs starting at `from`. */
function criteriaPairs(
	args: readonly FunctionArg[],
	from: number,
	ctx: FunctionContext,
): { grids: Grid[]; matchers: Matcher[] } | CellError {
	if ((args.length - from) % 2 !== 0) {
		return ERRORS.VALUE;
	}
	const grids: Grid[] = [];
	const matchers: Matcher[] = [];
	for (let i = from; i < args.length; i += 2) {
		const g = toGrid(args[i], ctx);
		if (isCellError(g)) {
			return g;
		}
		if (grids.length > 0 && !sameShape(g, grids[0])) {
			return ERRORS.VALUE;
		}
		grids.push(g);
		matchers.push(createCriteriaMatcher(criteriaValue(args[i + 1]), ctx.date1904));
	}
	return { grids, matchers };
}

function allMatch(
	grids: readonly Grid[],
	matchers: readonly Matcher[],
	r: number,
	c: number,
): boolean {
	for (let k = 0; k < grids.length; k++) {
		if (!matchers[k](gridAt(grids[k], r, c))) {
			return false;
		}
	}
	return true;
}

const COUNTIF: FunctionDefinition = define({
	name: 'COUNTIF',
	minArgs: 2,
	maxArgs: 2,
	params: ['ref', 'value'],
	propagateErrors: false,
	impl: (args, ctx) => {
		const range = args[0];
		const matcher = createCriteriaMatcher(criteriaValue(args[1]), ctx.date1904);
		if (isRefValue(range) && (range.sheetEnd === undefined || range.sheetEnd === range.sheet)) {
			// Visit stored cells only; blanks are counted arithmetically.
			let matched = 0;
			let visited = 0;
			ctx.forEachCell(range, (v) => {
				visited++;
				if (matcher(v)) {
					matched++;
				}
			});
			return matcher(null) ? matched + cellCount(range) - visited : matched;
		}
		const g = toGrid(range, ctx);
		if (isCellError(g)) {
			return g;
		}
		let n = 0;
		for (let r = 0; r < g.rows; r++) {
			for (let c = 0; c < g.cols; c++) {
				if (matcher(gridAt(g, r, c))) {
					n++;
				}
			}
		}
		return n;
	},
});

const COUNTIFS: FunctionDefinition = define({
	name: 'COUNTIFS',
	minArgs: 2,
	maxArgs: 254,
	params: pairedParams([]),
	propagateErrors: false,
	impl: (args, ctx) => {
		const pairs = criteriaPairs(args, 0, ctx);
		if (isCellError(pairs)) {
			return pairs;
		}
		const { grids, matchers } = pairs;
		const shape = grids[0];
		const rows = Math.min(shape.rows, Math.max(...grids.map((g) => g.filledRows)));
		const cols = Math.min(shape.cols, Math.max(...grids.map((g) => g.filledCols)));
		let n = 0;
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				if (allMatch(grids, matchers, r, c)) {
					n++;
				}
			}
		}
		// Positions outside every filled area are blank in all ranges.
		const rest = shape.rows * shape.cols - rows * cols;
		if (rest > 0 && matchers.every((m) => m(null))) {
			n += rest;
		}
		return n;
	},
});

type Reducer = { add: (v: number) => void; result: () => FormulaValue };

/** Fold the numeric cells of `values` at positions where every criterion matches. */
function reduceMatches(
	values: Grid,
	grids: readonly Grid[],
	matchers: readonly Matcher[],
	reducer: Reducer,
): FormulaValue {
	const rows = Math.min(values.rows, values.filledRows);
	const cols = Math.min(values.cols, values.filledCols);
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			const v = gridAt(values, r, c);
			if (v === null || !allMatch(grids, matchers, r, c)) {
				continue;
			}
			if (typeof v === 'number') {
				reducer.add(v);
			} else if (isCellError(v)) {
				return v;
			}
		}
	}
	return reducer.result();
}

function averager(): Reducer {
	let sum = 0;
	let n = 0;
	return {
		add: (v) => {
			sum += v;
			n++;
		},
		result: () => (n === 0 ? ERRORS.DIV0 : sum / n),
	};
}

function extremer(largest: boolean): Reducer {
	let best = largest ? -Infinity : Infinity;
	let n = 0;
	return {
		add: (v) => {
			best = largest ? Math.max(best, v) : Math.min(best, v);
			n++;
		},
		result: () => (n === 0 ? 0 : best),
	};
}

const AVERAGEIF: FunctionDefinition = define({
	name: 'AVERAGEIF',
	minArgs: 2,
	maxArgs: 3,
	params: ['ref', 'value', 'ref'],
	propagateErrors: false,
	impl: (args, ctx) => {
		const range = toGrid(args[0], ctx);
		if (isCellError(range)) {
			return range;
		}
		const matcher = createCriteriaMatcher(criteriaValue(args[1]), ctx.date1904);
		let values = range;
		const avg = args[2];
		if (args.length > 2 && (isRefValue(avg) || isArrayValue(avg))) {
			// The average range is read from its top-left cell in the shape of `range`.
			const g = toGrid(isRefValue(avg) ? resizeRef(avg, range.rows, range.cols) : avg, ctx);
			if (isCellError(g)) {
				return g;
			}
			values = { ...g, rows: range.rows, cols: range.cols };
		} else if (args.length > 2 && isCellError(avg)) {
			return avg;
		}
		return reduceMatches(values, [range], [matcher], averager());
	},
});

function multiCriteria(name: string, makeReducer: () => Reducer): FunctionDefinition {
	return define({
		name,
		minArgs: 3,
		maxArgs: 255,
		params: pairedParams(['ref']),
		propagateErrors: false,
		impl: (args, ctx) => {
			const values = toGrid(args[0], ctx);
			if (isCellError(values)) {
				return values;
			}
			const pairs = criteriaPairs(args, 1, ctx);
			if (isCellError(pairs)) {
				return pairs;
			}
			if (!sameShape(values, pairs.grids[0])) {
				return ERRORS.VALUE;
			}
			return reduceMatches(values, pairs.grids, pairs.matchers, makeReducer());
		},
	});
}

export const CONDITIONAL_FUNCTIONS: FunctionDefinition[] = [
	COUNTIF,
	COUNTIFS,
	AVERAGEIF,
	multiCriteria('AVERAGEIFS', averager),
	multiCriteria('MAXIFS', () => extremer(true)),
	multiCriteria('MINIFS', () => extremer(false)),
];
