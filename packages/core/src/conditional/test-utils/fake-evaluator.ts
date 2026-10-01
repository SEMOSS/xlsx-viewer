import { cellKey, keyCol, keyRow, lettersToColumn } from '../../address/cell-address';
import { compareScalars, toNumber, toText } from '../../formula/coerce';
import type { FormulaValue, RefValue, Scalar } from '../../formula/types';
import { isArrayValue, isRefValue, makeArray } from '../../formula/types';
import { findSheetIndex } from '../../sheet/workbook';
import type { Workbook } from '../../types/workbook';
import { cellError, ERRORS, isCellError, parseErrorCode } from '../../values/errors';
import type { FormulaEvaluator } from '../formula-eval';

/**
 * A deliberately tiny formula evaluator for tests: literals, A1 references
 * and ranges (optionally sheet-qualified, returned as `RefValue`s like the
 * real engine), defined names, arithmetic, `&`, comparisons and a handful of
 * functions (ROW, COLUMN, MOD, AND, OR, NOT, IF, ISNUMBER, ISBLANK, ISERROR,
 * LEN, SUM, MAX, MIN, COUNT, COUNTIF, ARR (row array of its arguments), NA,
 * THROW (throws)). `calls` counts invocations.
 */
export type FakeEvaluator = FormulaEvaluator & { calls: number; formulas: string[] };

type Token =
	| { t: 'num'; v: number }
	| { t: 'str'; v: string }
	| { t: 'err'; v: string }
	| { t: 'ref'; sheet?: string; a: string; b?: string }
	| { t: 'id'; v: string }
	| { t: 'op'; v: string };

const PATTERNS: Array<[RegExp, (m: RegExpExecArray) => Token | null]> = [
	[/^\s+/, () => null],
	[/^"(?:[^"]|"")*"/, (m) => ({ t: 'str', v: m[0].slice(1, -1).replace(/""/g, '"') })],
	[/^#(?:N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!)/, (m) => ({ t: 'err', v: m[0] })],
	[
		/^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_]*))!)?(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?(?![A-Za-z0-9_(])/,
		(m) => ({ t: 'ref', sheet: m[1]?.replace(/''/g, "'") ?? m[2], a: m[3], b: m[4] }),
	],
	[/^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/, (m) => ({ t: 'num', v: Number(m[0]) })],
	[/^[A-Za-z_][A-Za-z0-9_.]*/, (m) => ({ t: 'id', v: m[0] })],
	[/^(?:<>|<=|>=|[-+*/^&=<>(),%])/, (m) => ({ t: 'op', v: m[0] })],
];

function tokenize(src: string): Token[] {
	const out: Token[] = [];
	let rest = src;
	outer: while (rest.length) {
		for (const [re, make] of PATTERNS) {
			const m = re.exec(rest);
			if (m) {
				const tok = make(m);
				if (tok) {
					out.push(tok);
				}
				rest = rest.slice(m[0].length);
				continue outer;
			}
		}
		throw new Error(`fake evaluator: cannot tokenize "${rest}"`);
	}
	return out;
}

function parseCorner(text: string): { row: number; col: number } {
	const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(text)!;
	return { row: Number(m[2]) - 1, col: lettersToColumn(m[1]) };
}

class Evaluation {
	private pos = 0;
	private readonly toks: Token[];

	constructor(
		private readonly wb: Workbook,
		src: string,
		private readonly sheet: number,
		private readonly row: number,
		private readonly col: number,
		private readonly depth = 0,
	) {
		this.toks = tokenize(src);
	}

	run(): FormulaValue {
		const v = this.comparison();
		if (this.pos < this.toks.length) {
			throw new Error('fake evaluator: trailing tokens');
		}
		return v;
	}

	private peekOp(...ops: string[]): string | undefined {
		const tok = this.toks[this.pos];
		return tok?.t === 'op' && ops.includes(tok.v) ? tok.v : undefined;
	}

	private scalar(v: FormulaValue): Scalar {
		if (isRefValue(v)) {
			return this.wb.sheets[v.sheet]?.cells.get(cellKey(v.range.top, v.range.left))?.value ?? null;
		}
		if (isArrayValue(v)) {
			return v.values[0][0];
		}
		return v as Scalar;
	}

	private binary(
		next: () => FormulaValue,
		ops: string[],
		apply: (op: string, a: Scalar, b: Scalar) => Scalar,
	): FormulaValue {
		let left = next();
		for (let op = this.peekOp(...ops); op; op = this.peekOp(...ops)) {
			this.pos++;
			const a = this.scalar(left);
			const b = this.scalar(next());
			left = isCellError(a) ? a : isCellError(b) ? b : apply(op, a, b);
		}
		return left;
	}

	private comparison(): FormulaValue {
		return this.binary(
			() => this.concat(),
			['=', '<>', '<', '>', '<=', '>='],
			(op, a, b) => {
				const c = compareScalars(a, b);
				return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[
					op
				]!;
			},
		);
	}

	private concat(): FormulaValue {
		return this.binary(
			() => this.additive(),
			['&'],
			(_op, a, b) => `${toText(a)}${toText(b)}`,
		);
	}

	private additive(): FormulaValue {
		return this.binary(
			() => this.multiplicative(),
			['+', '-'],
			(op, a, b) => {
				const x = toNumber(a);
				const y = toNumber(b);
				return typeof x !== 'number' ? x : typeof y !== 'number' ? y : op === '+' ? x + y : x - y;
			},
		);
	}

	private multiplicative(): FormulaValue {
		return this.binary(
			() => this.unary(),
			['*', '/'],
			(op, a, b) => {
				const x = toNumber(a);
				const y = toNumber(b);
				if (typeof x !== 'number') return x;
				if (typeof y !== 'number') return y;
				return op === '*' ? x * y : y === 0 ? ERRORS.DIV0 : x / y;
			},
		);
	}

	private unary(): FormulaValue {
		if (this.peekOp('-')) {
			this.pos++;
			const n = toNumber(this.scalar(this.unary()));
			return typeof n === 'number' ? -n : n;
		}
		const v = this.primary();
		if (this.peekOp('%')) {
			this.pos++;
			const n = toNumber(this.scalar(v));
			return typeof n === 'number' ? n / 100 : n;
		}
		return v;
	}

	private primary(): FormulaValue {
		const tok = this.toks[this.pos++];
		if (!tok) throw new Error('fake evaluator: unexpected end');
		switch (tok.t) {
			case 'num':
			case 'str':
				return tok.v;
			case 'err':
				return cellError(parseErrorCode(tok.v)!);
			case 'ref': {
				const sheet = tok.sheet === undefined ? this.sheet : findSheetIndex(this.wb, tok.sheet);
				if (sheet < 0) return ERRORS.REF;
				const a = parseCorner(tok.a);
				const b = tok.b ? parseCorner(tok.b) : a;
				return {
					kind: 'ref',
					sheet,
					range: {
						top: Math.min(a.row, b.row),
						left: Math.min(a.col, b.col),
						bottom: Math.max(a.row, b.row),
						right: Math.max(a.col, b.col),
					},
				};
			}
			case 'op': {
				if (tok.v !== '(') throw new Error(`fake evaluator: unexpected ${tok.v}`);
				const v = this.comparison();
				this.pos++;
				return v;
			}
			case 'id':
				return this.identifier(tok.v);
		}
	}

	private identifier(name: string): FormulaValue {
		const upper = name.toUpperCase();
		if (!this.peekOp('(')) {
			if (upper === 'TRUE' || upper === 'FALSE') return upper === 'TRUE';
			const dn = this.wb.definedNames.find((d) => d.name.toUpperCase() === upper);
			if (!dn || this.depth > 5) return ERRORS.NAME;
			return new Evaluation(
				this.wb,
				dn.formula,
				this.sheet,
				this.row,
				this.col,
				this.depth + 1,
			).run();
		}
		this.pos++;
		const args: FormulaValue[] = [];
		while (!this.peekOp(')')) {
			if (this.peekOp(',')) {
				args.push(null);
				this.pos++;
				continue;
			}
			args.push(this.comparison());
			if (this.peekOp(',')) this.pos++;
		}
		this.pos++;
		return this.call(upper, args);
	}

	private values(args: FormulaValue[]): Scalar[] {
		const out: Scalar[] = [];
		for (const arg of args) {
			if (isRefValue(arg)) {
				const cells = this.wb.sheets[arg.sheet]?.cells;
				for (const [key, cell] of cells ?? []) {
					const r = keyRow(key);
					const c = keyCol(key);
					if (
						r >= arg.range.top &&
						r <= arg.range.bottom &&
						c >= arg.range.left &&
						c <= arg.range.right
					) {
						out.push(cell.value);
					}
				}
			} else {
				out.push(this.scalar(arg));
			}
		}
		return out;
	}

	private call(name: string, args: FormulaValue[]): FormulaValue {
		const s = (i: number) => this.scalar(args[i] ?? null);
		const nums = () => this.values(args).filter((v): v is number => typeof v === 'number');
		switch (name) {
			case 'ROW':
				return (args[0] && isRefValue(args[0]) ? (args[0] as RefValue).range.top : this.row) + 1;
			case 'COLUMN':
				return (args[0] && isRefValue(args[0]) ? (args[0] as RefValue).range.left : this.col) + 1;
			case 'MOD': {
				const a = toNumber(s(0));
				const b = toNumber(s(1));
				return typeof a === 'number' && typeof b === 'number'
					? a - b * Math.floor(a / b)
					: ERRORS.VALUE;
			}
			case 'AND':
				return args.every((_, i) => s(i) === true || (typeof s(i) === 'number' && s(i) !== 0));
			case 'OR':
				return args.some((_, i) => s(i) === true || (typeof s(i) === 'number' && s(i) !== 0));
			case 'NOT':
				return !(s(0) === true || (typeof s(0) === 'number' && s(0) !== 0));
			case 'IF':
				return s(0) === true || (typeof s(0) === 'number' && s(0) !== 0)
					? (args[1] ?? true)
					: (args[2] ?? false);
			case 'ISNUMBER':
				return typeof s(0) === 'number';
			case 'ISBLANK':
				return s(0) === null;
			case 'ISERROR':
				return isCellError(s(0));
			case 'LEN': {
				const t = toText(s(0));
				return typeof t === 'string' ? t.length : t;
			}
			case 'SUM':
				return nums().reduce((a, b) => a + b, 0);
			case 'MAX':
				return nums().length ? Math.max(...nums()) : 0;
			case 'MIN':
				return nums().length ? Math.min(...nums()) : 0;
			case 'COUNT':
				return nums().length;
			case 'COUNTIF': {
				const target = s(1);
				return this.values([args[0]]).filter(
					(v) => !isCellError(v) && v !== null && compareScalars(v, target) === 0,
				).length;
			}
			case 'ARR':
				return makeArray([args.map((_, i) => s(i))]);
			case 'NA':
				return ERRORS.NA;
			case 'THROW':
				throw new Error('boom');
			default:
				return ERRORS.NAME;
		}
	}
}

/** Create a fake evaluator bound to a workbook. */
export function createFakeEvaluator(workbook: Workbook): FakeEvaluator {
	const evaluate = ((formula: string, sheetIndex: number, row: number, col: number) => {
		evaluate.calls++;
		evaluate.formulas.push(formula);
		return new Evaluation(workbook, formula, sheetIndex, row, col).run();
	}) as FakeEvaluator;
	evaluate.calls = 0;
	evaluate.formulas = [];
	return evaluate;
}
