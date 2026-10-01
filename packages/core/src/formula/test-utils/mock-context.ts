import { parseRange } from '../../address/range';
import type { RangeAddress } from '../../types/address';
import { cellError, ERRORS } from '../../values/errors';
import type { AstNode } from '../ast';
import { invokeFunction, type ArrayMode } from '../invoke';
import type {
	ArrayValue,
	FormulaValue,
	FunctionArg,
	FunctionContext,
	FunctionDefinition,
	LambdaValue,
	RefValue,
	Scalar,
} from '../types';

/**
 * In-memory FunctionContext for unit-testing functions without a workbook.
 *
 * ```ts
 * const ctx = new MockContext({ A1: 1, A2: 2, B1: 'x' });
 * call(SUM, [ctx.ref('A1:A2'), 3], ctx); // 6
 * ```
 */
export class MockContext implements FunctionContext {
	readonly sheet = 0;
	row = 0;
	col = 0;
	date1904 = false;
	/** sheet -> "row,col" -> value */
	private readonly data: Array<Map<string, Scalar>> = [new Map()];
	private readonly formulas = new Map<string, string>();
	private readonly hiddenRows = new Set<string>();
	private readonly subtotalCells = new Set<string>();
	readonly sheetNames: string[] = ['Sheet1'];
	fixedNow = 45292.5; // 2024-01-01 12:00
	randomValues: number[] = [];
	scope: ReadonlyMap<string, FormulaValue> = new Map();

	constructor(
		cells: Record<string, Scalar> = {},
		extraSheets: Record<string, Record<string, Scalar>> = {},
	) {
		this.setCells(cells, 0);
		for (const [name, sheetCells] of Object.entries(extraSheets)) {
			this.sheetNames.push(name);
			this.data.push(new Map());
			this.setCells(sheetCells, this.sheetNames.length - 1);
		}
	}

	/** Set cells by A1 address on a sheet. */
	setCells(cells: Record<string, Scalar>, sheet = 0): void {
		for (const [addr, value] of Object.entries(cells)) {
			const r = parseRange(addr);
			if (!r) {
				throw new Error(`Bad address ${addr}`);
			}
			this.data[sheet].set(`${r.top},${r.left}`, value);
		}
	}

	setFormula(addr: string, formula: string, sheet = 0): void {
		const r = parseRange(addr)!;
		this.formulas.set(`${sheet}:${r.top},${r.left}`, formula);
	}

	hideRow(row: number, sheet = 0): void {
		this.hiddenRows.add(`${sheet}:${row}`);
	}

	markSubtotal(addr: string, sheet = 0): void {
		const r = parseRange(addr)!;
		this.subtotalCells.add(`${sheet}:${r.top},${r.left}`);
	}

	/** Position the "formula cell" (for implicit intersection, ROW(), ...). */
	at(addr: string): this {
		const r = parseRange(addr)!;
		this.row = r.top;
		this.col = r.left;
		return this;
	}

	/** Reference value for an A1 range on a sheet. */
	ref(a1: string, sheet = 0): RefValue {
		const range = parseRange(a1);
		if (!range) {
			throw new Error(`Bad range ${a1}`);
		}
		return { kind: 'ref', sheet, range };
	}

	getCellValue(sheet: number, row: number, col: number): Scalar {
		return this.data[sheet]?.get(`${row},${col}`) ?? null;
	}

	resolveRef(ref: RefValue): ArrayValue {
		const r = this.clamp(ref);
		const values: Scalar[][] = [];
		for (let row = r.top; row <= r.bottom; row++) {
			const line: Scalar[] = [];
			for (let col = r.left; col <= r.right; col++) {
				line.push(this.getCellValue(ref.sheet, row, col));
			}
			values.push(line);
		}
		return { kind: 'array', rows: values.length, cols: values[0]?.length ?? 0, values };
	}

	forEachCell(
		ref: RefValue,
		visit: (value: Scalar, row: number, col: number, sheet: number) => void,
	): void {
		const r = this.clamp(ref);
		const entries: Array<[number, number, Scalar]> = [];
		for (const [key, value] of this.data[ref.sheet] ?? []) {
			const [row, col] = key.split(',').map(Number);
			if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right && value !== null) {
				entries.push([row, col, value]);
			}
		}
		entries.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
		for (const [row, col, value] of entries) {
			visit(value, row, col, ref.sheet);
		}
	}

	private clamp(ref: RefValue): RangeAddress {
		const used = this.usedRange(ref.sheet);
		const r = ref.range;
		if (!used) {
			return {
				top: r.top,
				left: r.left,
				bottom: Math.min(r.bottom, r.top),
				right: Math.min(r.right, r.left),
			};
		}
		return {
			top: r.top,
			left: r.left,
			bottom: Math.min(r.bottom, Math.max(used.bottom, r.top)),
			right: Math.min(r.right, Math.max(used.right, r.left)),
		};
	}

	getFormula(sheet: number, row: number, col: number): string | undefined {
		return this.formulas.get(`${sheet}:${row},${col}`);
	}

	isRowHidden(sheet: number, row: number): boolean {
		return this.hiddenRows.has(`${sheet}:${row}`);
	}

	isSubtotalCell(sheet: number, row: number, col: number): boolean {
		return this.subtotalCells.has(`${sheet}:${row},${col}`);
	}

	usedRange(sheet: number): RangeAddress | null {
		let bottom = -1;
		let right = -1;
		for (const key of this.data[sheet]?.keys() ?? []) {
			const [row, col] = key.split(',').map(Number);
			bottom = Math.max(bottom, row);
			right = Math.max(right, col);
		}
		return bottom < 0 ? null : { top: 0, left: 0, bottom, right };
	}

	sheetCount(): number {
		return this.sheetNames.length;
	}

	sheetName(index: number): string {
		return this.sheetNames[index] ?? '';
	}

	sheetIndex(name: string): number {
		return this.sheetNames.findIndex((n) => n.toLowerCase() === name.toLowerCase());
	}

	evaluate(node: AstNode): FormulaValue {
		switch (node.type) {
			case 'number':
			case 'string':
			case 'boolean':
				return node.value;
			case 'error':
				return cellError(node.code);
			case 'missing':
				return null;
			default:
				throw new Error('MockContext.evaluate only supports literal nodes');
		}
	}

	evaluateWith(node: AstNode, _bindings: ReadonlyMap<string, FormulaValue>): FormulaValue {
		return this.evaluate(node);
	}

	callLambda(_fn: LambdaValue, _args: FormulaValue[]): FormulaValue {
		return ERRORS.CALC;
	}

	now(): number {
		return this.fixedNow;
	}

	random(): number {
		return this.randomValues.length ? this.randomValues.shift()! : 0.5;
	}
}

/** Invoke a function definition the way the evaluator does (argument preparation and lifting). */
export function call(
	def: FunctionDefinition,
	args: FunctionArg[],
	ctx: FunctionContext = new MockContext(),
	mode: ArrayMode = 'dynamic',
): FormulaValue {
	return invokeFunction(def, args, ctx, mode);
}

/** Array literal helper: `arr([1, 2], [3, 4])`. */
export function arr(...rows: Scalar[][]): ArrayValue {
	return { kind: 'array', rows: rows.length, cols: rows[0]?.length ?? 0, values: rows };
}

/** Find a definition by name in a module's list (throws when missing). */
export function fn(defs: readonly FunctionDefinition[], name: string): FunctionDefinition {
	const def = defs.find((d) => d.name === name);
	if (!def) {
		throw new Error(`Function ${name} not found`);
	}
	return def;
}
