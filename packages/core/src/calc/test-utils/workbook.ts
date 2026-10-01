import { a1, cellKey } from '../../address/cell-address';
import { parseRange } from '../../address/range';
import { createWorkbook } from '../../sheet/workbook';
import type { FormulaValue } from '../../formula/types';
import { isArrayValue, isLambdaValue, isMultiRefValue, isRefValue } from '../../formula/types';
import type { Cell, CellValue } from '../../types/cell';
import type { Workbook } from '../../types/workbook';
import { CalcEngine } from '../engine';
import type { CalcEngineOptions, CellRef, RecalcResult } from '../types';

/**
 * Test helpers: build workbooks from compact cell maps and edit them the
 * way the editor does.
 *
 * - `'=A1+1'` is a formula typed by the user (dynamic-array mode, no value);
 * - `{ f: 'A1+1', v: 2 }` a legacy formula loaded with a cached value;
 * - `{ f: 'A1:A3*2', cse: 'B1:B3' }` a legacy CSE array formula;
 * - `{ f: 'SEQUENCE(3)', dynamic: true, v: 1 }` a loaded dynamic-array formula;
 * - anything else is a constant.
 */
export type CellSpec =
	| CellValue
	| {
			f: string;
			v?: CellValue;
			cse?: string;
			dynamic?: boolean;
			spill?: string;
			style?: number;
	  };

function position(addr: string): { row: number; col: number } {
	const r = parseRange(addr);
	if (r === null) {
		throw new Error(`Bad address ${addr}`);
	}
	return { row: r.top, col: r.left };
}

function toCell(spec: CellSpec, addr: string): Cell {
	if (typeof spec === 'string' && spec.startsWith('=')) {
		const { row, col } = position(addr);
		return {
			value: null,
			formula: spec.slice(1),
			array: { ref: { top: row, left: col, bottom: row, right: col }, dynamic: true },
		};
	}
	if (spec !== null && typeof spec === 'object' && 'f' in spec) {
		const cell: Cell = { value: spec.v ?? null, formula: spec.f };
		if (spec.cse !== undefined) {
			cell.array = { ref: parseRange(spec.cse)!, dynamic: false };
		} else if (spec.dynamic) {
			const { row, col } = position(addr);
			cell.array = {
				ref: spec.spill
					? parseRange(spec.spill)!
					: { top: row, left: col, bottom: row, right: col },
				dynamic: true,
			};
		}
		if (spec.style !== undefined) {
			cell.style = spec.style;
		}
		return cell;
	}
	return { value: spec as CellValue };
}

/** A workbook with the given sheets (in order) and cells. */
export function makeWorkbook(sheets: Record<string, Record<string, CellSpec>>): Workbook {
	const names = Object.keys(sheets);
	const wb = createWorkbook({ sheetNames: names.length ? names : ['Sheet1'] });
	names.forEach((name, i) => {
		for (const [addr, spec] of Object.entries(sheets[name])) {
			const { row, col } = position(addr);
			wb.sheets[i].cells.set(cellKey(row, col), toCell(spec, addr));
		}
	});
	return wb;
}

/** Split `Sheet2!B3` / `B3` into a sheet index and position. */
export function locate(wb: Workbook, ref: string): CellRef {
	const bang = ref.lastIndexOf('!');
	const sheetName = bang >= 0 ? ref.slice(0, bang).replace(/^'|'$/g, '') : wb.sheets[0].name;
	const sheet = wb.sheets.findIndex((s) => s.name.toLowerCase() === sheetName.toLowerCase());
	if (sheet < 0) {
		throw new Error(`Unknown sheet in ${ref}`);
	}
	return { sheet, ...position(bang >= 0 ? ref.slice(bang + 1) : ref) };
}

/** The stored cell at a reference. */
export function cellAt(wb: Workbook, ref: string): Cell | undefined {
	const { sheet, row, col } = locate(wb, ref);
	return wb.sheets[sheet].cells.get(cellKey(row, col));
}

/** The value at a reference (null when blank). */
export function valueAt(wb: Workbook, ref: string): CellValue {
	return cellAt(wb, ref)?.value ?? null;
}

/** Values of a block, row by row. */
export function valuesAt(wb: Workbook, range: string): CellValue[][] {
	const bang = range.lastIndexOf('!');
	const prefix = bang >= 0 ? range.slice(0, bang + 1) : '';
	const r = parseRange(bang >= 0 ? range.slice(bang + 1) : range)!;
	const out: CellValue[][] = [];
	for (let row = r.top; row <= r.bottom; row++) {
		const line: CellValue[] = [];
		for (let col = r.left; col <= r.right; col++) {
			line.push(cellAt(wb, `${prefix}${a1(row, col)}`)?.value ?? null);
		}
		out.push(line);
	}
	return out;
}

/** Store a cell the way the editor does (formulas as dynamic arrays) and notify the engine. */
export function edit(engine: CalcEngine, ref: string, spec: CellSpec | undefined): RecalcResult {
	const wb = engine.workbook;
	const at = locate(wb, ref);
	const key = cellKey(at.row, at.col);
	const cells = wb.sheets[at.sheet].cells;
	if (spec === undefined) {
		const old = cells.get(key);
		if (old?.style !== undefined) {
			cells.set(key, { value: null, style: old.style });
		} else {
			cells.delete(key);
		}
	} else {
		const addr = ref.slice(ref.lastIndexOf('!') + 1);
		const cell = toCell(spec, addr);
		const old = cells.get(key);
		if (old?.style !== undefined && cell.style === undefined) {
			cell.style = old.style;
		}
		cells.set(key, cell);
	}
	return engine.onCellsChanged([at]);
}

/** A workbook plus an engine that has been rebuilt and fully recalculated. */
export function calculated(
	sheets: Record<string, Record<string, CellSpec>>,
	options: CalcEngineOptions = {},
	setup?: (wb: Workbook) => void,
): { wb: Workbook; engine: CalcEngine; result: RecalcResult } {
	const wb = makeWorkbook(sheets);
	setup?.(wb);
	const engine = new CalcEngine(wb, options);
	engine.rebuild();
	const result = engine.recalculateAll();
	return { wb, engine, result };
}

/** Format changed cells as `Sheet1!A1` strings for compact assertions. */
export function refs(wb: Workbook, list: readonly CellRef[]): string[] {
	return list.map((c) => `${wb.sheets[c.sheet].name}!${a1(c.row, c.col)}`);
}

/** A formula result as plain data: scalars as is, references and arrays as rows of values. */
export function plain(engine: CalcEngine, v: FormulaValue): CellValue | CellValue[][] | string {
	if (isRefValue(v)) {
		if (v.sheetEnd !== undefined && v.sheetEnd !== v.sheet) {
			return `3D ${v.sheet}:${v.sheetEnd}`;
		}
		const r = v.range;
		const ws = engine.workbook.sheets[v.sheet];
		if (r.top === r.bottom && r.left === r.right) {
			return ws.cells.get(cellKey(r.top, r.left))?.value ?? null;
		}
		const out: CellValue[][] = [];
		for (let row = r.top; row <= r.bottom; row++) {
			const line: CellValue[] = [];
			for (let col = r.left; col <= r.right; col++) {
				line.push(ws.cells.get(cellKey(row, col))?.value ?? null);
			}
			out.push(line);
		}
		return out;
	}
	if (isMultiRefValue(v)) {
		return `refs(${v.refs.length})`;
	}
	if (isArrayValue(v)) {
		return v.values;
	}
	if (isLambdaValue(v)) {
		return 'lambda';
	}
	return v as CellValue;
}

/** Evaluate a formula in a cell (default Z100 of the first sheet) and return plain data. */
export function evalIn(
	engine: CalcEngine,
	formula: string,
	at = 'Z100',
	dynamicArrays = true,
): CellValue | CellValue[][] | string {
	const { sheet, row, col } = locate(engine.workbook, at);
	return plain(engine, engine.evaluateFormula(formula, sheet, row, col, { dynamicArrays }));
}
