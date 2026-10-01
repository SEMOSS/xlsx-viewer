import { a1, keyCol, keyRow } from '../../address/cell-address';
import type { AstNode } from '../../formula/ast';
import { parseFormula } from '../../formula/language/parser';
import { getDefaultRegistry } from '../../formula/registry';
import type { CellValue } from '../../types/cell';
import type { Workbook } from '../../types/workbook';
import { isCellError } from '../../values/errors';
import { VOLATILE_FUNCTIONS } from '../analysis';
import { CalcEngine } from '../engine';

export interface Mismatch {
	sheet: string;
	cell: string;
	formula: string;
	cached: CellValue;
	computed: CellValue;
}

export interface Comparison {
	compared: number;
	/** Formulas skipped because they use functions not registered yet (or volatile ones). */
	skipped: number;
	mismatches: Mismatch[];
}

/** Functions whose results depend on the environment rather than the workbook. */
const ENVIRONMENT = new Set(['CELL', 'INFO']);

function collect(node: AstNode, fns: Set<string>, names: Set<string>): void {
	switch (node.type) {
		case 'function':
			fns.add(node.name);
			node.args.forEach((a) => collect(a, fns, names));
			return;
		case 'name':
			names.add(node.name.toUpperCase());
			return;
		case 'call':
			collect(node.callee, fns, names);
			node.args.forEach((a) => collect(a, fns, names));
			return;
		case 'unary':
		case 'percent':
			collect(node.operand, fns, names);
			return;
		case 'binary':
			collect(node.left, fns, names);
			collect(node.right, fns, names);
			return;
		default:
			return;
	}
}

/**
 * Every function the formula calls (directly or through defined names) is
 * implemented and deterministic. LET / LAMBDA-local functions count as
 * unknown, so such formulas are conservatively skipped.
 */
function computable(formula: string, wb: Workbook): boolean {
	const registry = getDefaultRegistry();
	const defs = new Map(wb.definedNames.map((d) => [d.name.toUpperCase(), d.formula]));
	const seen = new Set<string>();
	const pending = [formula];
	while (pending.length > 0) {
		let ast: AstNode;
		try {
			ast = parseFormula(pending.pop()!);
		} catch {
			return false;
		}
		const fns = new Set<string>();
		const names = new Set<string>();
		collect(ast, fns, names);
		for (const fn of fns) {
			if (VOLATILE_FUNCTIONS.has(fn) || ENVIRONMENT.has(fn)) {
				return false;
			}
			if (registry.has(fn) || fn === 'ANCHORARRAY' || fn === 'SINGLE') {
				continue;
			}
			if (!defs.has(fn)) {
				return false;
			}
			names.add(fn);
		}
		for (const name of names) {
			const def = defs.get(name);
			if (def !== undefined && !seen.has(name)) {
				seen.add(name);
				pending.push(def);
			}
		}
	}
	return true;
}

/** Values agree: numbers within 1e-9 relative, everything else exactly (errors by code). */
export function sameValue(a: CellValue, b: CellValue): boolean {
	if (typeof a === 'number' && typeof b === 'number') {
		return a === b || Math.abs(a - b) <= 1e-9 * Math.max(Math.abs(a), Math.abs(b), 1e-300);
	}
	if (isCellError(a) || isCellError(b)) {
		return isCellError(a) && isCellError(b) && a.code === b.code;
	}
	if ((a === null && b === '') || (a === '' && b === null)) {
		return true;
	}
	return a === b;
}

/**
 * Recalculate a loaded workbook and compare every computable formula cell
 * (and the cells its arrays fill) with the value Excel cached in the file.
 */
export function compareWithCachedValues(wb: Workbook): Comparison {
	const cached = wb.sheets.map((s) => new Map([...s.cells].map(([k, c]) => [k, c.value])));
	const engine = new CalcEngine(wb);
	engine.rebuild();
	engine.recalculateAll();
	const result: Comparison = { compared: 0, skipped: 0, mismatches: [] };
	wb.sheets.forEach((sheet, s) => {
		for (const [key, cell] of sheet.cells) {
			const anchor =
				cell.formula !== undefined
					? cell
					: cell.arrayAnchor !== undefined
						? sheet.cells.get(cell.arrayAnchor)
						: undefined;
			const formula = anchor?.formula;
			if (formula === undefined || !cached[s].has(key)) {
				continue;
			}
			if (anchor?.stale || !computable(formula, wb)) {
				result.skipped++;
				continue;
			}
			result.compared++;
			const before = cached[s].get(key) ?? null;
			if (!sameValue(before, cell.value)) {
				result.mismatches.push({
					sheet: sheet.name,
					cell: a1(keyRow(key), keyCol(key)),
					formula,
					cached: before,
					computed: cell.value,
				});
			}
		}
	});
	return result;
}
