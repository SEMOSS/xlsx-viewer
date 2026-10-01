import { a1, keyCol, keyRow } from '../address/cell-address';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Workbook } from '../types/workbook';
import { isCellError } from '../values/errors';
import type { FormulaCache } from './formula-cache';
import type { DependencyGraph } from './graph';
import { FormulaNode, type FormulaMode } from './node';
import { resolveDependencies, type ResolveContext } from './resolve';
import type { CellRef } from './types';

/** What `createNode` needs from the engine. */
export interface BuildContext extends ResolveContext {
	readonly workbook: Workbook;
	readonly graph: DependencyGraph;
	readonly cache: FormulaCache;
}

/** Problems found while building, reported once as workbook warnings. */
export class BuildIssues {
	/** Unsupported function name -> first cell using it. */
	readonly functions = new Map<string, CellRef>();
	external: CellRef | null = null;
	parse: CellRef | null = null;
}

/** Evaluation mode of a formula cell. */
export function modeOf(cell: Cell): FormulaMode {
	if (cell.array === undefined) {
		return 'legacy';
	}
	return cell.array.dynamic ? 'dynamic' : 'cse';
}

function isMultiCell(r: RangeAddress): boolean {
	return r.top !== r.bottom || r.left !== r.right;
}

/** The CSE block of an anchor (normalised; null when it does not start at the anchor). */
function cseBlock(cell: Cell, row: number, col: number): RangeAddress | null {
	const r = cell.array?.ref;
	if (r === undefined || !isMultiCell(r) || r.top !== row || r.left !== col) {
		return null;
	}
	return { top: r.top, left: r.left, bottom: r.bottom, right: r.right };
}

/**
 * Create the graph node of a formula cell and register its precedents.
 * Formulas loaded from a file (`fromFile`) that use unsupported functions,
 * external references or unparsable text keep their cached value: the node
 * is frozen and the cell flagged `stale`.
 */
export function createNode(
	ctx: BuildContext,
	sheet: number,
	key: number,
	cell: Cell,
	fromFile: boolean,
	issues: BuildIssues | null,
	bulk = false,
): FormulaNode {
	const row = keyRow(key);
	const col = keyCol(key);
	const formula = cell.formula ?? '';
	const parsed = ctx.cache.get(formula);
	const mode = modeOf(cell);
	const node = new FormulaNode(sheet, row, col, key, formula, parsed, mode);
	const deps =
		parsed.analysis !== null ? resolveDependencies(parsed.analysis, sheet, row, col, ctx) : null;
	node.volatile = deps?.volatile ?? false;
	node.subtotal = parsed.analysis?.subtotal ?? false;
	ctx.graph.add(node, deps?.cells ?? [], deps?.ranges ?? [], bulk);
	if (mode === 'cse') {
		ctx.graph.setCse(node, cseBlock(cell, row, col));
	}
	if (mode === 'dynamic') {
		node.arrayResult = true;
		const ref = cell.array!.ref;
		const spilled = !(isCellError(cell.value) && cell.value.code === '#SPILL!');
		if (spilled && isMultiCell(ref) && ref.top === row && ref.left === col) {
			ctx.graph.setSpill(node, {
				top: ref.top,
				left: ref.left,
				bottom: ref.bottom,
				right: ref.right,
			});
		}
	}
	const unsupported = deps?.unsupported ?? [];
	const external = deps?.external ?? false;
	const unreadable = parsed.ast === null || deps?.brokenName === true;
	if (fromFile && cell.value !== null && (unreadable || unsupported.length > 0 || external)) {
		node.frozen = true;
		if (cell.stale !== true) {
			ctx.workbook.sheets[sheet].cells.set(key, { ...cell, stale: true });
		}
		if (issues !== null) {
			const at = { sheet, row, col };
			for (const name of unsupported) {
				if (!issues.functions.has(name)) {
					issues.functions.set(name, at);
				}
			}
			if (external) {
				issues.external ??= at;
			}
			if (unreadable) {
				issues.parse ??= at;
			}
		}
	}
	return node;
}

/** Add the build issues to `workbook.warnings`, once per function name / kind. */
export function reportIssues(workbook: Workbook, issues: BuildIssues): void {
	const existing = new Set(workbook.warnings.map((w) => `${w.code}|${w.message}`));
	const push = (code: string, message: string, at: CellRef): void => {
		if (existing.has(`${code}|${message}`)) {
			return;
		}
		existing.add(`${code}|${message}`);
		workbook.warnings.push({
			code,
			message,
			sheet: workbook.sheets[at.sheet]?.name,
			cell: a1(at.row, at.col),
		});
	};
	for (const [name, at] of issues.functions) {
		push(
			'unsupported-function',
			`The function ${name} is not supported; formulas using it keep the values saved in the file`,
			at,
		);
	}
	if (issues.external !== null) {
		push(
			'external-reference',
			'Formulas referring to other workbooks keep the values saved in the file',
			issues.external,
		);
	}
	if (issues.parse !== null) {
		push(
			'formula-parse-error',
			'Some formulas could not be read; they keep the values saved in the file',
			issues.parse,
		);
	}
}
