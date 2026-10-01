import type { Evaluator } from '../formula/evaluator/evaluator';
import { arrayCellResult, legacyCellResult } from '../formula/evaluator/results';
import type { ArrayValue, FormulaValue, Scalar } from '../formula/types';
import { isArrayValue, isMultiRefValue, isRefValue } from '../formula/types';
import type { Workbook } from '../types/workbook';
import { ERRORS } from '../values/errors';
import type { DependencyGraph } from './graph';
import { PendingCells, type WorkbookHost } from './host';
import type { FormulaNode } from './node';
import { dirtyPrecedents, readsRange } from './precedents';
import type { CellRef } from './types';
import { ChangeSet, type ResultWriter } from './writer';

/** Evaluations of one node within a recalculation beyond which it is treated as circular. */
const MAX_EVALS = 32;
/** Follow-up rounds for formulas re-dirtied by spills that changed shape. */
const MAX_ROUNDS = 64;

interface TarjanFrame {
	node: FormulaNode;
	succ: FormulaNode[];
	i: number;
	self: boolean;
}

function isMultiCell(v: FormulaValue): boolean {
	if (isRefValue(v)) {
		return v.range.top !== v.range.bottom || v.range.left !== v.range.right;
	}
	return isMultiRefValue(v);
}

function byPosition(a: FormulaNode, b: FormulaNode): number {
	return a.sheet - b.sheet || a.row - b.row || a.col - b.col;
}

/**
 * Runs a recalculation: marks the dirty closure, orders it with Tarjan's
 * algorithm over the static precedents (strongly connected components are
 * circular references), evaluates each cell, and handles cells read
 * dynamically (INDIRECT, OFFSET, spills) by suspending and retrying the
 * reader. Everything is iterative, so long chains do not grow the stack.
 */
export class Scheduler {
	private recalcId = 0;
	private stamp = 0;
	private index = 0;
	private late: FormulaNode[] = [];
	private marked: FormulaNode[] = [];
	private circular = new ChangeSet();
	inPass = false;

	constructor(
		private readonly workbook: Workbook,
		private readonly graph: DependencyGraph,
		private readonly host: WorkbookHost,
		private readonly writer: ResultWriter,
		private readonly evaluator: Evaluator,
	) {}

	/**
	 * Recalculate `roots` and everything that depends on them or on the
	 * `changed` cells. `complete`: the roots already are every formula (full
	 * recalculation), so dependents need not be looked up.
	 */
	run(roots: Iterable<FormulaNode>, changed: readonly CellRef[], complete = false): CellRef[] {
		this.recalcId++;
		this.circular = new ChangeSet();
		this.marked = [];
		this.late = [];
		this.host.beginPass();
		this.inPass = true;
		try {
			const list: FormulaNode[] = [];
			this.markClosure(roots, changed, list, !complete);
			this.process(list);
			for (let round = 0; this.late.length > 0 && round < MAX_ROUNDS; round++) {
				const next = this.late;
				this.late = [];
				this.process(next);
			}
		} finally {
			for (const n of this.marked) {
				this.graph.setClean(n);
				n.inEval = false;
				n.onStack = false;
			}
			this.marked = [];
			this.late = [];
			this.inPass = false;
			this.host.endPass();
		}
		return this.circular.sorted();
	}

	/** Mark a formula dirty during a pass (a spill it reads changed shape after it was computed). */
	redirty(node: FormulaNode): void {
		if (node.evalCount >= MAX_EVALS) {
			return;
		}
		const list: FormulaNode[] = [];
		this.markClosure([node], [], list);
		for (const n of list) {
			this.late.push(n);
		}
	}

	private markDirty(n: FormulaNode): boolean {
		if (n.frozen || n.dirty) {
			return false;
		}
		if (n.recalcId !== this.recalcId) {
			n.recalcId = this.recalcId;
			n.evalCount = 0;
		}
		this.graph.setDirty(n);
		this.marked.push(n);
		return true;
	}

	private markClosure(
		roots: Iterable<FormulaNode>,
		changed: readonly CellRef[],
		out: FormulaNode[],
		propagate = true,
	): void {
		const stack: FormulaNode[] = [];
		const mark = (n: FormulaNode): void => {
			if (this.markDirty(n)) {
				out.push(n);
				stack.push(n);
			}
		};
		const round = this.graph.nextRound();
		for (const c of changed) {
			this.graph.forEachDependentOfCell(c.sheet, c.row, c.col, mark, round);
		}
		for (const r of roots) {
			mark(r);
		}
		if (!propagate) {
			return;
		}
		while (stack.length > 0) {
			const n = stack.pop()!;
			this.graph.forEachDependentOfCell(n.sheet, n.row, n.col, mark, round);
			if (n.spill !== null) {
				this.graph.forEachDependentOfRange(n.sheet, n.spill, mark);
			}
			if (n.cse !== null) {
				this.graph.forEachDependentOfRange(n.sheet, n.cse, mark);
			}
		}
	}

	private process(list: readonly FormulaNode[]): void {
		const stamp = ++this.stamp;
		for (const n of list) {
			if (n.dirty && n.tStamp !== stamp) {
				this.tarjan(n, stamp);
			}
		}
	}

	private tarjan(root: FormulaNode, stamp: number): void {
		const frames: TarjanFrame[] = [];
		const stack: FormulaNode[] = [];
		const open = (v: FormulaNode): void => {
			v.tStamp = stamp;
			v.tIndex = v.tLow = this.index++;
			v.onStack = true;
			stack.push(v);
			const succ: FormulaNode[] = [];
			const self = dirtyPrecedents(this.graph, v, succ);
			frames.push({ node: v, succ, i: 0, self });
		};
		open(root);
		while (frames.length > 0) {
			const f = frames[frames.length - 1];
			if (f.i < f.succ.length) {
				const w = f.succ[f.i++];
				if (!w.dirty) {
					continue;
				}
				if (w.tStamp !== stamp) {
					open(w);
				} else if (w.onStack && w.tIndex < f.node.tLow) {
					f.node.tLow = w.tIndex;
				}
				continue;
			}
			frames.pop();
			const v = f.node;
			if (frames.length > 0) {
				const parent = frames[frames.length - 1].node;
				if (v.tLow < parent.tLow) {
					parent.tLow = v.tLow;
				}
			}
			if (v.tLow !== v.tIndex) {
				continue;
			}
			const component: FormulaNode[] = [];
			let w: FormulaNode;
			do {
				w = stack.pop()!;
				w.onStack = false;
				component.push(w);
			} while (w !== v);
			if (component.length === 1 && !f.self) {
				this.evaluate(v, false);
			} else {
				this.cycle(component);
			}
		}
	}

	/**
	 * Evaluate a node, first computing any dirty cells it turns out to read
	 * (explicit stack instead of recursion). With `force` the node is
	 * evaluated even though it is not dirty (iterative calculation).
	 */
	private evaluate(node: FormulaNode, force: boolean): void {
		const stack: FormulaNode[] = [node];
		node.inEval = true;
		let forced: FormulaNode | null = force ? node : null;
		while (stack.length > 0) {
			const top = stack[stack.length - 1];
			if (!top.dirty && top !== forced) {
				top.inEval = false;
				stack.pop();
				continue;
			}
			try {
				this.compute(top, top === forced);
				if (top === forced) {
					forced = null;
				}
				top.inEval = false;
				stack.pop();
			} catch (e) {
				if (!(e instanceof PendingCells)) {
					throw e;
				}
				let cycleAt = -1;
				for (const n of e.nodes) {
					if (n.inEval) {
						const i = stack.indexOf(n);
						if (i >= 0 && (cycleAt < 0 || i < cycleAt)) {
							cycleAt = i;
						}
					}
				}
				if (cycleAt >= 0) {
					for (const member of stack.splice(cycleAt)) {
						member.inEval = false;
						this.circularValue(member);
					}
					continue;
				}
				let pushed = false;
				for (const n of e.nodes) {
					if (n.dirty && !n.inEval) {
						n.inEval = true;
						stack.push(n);
						pushed = true;
					}
				}
				if (!pushed) {
					// Defensive: no way to make progress, treat the reader as circular.
					stack.pop();
					top.inEval = false;
					this.circularValue(top);
				}
			}
		}
	}

	/** Evaluate one formula and store its result (throws PendingCells to suspend). */
	private compute(node: FormulaNode, forced: boolean): void {
		node.evalCount++;
		const cell = this.workbook.sheets[node.sheet]?.cells.get(node.key);
		if (cell === undefined || cell.formula !== node.formula) {
			this.graph.setClean(node);
			return;
		}
		if (!forced && node.evalCount > MAX_EVALS) {
			this.circularValue(node);
			return;
		}
		let result: Scalar | ArrayValue = ERRORS.NAME;
		let arrayResult = false;
		const ast = node.parsed.ast;
		if (ast !== null) {
			try {
				const mode = node.mode === 'legacy' ? 'legacy' : 'dynamic';
				const raw = this.evaluator.evaluate(ast, node.sheet, node.row, node.col, mode);
				if (node.mode === 'legacy') {
					result = legacyCellResult(raw, this.host, node.row, node.col);
				} else {
					result = arrayCellResult(raw, this.host);
					arrayResult = isArrayValue(raw) || isMultiCell(raw);
				}
			} catch (e) {
				if (e instanceof PendingCells) {
					throw e;
				}
				// A failing function implementation must not stop the recalculation.
				result = ERRORS.VALUE;
			}
		}
		this.writer.apply(node, result, arrayResult);
		if (node.spill !== null && this.spillReadsItself(node)) {
			// `=SEQUENCE(3)+A2` in A1 reads its own spill: a circular reference.
			this.circularValue(node);
			return;
		}
		this.graph.setClean(node);
	}

	private spillReadsItself(node: FormulaNode): boolean {
		const spill = node.spill!;
		return readsRange(node, node.sheet, spill);
	}

	/** A circular reference without iterative calculation: the cells show 0. */
	private circularValue(node: FormulaNode): void {
		this.writer.apply(node, 0, false);
		this.graph.setClean(node);
		this.circular.add(node.sheet, node.row, node.col);
	}

	private cycle(component: FormulaNode[]): void {
		if (!this.workbook.calc.iterate) {
			for (const n of component.sort(byPosition)) {
				this.circularValue(n);
			}
			return;
		}
		const members = component.sort(byPosition);
		for (const m of members) {
			this.graph.setClean(m);
		}
		const count = Math.max(1, Math.floor(this.workbook.calc.iterateCount ?? 100));
		const delta = this.workbook.calc.iterateDelta ?? 0.001;
		const valueOf = (n: FormulaNode): unknown =>
			this.workbook.sheets[n.sheet]?.cells.get(n.key)?.value;
		for (let k = 0; k < count; k++) {
			let change = 0;
			for (const m of members) {
				const before = valueOf(m);
				this.evaluate(m, true);
				const after = valueOf(m);
				const diff =
					typeof before === 'number' && typeof after === 'number'
						? Math.abs(after - before)
						: before === after
							? 0
							: Infinity;
				change = Math.max(change, diff);
			}
			if (change < delta) {
				break;
			}
		}
	}
}
