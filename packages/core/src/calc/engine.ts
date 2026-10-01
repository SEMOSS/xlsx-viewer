import { cellKey, keyCol, keyRow } from '../address/cell-address';
import { Evaluator } from '../formula/evaluator/evaluator';
import { legacyCellResult } from '../formula/evaluator/results';
import { getDefaultRegistry, type FunctionRegistry } from '../formula/registry';
import type { FormulaValue } from '../formula/types';
import type { RangeAddress } from '../types/address';
import type { Workbook } from '../types/workbook';
import { ERRORS, isCellError } from '../values/errors';
import { BuildIssues, createNode, modeOf, reportIssues } from './build';
import { FormulaCache } from './formula-cache';
import { DependencyGraph } from './graph';
import { PendingCells, WorkbookHost } from './host';
import { NameTable, SheetList, TableTable } from './lookups';
import type { FormulaNode } from './node';
import { Scheduler } from './scheduler';
import type {
	CalcEngineOptions,
	CellRef,
	EvaluateOptions,
	PrecedentArea,
	RecalcResult,
} from './types';
import { ChangeSet, ResultWriter, type WriterHooks } from './writer';

function sameRange(a: RangeAddress | null, b: RangeAddress | undefined): boolean {
	if (a === null || b === undefined) {
		return a === null && (b === undefined || (b.top === b.bottom && b.left === b.right));
	}
	return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}

/**
 * Calculation engine: keeps a dependency graph of every formula in the
 * workbook and recalculates in dependency order.
 *
 * - `rebuild()` builds the graph and trusts the values saved in the file
 *   (only formulas without a cached value, or every formula when the file
 *   asks for a full calculation on load, are computed).
 * - `recalculateAll()` recomputes every formula; `onCellsChanged()` the
 *   formulas affected by edited cells; `recalculateVolatile()` the volatile
 *   ones (NOW, RAND, OFFSET, INDIRECT, ...) and their dependents.
 * - Dynamic-array formulas spill into neighbouring cells (or show #SPILL!),
 *   legacy CSE arrays fill their block, circular references show 0 unless
 *   iterative calculation is enabled.
 */
export class CalcEngine {
	private readonly registry: FunctionRegistry = getDefaultRegistry();
	private readonly cache = new FormulaCache();
	private readonly graph = new DependencyGraph();
	private readonly names = new NameTable();
	private readonly tables = new TableTable();
	private readonly sheetList = new SheetList();
	private readonly host: WorkbookHost;
	private readonly evaluator: Evaluator;
	private readonly writer: ResultWriter;
	private readonly scheduler: Scheduler;
	private built = false;
	private firstBuild = true;
	/** Outside a pass (while edits are applied) writer notifications are collected here. */
	private collected: { cells: CellRef[]; roots: FormulaNode[] } | null = null;

	constructor(
		readonly workbook: Workbook,
		readonly options: CalcEngineOptions = {},
	) {
		this.host = new WorkbookHost(
			workbook,
			this.graph,
			this.names,
			this.tables,
			options,
			this.registry,
		);
		this.evaluator = new Evaluator(this.host);
		const hooks: WriterHooks = {
			memberChanged: (sheet, row, col) => this.memberChanged(sheet, row, col),
			released: (sheet, range, by) => this.released(sheet, range, by),
		};
		this.writer = new ResultWriter(workbook, this.graph, this.host, this.tables, hooks);
		this.scheduler = new Scheduler(workbook, this.graph, this.host, this.writer, this.evaluator);
	}

	/** Rebuild the dependency graph from all formulas (after load or structural edits). */
	rebuild(): void {
		const wb = this.workbook;
		this.graph.reset();
		this.names.build(wb, this.cache);
		this.tables.build(wb);
		this.sheetList.build(wb);
		this.host.reset();
		this.evaluator.reset();
		const fullCalc = this.firstBuild && wb.calc.fullCalcOnLoad === true;
		const issues = new BuildIssues();
		const roots: FormulaNode[] = [];
		const ctx = this.buildContext();
		wb.sheets.forEach((ws, sheet) => {
			for (const [key, cell] of ws.cells) {
				if (cell.formula === undefined) {
					continue;
				}
				const node = createNode(ctx, sheet, key, cell, true, issues, true);
				if (node.frozen) {
					continue;
				}
				const spillError = isCellError(cell.value) && cell.value.code === '#SPILL!';
				if (fullCalc || cell.value === null || cell.stale === true || spillError) {
					roots.push(node);
				}
			}
		});
		this.graph.finishBulk();
		reportIssues(wb, issues);
		this.built = true;
		this.firstBuild = false;
		if (roots.length > 0) {
			this.writer.changes = new ChangeSet();
			this.scheduler.run(roots, []);
		}
	}

	/** Recalculate every formula (F9). */
	recalculateAll(): RecalcResult {
		this.sync();
		const roots: FormulaNode[] = [];
		for (const n of this.graph.allNodes()) {
			if (!n.frozen) {
				roots.push(n);
			}
		}
		return this.run(roots, [], true, true);
	}

	/** Recalculate the dependents of cells whose value or formula changed. */
	onCellsChanged(changes: readonly CellRef[]): RecalcResult {
		if (!this.built || this.structureChanged()) {
			this.rebuild();
			return this.recalculateAll();
		}
		this.writer.changes = new ChangeSet();
		const cells: CellRef[] = [];
		const roots: FormulaNode[] = [];
		this.collected = { cells, roots };
		const seen = new Set<number>();
		try {
			for (const c of changes) {
				const ws = this.workbook.sheets[c.sheet];
				const key = cellKey(c.row, c.col);
				const packed = c.sheet * 2 ** 34 + key;
				if (ws === undefined || seen.has(packed)) {
					continue;
				}
				seen.add(packed);
				this.host.noteCell(c.sheet, c.row, c.col);
				this.applyCellChange(c.sheet, key, roots);
				cells.push({ sheet: c.sheet, row: c.row, col: c.col });
				const g = this.graph.peek(c.sheet);
				const touch = (a: FormulaNode) => {
					if (a.key !== key) {
						roots.push(a);
					}
				};
				g?.spills.queryPoint(c.row, c.col, touch);
				g?.blocked.queryPoint(c.row, c.col, touch);
			}
		} finally {
			this.collected = null;
		}
		for (const v of this.graph.volatile) {
			roots.push(v);
		}
		return this.run(roots, cells, false);
	}

	/** Recalculate volatile formulas (NOW, RAND, OFFSET, INDIRECT, ...) and their dependents. */
	recalculateVolatile(): RecalcResult {
		this.sync();
		return this.run([...this.graph.volatile], []);
	}

	/** Evaluate a formula (a leading `=` is ignored) as if it were in the given cell, without storing it. */
	evaluateFormula(
		formula: string,
		sheet: number,
		row: number,
		col: number,
		options: EvaluateOptions = {},
	): FormulaValue {
		if (!this.built) {
			this.rebuild();
		}
		const text = formula.startsWith('=') ? formula.slice(1) : formula;
		const parsed = this.cache.get(text);
		if (parsed.ast === null) {
			return ERRORS.NAME;
		}
		const dynamic = options.dynamicArrays !== false;
		const checking = this.host.checking;
		this.host.checking = false;
		try {
			const raw = this.evaluator.evaluate(
				parsed.ast,
				sheet,
				row,
				col,
				dynamic ? 'dynamic' : 'legacy',
			);
			return dynamic ? raw : legacyCellResult(raw, this.host, row, col);
		} catch (e) {
			return e instanceof PendingCells ? ERRORS.CALC : ERRORS.VALUE;
		} finally {
			this.host.checking = checking;
		}
	}

	/** Direct precedents of a formula cell: the cells and ranges it reads (names expanded). */
	getPrecedents(ref: CellRef): PrecedentArea[] {
		this.sync();
		const node = this.graph.node(ref.sheet, cellKey(ref.row, ref.col));
		if (node === undefined) {
			return [];
		}
		const out: PrecedentArea[] = [];
		const cells = node.deps.cells;
		for (let i = 0; i < cells.length; i += 2) {
			const row = keyRow(cells[i + 1]);
			const col = keyCol(cells[i + 1]);
			out.push({ sheet: cells[i], range: { top: row, left: col, bottom: row, right: col } });
		}
		for (const dep of node.deps.ranges) {
			out.push({ sheet: dep.sheet, range: { ...dep.range } });
		}
		return out;
	}

	/** Formula cells that read the given cell directly (through a reference, range or name). */
	getDependents(ref: CellRef): CellRef[] {
		this.sync();
		const found = new Set<FormulaNode>();
		this.graph.forEachDependentOfCell(ref.sheet, ref.row, ref.col, (n) => found.add(n));
		return [...found]
			.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col }))
			.sort((a, b) => a.sheet - b.sheet || a.row - b.row || a.col - b.col);
	}

	/** The formula in the cell is a dynamic-array formula whose last result was an array. */
	isDynamicResult(ref: CellRef): boolean {
		const node = this.graph.node(ref.sheet, cellKey(ref.row, ref.col));
		return node !== undefined && node.mode === 'dynamic' && node.arrayResult;
	}

	/** Current spill range of a dynamic-array anchor (null when it does not spill). */
	getSpillRange(ref: CellRef): RangeAddress | null {
		const node = this.graph.node(ref.sheet, cellKey(ref.row, ref.col));
		return node?.spill ? { ...node.spill } : null;
	}

	/** Range a #SPILL! formula would fill if it were not blocked. */
	getBlockedSpillRange(ref: CellRef): RangeAddress | null {
		const node = this.graph.node(ref.sheet, cellKey(ref.row, ref.col));
		return node?.blocked ? { ...node.blocked } : null;
	}

	private buildContext() {
		return {
			workbook: this.workbook,
			graph: this.graph,
			cache: this.cache,
			registry: this.registry,
			names: this.names,
			tables: this.tables,
			sheetIndex: (name: string) => this.host.sheetIndex(name),
		};
	}

	private structureChanged(): boolean {
		const wb = this.workbook;
		return this.sheetList.changed(wb) || this.names.changed(wb) || this.tables.changed(wb);
	}

	/** Rebuild when the graph is missing or sheets / names / tables changed without a rebuild. */
	private sync(): void {
		if (!this.built || this.structureChanged()) {
			this.rebuild();
		}
	}

	private run(
		roots: FormulaNode[],
		cells: CellRef[],
		fresh = true,
		complete = false,
	): RecalcResult {
		if (fresh) {
			this.writer.changes = new ChangeSet();
		}
		const circular = this.scheduler.run(roots, cells, complete);
		return { changed: this.writer.changes.sorted(), circular };
	}

	/** Bring the graph in line with an edited cell. */
	private applyCellChange(sheet: number, key: number, roots: FormulaNode[]): void {
		const cell = this.workbook.sheets[sheet].cells.get(key);
		const node = this.graph.node(sheet, key);
		if (cell?.formula !== undefined) {
			const mode = modeOf(cell);
			if (
				node !== undefined &&
				node.formula === cell.formula &&
				node.mode === mode &&
				(mode !== 'cse' || sameRange(node.cse, cell.array?.ref))
			) {
				// A stale formula (unsupported function) keeps its value; its readers still recalculate.
				if (!(node.frozen && cell.stale === true)) {
					node.frozen = false;
					roots.push(node);
				}
				return;
			}
			if (node !== undefined) {
				this.dropNode(node);
			}
			// Cells still flagged stale (moved by a sort, say) keep their value when unsupported.
			const created = createNode(this.buildContext(), sheet, key, cell, cell.stale === true, null);
			if (!created.frozen) {
				roots.push(created);
			}
			return;
		}
		if (node !== undefined) {
			this.dropNode(node);
		}
	}

	private dropNode(node: FormulaNode): void {
		this.writer.clearOutputs(node);
		this.graph.remove(node);
	}

	private memberChanged(sheet: number, row: number, col: number): void {
		if (this.collected !== null) {
			this.collected.cells.push({ sheet, row, col });
			return;
		}
		if (this.scheduler.inPass) {
			this.graph.forEachDependentOfCell(sheet, row, col, (d) => {
				if (!d.dirty && !d.frozen) {
					this.scheduler.redirty(d);
				}
			});
		}
	}

	private released(sheet: number, range: RangeAddress, by: FormulaNode): void {
		this.graph.peek(sheet)?.blocked.queryRange(range, (n) => {
			if (n === by) {
				return;
			}
			if (this.collected !== null) {
				this.collected.roots.push(n);
			} else if (this.scheduler.inPass && !n.dirty) {
				this.scheduler.redirty(n);
			}
		});
	}
}
