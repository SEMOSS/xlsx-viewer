import { cellKey, keyCol, keyRow } from '../address/cell-address';
import type { RangeAddress } from '../types/address';
import { FormulaColumn } from './formula-column';
import type { FormulaNode, RangeDep } from './node';
import { RangeIndex } from './range-index';

type Dependents = FormulaNode | FormulaNode[];

/** Graph state of one sheet. */
export class SheetGraph {
	/** Formula nodes by cell key. */
	readonly nodes = new Map<number, FormulaNode>();
	/** Rows of formula cells per column, sorted. */
	readonly formulaCols = new Map<number, FormulaColumn>();
	/** Formulas reading a single cell, by cell key. */
	readonly cellDeps = new Map<number, Dependents>();
	/** Distinct ranges read by formulas, spatially indexed. */
	readonly ranges = new RangeIndex<RangeDep>();
	readonly rangeDeps = new Map<string, RangeDep>();
	/** Active multi-cell spill ranges, by anchor. */
	readonly spills = new RangeIndex<FormulaNode>();
	/** Ranges blocked dynamic-array formulas (#SPILL!) want to fill. */
	readonly blocked = new RangeIndex<FormulaNode>();
	dirtyCount = 0;
	/** Columns are appended unsorted while building; sorted before use. */
	unsorted = false;
}

/**
 * Dependency graph: formula nodes per sheet, their registered precedents
 * and the reverse index answering "which formulas read this cell".
 */
export class DependencyGraph {
	private sheets: SheetGraph[] = [];
	readonly volatile = new Set<FormulaNode>();
	dirtyTotal = 0;
	private round = 0;

	reset(): void {
		this.sheets = [];
		this.volatile.clear();
		this.dirtyTotal = 0;
	}

	sheet(index: number): SheetGraph {
		let g = this.sheets[index];
		if (g === undefined) {
			g = new SheetGraph();
			this.sheets[index] = g;
		}
		return g;
	}

	peek(index: number): SheetGraph | undefined {
		return this.sheets[index];
	}

	node(sheet: number, key: number): FormulaNode | undefined {
		return this.sheets[sheet]?.nodes.get(key);
	}

	*allNodes(): Generator<FormulaNode> {
		for (const g of this.sheets) {
			if (g !== undefined) {
				yield* g.nodes.values();
			}
		}
	}

	get nodeCount(): number {
		let n = 0;
		for (const g of this.sheets) {
			n += g?.nodes.size ?? 0;
		}
		return n;
	}

	/** Add a node with its precedents: (sheet, key) pairs and (sheet, range) areas. */
	add(
		node: FormulaNode,
		cells: readonly number[],
		ranges: ReadonlyArray<{ sheet: number; range: RangeAddress }>,
		bulk = false,
	): void {
		const home = this.sheet(node.sheet);
		home.nodes.set(node.key, node);
		let column = home.formulaCols.get(node.col);
		if (column === undefined) {
			column = new FormulaColumn();
			home.formulaCols.set(node.col, column);
		}
		if (bulk) {
			column.push(node.row);
			home.unsorted = true;
		} else {
			column.insert(node.row);
		}
		for (let i = 0; i < cells.length; i += 2) {
			const g = this.sheet(cells[i]);
			const key = cells[i + 1];
			const cur = g.cellDeps.get(key);
			if (cur === undefined) {
				g.cellDeps.set(key, node);
			} else if (Array.isArray(cur)) {
				cur.push(node);
			} else {
				g.cellDeps.set(key, [cur, node]);
			}
		}
		const deps: RangeDep[] = [];
		for (const { sheet, range } of ranges) {
			const g = this.sheet(sheet);
			const id = `${range.top}:${range.left}:${range.bottom}:${range.right}`;
			let dep = g.rangeDeps.get(id);
			if (dep === undefined) {
				dep = { sheet, range, id, dependents: new Set(), entry: null };
				dep.entry = g.ranges.insert(range, dep);
				g.rangeDeps.set(id, dep);
			}
			dep.dependents.add(node);
			deps.push(dep);
		}
		node.deps = { cells: cells.slice(), ranges: deps };
		if (node.volatile) {
			this.volatile.add(node);
		}
	}

	/** Sort the formula column index after bulk additions. */
	finishBulk(): void {
		for (const g of this.sheets) {
			if (g?.unsorted) {
				for (const column of g.formulaCols.values()) {
					column.sort();
				}
				g.unsorted = false;
			}
		}
	}

	remove(node: FormulaNode): void {
		const home = this.sheets[node.sheet];
		if (home === undefined || home.nodes.get(node.key) !== node) {
			return;
		}
		this.setClean(node);
		home.nodes.delete(node.key);
		const column = home.formulaCols.get(node.col);
		if (column !== undefined) {
			column.remove(node.row);
			if (column.rows.length === 0) {
				home.formulaCols.delete(node.col);
			}
		}
		const cells = node.deps.cells;
		for (let i = 0; i < cells.length; i += 2) {
			const g = this.sheets[cells[i]];
			const key = cells[i + 1];
			const cur = g?.cellDeps.get(key);
			if (cur === undefined) {
				continue;
			}
			if (cur === node) {
				g!.cellDeps.delete(key);
			} else if (Array.isArray(cur)) {
				const j = cur.indexOf(node);
				if (j >= 0) {
					cur.splice(j, 1);
				}
				if (cur.length === 1) {
					g!.cellDeps.set(key, cur[0]);
				}
			}
		}
		for (const dep of node.deps.ranges) {
			dep.dependents.delete(node);
			if (dep.dependents.size === 0) {
				const g = this.sheets[dep.sheet];
				if (g !== undefined) {
					if (dep.entry !== null) {
						g.ranges.remove(dep.entry);
					}
					g.rangeDeps.delete(dep.id);
				}
			}
		}
		node.deps = { cells: [], ranges: [] };
		this.volatile.delete(node);
		this.setSpill(node, null);
		this.setBlocked(node, null);
		if (node.cseEntry !== null) {
			home.spills.remove(node.cseEntry);
			node.cseEntry = null;
		}
	}

	setDirty(node: FormulaNode): void {
		if (node.dirty) {
			return;
		}
		node.dirty = true;
		const g = this.sheet(node.sheet);
		g.dirtyCount++;
		g.formulaCols.get(node.col)?.mark(node.row, 1);
		this.dirtyTotal++;
	}

	setClean(node: FormulaNode): void {
		if (!node.dirty) {
			return;
		}
		node.dirty = false;
		const g = this.sheet(node.sheet);
		g.dirtyCount--;
		g.formulaCols.get(node.col)?.mark(node.row, -1);
		this.dirtyTotal--;
	}

	/** Update the active spill range of an anchor. */
	setSpill(node: FormulaNode, range: RangeAddress | null): void {
		const g = this.sheet(node.sheet);
		if (node.spillEntry !== null) {
			g.spills.remove(node.spillEntry);
			node.spillEntry = null;
		}
		node.spill = range;
		if (range !== null) {
			node.spillEntry = g.spills.insert(range, node);
		}
	}

	/** Register the block of a CSE array formula, so readers of its cells find the anchor. */
	setCse(node: FormulaNode, range: RangeAddress | null): void {
		const g = this.sheet(node.sheet);
		if (node.cseEntry !== null) {
			g.spills.remove(node.cseEntry);
			node.cseEntry = null;
		}
		node.cse = range;
		if (range !== null) {
			node.cseEntry = g.spills.insert(range, node);
		}
	}

	/** Update the range a blocked dynamic-array formula wants to fill. */
	setBlocked(node: FormulaNode, range: RangeAddress | null): void {
		const g = this.sheet(node.sheet);
		if (node.blockedEntry !== null) {
			g.blocked.remove(node.blockedEntry);
			node.blockedEntry = null;
		}
		node.blocked = range;
		if (range !== null) {
			node.blockedEntry = g.blocked.insert(range, node);
		}
	}

	/** A fresh round for {@link forEachDependentOfCell} (ranges are then visited once per round). */
	nextRound(): number {
		return ++this.round;
	}

	/**
	 * Visit the formulas that read the cell (directly or through a range).
	 * Within one `round` a shared range is visited only the first time.
	 */
	forEachDependentOfCell(
		sheet: number,
		row: number,
		col: number,
		visit: (n: FormulaNode) => void,
		round = 0,
	): void {
		const g = this.sheets[sheet];
		if (g === undefined) {
			return;
		}
		const d = g.cellDeps.get(cellKey(row, col));
		if (d !== undefined) {
			if (Array.isArray(d)) {
				d.forEach(visit);
			} else {
				visit(d);
			}
		}
		g.ranges.queryPoint(row, col, (dep) => dep.dependents.forEach(visit), round);
	}

	/** Visit the formulas that read any cell of a range. */
	forEachDependentOfRange(
		sheet: number,
		range: RangeAddress,
		visit: (n: FormulaNode) => void,
	): void {
		const g = this.sheets[sheet];
		if (g === undefined) {
			return;
		}
		g.ranges.queryRange(range, (dep) => dep.dependents.forEach(visit));
		const area = (range.bottom - range.top + 1) * (range.right - range.left + 1);
		const each = (d: Dependents) => (Array.isArray(d) ? d.forEach(visit) : visit(d));
		if (area <= g.cellDeps.size) {
			for (let r = range.top; r <= range.bottom; r++) {
				for (let c = range.left; c <= range.right; c++) {
					const d = g.cellDeps.get(cellKey(r, c));
					if (d !== undefined) {
						each(d);
					}
				}
			}
			return;
		}
		for (const [key, d] of g.cellDeps) {
			const r = keyRow(key);
			const c = keyCol(key);
			if (r >= range.top && r <= range.bottom && c >= range.left && c <= range.right) {
				each(d);
			}
		}
	}

	/** Visit the dirty formula cells inside a range. */
	forEachDirtyIn(sheet: number, range: RangeAddress, visit: (n: FormulaNode) => void): void {
		const g = this.sheets[sheet];
		if (g === undefined || g.dirtyCount === 0) {
			return;
		}
		const scan = (col: number, column: FormulaColumn) => {
			if (column.dirty === 0) {
				return;
			}
			const nodeAt = (row: number) => g.nodes.get(cellKey(row, col));
			column.forEachDirty(
				range.top,
				range.bottom,
				(row) => nodeAt(row)?.dirty === true,
				(row) => visit(nodeAt(row)!),
			);
		};
		if (range.right - range.left + 1 > g.formulaCols.size) {
			for (const [col, column] of g.formulaCols) {
				if (col >= range.left && col <= range.right) {
					scan(col, column);
				}
			}
			return;
		}
		for (let col = range.left; col <= range.right; col++) {
			const column = g.formulaCols.get(col);
			if (column !== undefined) {
				scan(col, column);
			}
		}
	}

	/** Visit dirty anchors whose active spill intersects a range. */
	forEachDirtySpillIn(sheet: number, range: RangeAddress, visit: (n: FormulaNode) => void): void {
		const g = this.sheets[sheet];
		if (g === undefined || g.spills.size === 0 || g.dirtyCount === 0) {
			return;
		}
		g.spills.queryRange(range, (n) => {
			if (n.dirty) {
				visit(n);
			}
		});
	}
}
