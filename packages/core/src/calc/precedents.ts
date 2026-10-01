import { keyCol, keyRow } from '../address/cell-address';
import type { RangeAddress } from '../types/address';
import type { DependencyGraph } from './graph';
import type { FormulaNode } from './node';

/** The node reads a cell of `range` on `sheet` (other than its own cell). */
export function readsRange(node: FormulaNode, sheet: number, range: RangeAddress): boolean {
	const cells = node.deps.cells;
	for (let i = 0; i < cells.length; i += 2) {
		if (cells[i] !== sheet || cells[i + 1] === node.key) {
			continue;
		}
		const row = keyRow(cells[i + 1]);
		const col = keyCol(cells[i + 1]);
		if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) {
			return true;
		}
	}
	for (const dep of node.deps.ranges) {
		const r = dep.range;
		if (
			dep.sheet === sheet &&
			r.left <= range.right &&
			range.left <= r.right &&
			r.top <= range.bottom &&
			range.top <= r.bottom
		) {
			return true;
		}
	}
	return false;
}

/**
 * Collect the dirty static precedents of a node into `out`: formulas in the
 * cells and ranges it reads, and array formulas whose output covers them.
 * Returns true when the node reads itself (a one-cell circular reference).
 */
export function dirtyPrecedents(
	graph: DependencyGraph,
	node: FormulaNode,
	out: FormulaNode[],
): boolean {
	if (graph.dirtyTotal === 0) {
		return false;
	}
	let self = false;
	const add = (n: FormulaNode) => {
		if (n === node) {
			self = true;
		} else {
			out.push(n);
		}
	};
	const cells = node.deps.cells;
	for (let i = 0; i < cells.length; i += 2) {
		const g = graph.peek(cells[i]);
		if (g === undefined || g.dirtyCount === 0) {
			continue;
		}
		const key = cells[i + 1];
		const n = g.nodes.get(key);
		if (n !== undefined) {
			if (n.dirty) {
				add(n);
			}
		} else if (g.spills.size > 0) {
			g.spills.queryPoint(keyRow(key), keyCol(key), (a) => {
				if (a.dirty) {
					add(a);
				}
			});
		}
	}
	for (const dep of node.deps.ranges) {
		graph.forEachDirtyIn(dep.sheet, dep.range, add);
		graph.forEachDirtySpillIn(dep.sheet, dep.range, add);
	}
	return self;
}
