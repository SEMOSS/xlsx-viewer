import { cellKey } from '../address/cell-address';
import { MAX_COL_INDEX, MAX_COLS, MAX_ROW_INDEX } from '../address/constants';
import { broadcastElement } from '../formula/invoke';
import type { ArrayValue, Scalar } from '../formula/types';
import type { RangeAddress } from '../types/address';
import type { Cell } from '../types/cell';
import type { Worksheet } from '../types/sheet';
import type { Workbook } from '../types/workbook';
import { ERRORS } from '../values/errors';
import type { DependencyGraph } from './graph';
import type { WorkbookHost } from './host';
import type { TableTable } from './lookups';
import type { FormulaNode } from './node';
import type { CellRef } from './types';

/** Cells whose value changed during a recalculation (deduplicated). */
export class ChangeSet {
	private readonly seen = new Set<number>();
	private readonly list: CellRef[] = [];

	add(sheet: number, row: number, col: number): void {
		const k = sheet * 2 ** 34 + cellKey(row, col);
		if (!this.seen.has(k)) {
			this.seen.add(k);
			this.list.push({ sheet, row, col });
		}
	}

	get size(): number {
		return this.list.length;
	}

	/** The changes in sheet / row / column order. */
	sorted(): CellRef[] {
		return this.list.sort((a, b) => a.sheet - b.sheet || a.row - b.row || a.col - b.col);
	}
}

/** Notifications from the writer, used to schedule follow-up recalculation. */
export interface WriterHooks {
	/** A cell written or cleared on behalf of an array formula changed value. */
	memberChanged(sheet: number, row: number, col: number): void;
	/** A spill gave cells back (formulas blocked by them may now spill). */
	released(sheet: number, range: RangeAddress, by: FormulaNode): void;
}

function single(node: FormulaNode): RangeAddress {
	return { top: node.row, left: node.col, bottom: node.row, right: node.col };
}

function sameRange(a: RangeAddress, b: RangeAddress): boolean {
	return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}

function isArray(v: Scalar | ArrayValue): v is ArrayValue {
	return typeof v === 'object' && v !== null && (v as { kind?: unknown }).kind === 'array';
}

/**
 * Stores formula results: the anchor's value, CSE array blocks and
 * dynamic-array spills (with #SPILL! detection). Cells are replaced, never
 * mutated; spilled cells keep their style.
 */
export class ResultWriter {
	changes = new ChangeSet();

	constructor(
		private readonly workbook: Workbook,
		private readonly graph: DependencyGraph,
		private readonly host: WorkbookHost,
		private readonly tables: TableTable,
		private readonly hooks: WriterHooks,
	) {}

	/** Store a cell result. `arrayResult`: the formula produced an array (dynamic `A1#` validity). */
	apply(node: FormulaNode, result: Scalar | ArrayValue, arrayResult: boolean): void {
		switch (node.mode) {
			case 'legacy':
				this.writeAnchor(node, isArray(result) ? result.values[0][0] : result, undefined);
				return;
			case 'cse':
				this.applyCse(node, result);
				return;
			case 'dynamic':
				this.applyDynamic(node, result, arrayResult);
				return;
		}
	}

	/** Remove what a formula wrote into other cells (spill and CSE members). */
	clearOutputs(node: FormulaNode): void {
		this.releaseSpill(node, null);
		const block = node.cse;
		if (block !== null) {
			for (let r = block.top; r <= block.bottom; r++) {
				for (let c = block.left; c <= block.right; c++) {
					if (r !== node.row || c !== node.col) {
						this.clearMember(node.sheet, r, c, node.key);
					}
				}
			}
		}
		this.graph.setBlocked(node, null);
	}

	private applyCse(node: FormulaNode, result: Scalar | ArrayValue): void {
		const block = node.cse ?? single(node);
		let anchor: Scalar = null;
		for (let r = block.top; r <= block.bottom; r++) {
			for (let c = block.left; c <= block.right; c++) {
				const v = isArray(result)
					? broadcastElement(result, r - block.top, c - block.left)
					: result;
				if (r === node.row && c === node.col) {
					anchor = v;
				} else {
					this.writeMember(node.sheet, r, c, v, node.key);
				}
			}
		}
		this.writeAnchor(node, anchor, undefined);
	}

	private applyDynamic(node: FormulaNode, result: Scalar | ArrayValue, arrayResult: boolean): void {
		node.arrayResult = arrayResult;
		if (!isArray(result)) {
			this.releaseSpill(node, null);
			this.graph.setBlocked(node, null);
			this.writeAnchor(node, result, single(node));
			return;
		}
		const target: RangeAddress = {
			top: node.row,
			left: node.col,
			bottom: node.row + result.rows - 1,
			right: node.col + result.cols - 1,
		};
		if (
			target.bottom > MAX_ROW_INDEX ||
			target.right > MAX_COL_INDEX ||
			this.isBlocked(node, target)
		) {
			this.releaseSpill(node, null);
			this.graph.setBlocked(node, target);
			this.writeAnchor(node, ERRORS.SPILL, single(node));
			return;
		}
		this.graph.setBlocked(node, null);
		for (let r = 0; r < result.rows; r++) {
			const line = result.values[r];
			for (let c = 0; c < result.cols; c++) {
				if (r !== 0 || c !== 0) {
					this.writeMember(node.sheet, node.row + r, node.col + c, line[c], node.key);
				}
			}
		}
		this.releaseSpill(node, target);
		this.graph.setSpill(node, target);
		this.writeAnchor(node, result.values[0][0], target);
	}

	/** A dynamic array result cannot spill into `target`. */
	private isBlocked(node: FormulaNode, target: RangeAddress): boolean {
		const ws = this.workbook.sheets[node.sheet];
		if (ws === undefined) {
			return true;
		}
		if (this.tables.overlaps(node.sheet, target.top, target.left, target.bottom, target.right)) {
			return true;
		}
		for (const m of ws.merges) {
			if (
				m.left <= target.right &&
				target.left <= m.right &&
				m.top <= target.bottom &&
				target.top <= m.bottom
			) {
				return true;
			}
		}
		const blocks = (key: number): boolean => {
			if (key === node.key) {
				return false;
			}
			const cell = ws.cells.get(key);
			if (cell === undefined || cell.value === null || cell.arrayAnchor === node.key) {
				return cell?.formula !== undefined;
			}
			if (cell.formula !== undefined || cell.arrayAnchor === undefined) {
				return true;
			}
			// Part of another array: blocks only while that formula still owns the cell.
			return this.owns(node.sheet, cell.arrayAnchor, key);
		};
		const area = (target.bottom - target.top + 1) * (target.right - target.left + 1);
		const idx = this.host.index(node.sheet);
		if (area <= 4 * idx.count + 64) {
			for (let r = target.top; r <= target.bottom; r++) {
				const base = cellKey(r, 0);
				for (let c = target.left; c <= target.right; c++) {
					if (blocks(base + c)) {
						return true;
					}
				}
			}
			return false;
		}
		return idx.keysIn(target).some(blocks);
	}

	/** The array formula anchored at `anchorKey` currently covers the cell `key`. */
	private owns(sheet: number, anchorKey: number, key: number): boolean {
		const anchor = this.graph.node(sheet, anchorKey);
		const area = anchor?.spill ?? anchor?.cse ?? null;
		if (area === null) {
			return false;
		}
		const row = Math.floor(key / MAX_COLS);
		const col = key - row * MAX_COLS;
		return row >= area.top && row <= area.bottom && col >= area.left && col <= area.right;
	}

	/** Clear the cells of the current spill outside `keep` (all of them when `keep` is null). */
	private releaseSpill(node: FormulaNode, keep: RangeAddress | null): void {
		const old = node.spill;
		if (old === null) {
			return;
		}
		if (keep !== null && sameRange(old, keep)) {
			return;
		}
		let freed = false;
		for (let r = old.top; r <= old.bottom; r++) {
			for (let c = old.left; c <= old.right; c++) {
				const inside =
					keep !== null && r <= keep.bottom && c <= keep.right && r >= keep.top && c >= keep.left;
				if (!inside && (r !== node.row || c !== node.col)) {
					freed = this.clearMember(node.sheet, r, c, node.key) || freed;
				}
			}
		}
		if (keep === null) {
			this.graph.setSpill(node, null);
		}
		if (freed) {
			this.hooks.released(node.sheet, old, node);
		}
	}

	private sheetOf(index: number): Worksheet | undefined {
		return this.workbook.sheets[index];
	}

	/** Write one element of an array result into a member cell. */
	private writeMember(
		sheet: number,
		row: number,
		col: number,
		value: Scalar,
		anchorKey: number,
	): void {
		const ws = this.sheetOf(sheet);
		if (ws === undefined) {
			return;
		}
		const key = cellKey(row, col);
		const cell = ws.cells.get(key);
		if (
			cell !== undefined &&
			cell.arrayAnchor === anchorKey &&
			cell.formula === undefined &&
			cell.value === value
		) {
			return;
		}
		const next: Cell = { value, arrayAnchor: anchorKey };
		if (cell?.style !== undefined) {
			next.style = cell.style;
		}
		ws.cells.set(key, next);
		ws.dirty = true;
		if (cell === undefined) {
			this.host.noteCell(sheet, row, col);
		} else {
			this.host.noteValue(sheet, row, col, value);
		}
		if (cell === undefined || cell.value !== value) {
			this.changes.add(sheet, row, col);
			this.hooks.memberChanged(sheet, row, col);
		}
	}

	/** Clear a member cell owned by `anchorKey`; returns true when a cell was released. */
	private clearMember(sheet: number, row: number, col: number, anchorKey: number): boolean {
		const ws = this.sheetOf(sheet);
		const key = cellKey(row, col);
		const cell = ws?.cells.get(key);
		if (
			ws === undefined ||
			cell === undefined ||
			cell.arrayAnchor !== anchorKey ||
			cell.formula !== undefined
		) {
			return false;
		}
		if (cell.style !== undefined) {
			ws.cells.set(key, { value: null, style: cell.style });
			this.host.noteValue(sheet, row, col, null);
		} else {
			ws.cells.delete(key);
			this.host.noteCell(sheet, row, col);
		}
		ws.dirty = true;
		if (cell.value !== null) {
			this.changes.add(sheet, row, col);
			this.hooks.memberChanged(sheet, row, col);
		}
		return true;
	}

	/** Store the anchor's own value (and, for dynamic arrays, its spill range). */
	writeAnchor(node: FormulaNode, value: Scalar, arrayRef: RangeAddress | undefined): void {
		const ws = this.sheetOf(node.sheet);
		const cell = ws?.cells.get(node.key);
		if (ws === undefined || cell === undefined || cell.formula !== node.formula) {
			return;
		}
		let array = cell.array;
		if (
			arrayRef !== undefined &&
			(array === undefined || !array.dynamic || !sameRange(array.ref, arrayRef))
		) {
			array = { ref: arrayRef, dynamic: true };
		}
		if (cell.value === value && array === cell.array && cell.stale !== true) {
			return;
		}
		const next: Cell = { ...cell, value };
		if (array !== undefined) {
			next.array = array;
		}
		if (next.stale !== undefined) {
			delete next.stale;
		}
		ws.cells.set(node.key, next);
		ws.dirty = true;
		this.host.noteValue(node.sheet, node.row, node.col, value);
		if (cell.value !== value) {
			this.changes.add(node.sheet, node.row, node.col);
		}
	}
}
