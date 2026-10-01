import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { toBoolean, toNumber } from '../../coerce';
import { toArrayValue } from '../../invoke';
import type { ArrayValue, FormulaValue, FunctionArg, FunctionContext, Scalar } from '../../types';
import { arrayArg, isOmitted, isOmittedArray, scalarArg } from '../lookup/args';
import { vectorValues } from './shared';
import { compareForSort } from './sort-order';
import { valueKey } from './unique';

/** One group at a field level; `children` group the next field. */
export interface GroupNode {
	/** Key as first spelled in the data (keys compare case-insensitively). */
	key: Scalar;
	/** Data rows of the group. */
	rows: number[];
	children: GroupNode[];
}

/** Group data rows by field levels `level` .. `levels - 1`, in order of first appearance. */
export function buildGroups(
	rows: readonly number[],
	keyOf: (row: number, level: number) => Scalar,
	levels: number,
	level = 0,
): GroupNode[] {
	const byKey = new Map<string, GroupNode>();
	const out: GroupNode[] = [];
	for (const r of rows) {
		const key = keyOf(r, level);
		const id = valueKey(key);
		let node = byKey.get(id);
		if (!node) {
			node = { key, rows: [], children: [] };
			byKey.set(id, node);
			out.push(node);
		}
		node.rows.push(r);
	}
	if (level + 1 < levels) {
		for (const node of out) {
			node.children = buildGroups(node.rows, keyOf, levels, level + 1);
		}
	}
	return out;
}

/** Sibling comparator for each field level. */
export type GroupOrder = (level: number) => (a: GroupNode, b: GroupNode) => number;

/** Sort every level of a group tree (stable). */
export function sortGroups(nodes: GroupNode[], order: GroupOrder, level = 0): void {
	nodes.sort(order(level));
	for (const node of nodes) {
		sortGroups(node.children, order, level + 1);
	}
}

/**
 * Sibling order for a sort spec: a number naming the level's field sorts by
 * key, one naming a value column sorts by that column's aggregate (ties by
 * key), otherwise keys ascend. Negative numbers sort descending; Excel's
 * sort order applies (blanks last).
 */
export function groupOrder(
	spec: readonly number[],
	fields: number,
	aggregateOf: (node: GroupNode, valueColumn: number) => Scalar,
): GroupOrder {
	return (level) => {
		const byField = spec.find((k) => Math.abs(k) === level + 1);
		if (byField !== undefined) {
			return (a, b) => compareForSort(a.key, b.key, byField < 0);
		}
		const byValue = spec.find((k) => Math.abs(k) > fields);
		if (byValue === undefined) {
			return (a, b) => compareForSort(a.key, b.key, false);
		}
		const column = Math.abs(byValue) - fields - 1;
		const cache = new Map<GroupNode, Scalar>();
		const measure = (n: GroupNode): Scalar => {
			if (!cache.has(n)) {
				cache.set(n, aggregateOf(n, column));
			}
			return cache.get(n) ?? null;
		};
		return (a, b) =>
			compareForSort(measure(a), measure(b), byValue < 0) || compareForSort(a.key, b.key, false);
	};
}

/** An output line (row, or column for PIVOTBY): a label per field level and the leaf groups it covers. */
export interface Slot {
	labels: Scalar[];
	leaves: number[];
}

export interface Layout {
	slots: Slot[];
	/** Data rows of each leaf group, in output order. */
	leafRows: number[][];
}

/**
 * Lay a sorted group tree out as output lines: one line per leaf group with
 * all its keys; with |depth| >= 2 a subtotal line for each group of the
 * first |depth| - 1 levels (keys up to its level, blanks after), placed
 * after its group (before it when depth is negative); with depth != 0 a
 * "Total" line at the end (start when negative).
 */
export function layoutGroups(tree: GroupNode[], levels: number, depth: number): Layout {
	const slots: Slot[] = [];
	const leafRows: number[][] = [];
	const atTop = depth < 0;
	const span = Math.abs(depth);
	const subtotalLevels = span >= 2 ? Math.min(span - 1, levels - 1) : 0;
	const blanks = (n: number): Scalar[] => new Array<Scalar>(n).fill(null);
	const placeholder: Slot = { labels: [], leaves: [] };
	const visit = (nodes: GroupNode[], level: number, prefix: Scalar[]): number[] => {
		const covered: number[] = [];
		for (const node of nodes) {
			const labels = [...prefix, node.key];
			if (level === levels - 1) {
				const leaf = leafRows.length;
				leafRows.push(node.rows);
				slots.push({ labels, leaves: [leaf] });
				covered.push(leaf);
				continue;
			}
			const subtotal = level < subtotalLevels;
			const at = slots.length;
			if (subtotal && atTop) {
				slots.push(placeholder);
			}
			const inner = visit(node.children, level + 1, labels);
			if (subtotal) {
				const line = { labels: [...labels, ...blanks(levels - 1 - level)], leaves: inner };
				if (atTop) {
					slots[at] = line;
				} else {
					slots.push(line);
				}
			}
			for (const l of inner) {
				covered.push(l);
			}
		}
		return covered;
	};
	if (span >= 1 && atTop) {
		slots.push(placeholder);
	}
	const all = visit(tree, 0, []);
	if (span >= 1) {
		const total = { labels: ['Total', ...blanks(levels - 1)], leaves: all };
		if (atTop) {
			slots[0] = total;
		} else {
			slots.push(total);
		}
	}
	return { slots, leafRows };
}

/** Data rows covered by a slot. */
export function slotRows(slot: Slot, leafRows: readonly number[][]): number[] {
	if (slot.leaves.length === 1) {
		return leafRows[slot.leaves[0]];
	}
	const out: number[] = [];
	for (const l of slot.leaves) {
		for (const r of leafRows[l]) {
			out.push(r);
		}
	}
	return out;
}

export interface HeaderMode {
	/** The first row of the fields / values holds headers. */
	hasHeaders: boolean;
	/** The result starts with a header row. */
	show: boolean;
	/** Header text is generated ("Row Field 1", "Value 1", ...). */
	generate: boolean;
}

/**
 * field_headers: omitted = automatic (the data has headers when its first
 * value is text and its second a number; they are shown when there are
 * several field levels), 0 no, 1 yes but hidden, 2 no but generated,
 * 3 yes and shown.
 */
export function readHeaderMode(
	args: readonly FunctionArg[],
	i: number,
	values: ArrayValue,
	multiLevel: boolean,
	date1904: boolean,
): HeaderMode | CellError {
	if (isOmitted(args, i)) {
		const has =
			values.rows >= 2 &&
			typeof values.values[0][0] === 'string' &&
			typeof values.values[1][0] === 'number';
		return { hasHeaders: has, show: has && multiLevel, generate: false };
	}
	const n = toNumber(scalarArg(args, i), date1904);
	if (isCellError(n)) {
		return n;
	}
	switch (Math.trunc(n)) {
		case 0:
			return { hasHeaders: false, show: false, generate: false };
		case 1:
			return { hasHeaders: true, show: false, generate: false };
		case 2:
			return { hasHeaders: false, show: true, generate: true };
		case 3:
			return { hasHeaders: true, show: true, generate: false };
		default:
			return ERRORS.VALUE;
	}
}

/** total_depth: omitted = grand totals, plus first-level subtotals when there are several levels. */
export function readTotalDepth(
	args: readonly FunctionArg[],
	i: number,
	levels: number,
	date1904: boolean,
): number | CellError {
	if (isOmitted(args, i)) {
		return levels > 1 ? 2 : 1;
	}
	const n = toNumber(scalarArg(args, i), date1904);
	return isCellError(n) ? n : Math.trunc(n);
}

/**
 * sort_order: 1-based column numbers over the fields followed by the value
 * columns, negative for descending. Several numbers may only name fields.
 */
export function readSortSpec(
	args: readonly FunctionArg[],
	i: number,
	fields: number,
	values: number,
	ctx: FunctionContext,
): number[] | CellError {
	if (isOmitted(args, i)) {
		return [];
	}
	const a = toArrayValue(args[i] as FormulaValue, ctx);
	if (isCellError(a)) {
		return a;
	}
	const spec: number[] = [];
	for (const line of a.values) {
		for (const v of line) {
			const n = toNumber(v, ctx.date1904);
			if (isCellError(n)) {
				return n;
			}
			const k = Math.trunc(n);
			if (k === 0 || Math.abs(k) > fields + values) {
				return ERRORS.VALUE;
			}
			spec.push(k);
		}
	}
	if (spec.length > 1 && spec.some((k) => Math.abs(k) > fields)) {
		return ERRORS.VALUE;
	}
	return spec;
}

/** filter_array: one logical per data row (header row included), or null when omitted. */
export function readFilter(
	args: readonly FunctionArg[],
	i: number,
	rows: number,
): boolean[] | null | CellError {
	if (isOmittedArray(args, i)) {
		return null;
	}
	const a = arrayArg(args, i);
	if (isCellError(a)) {
		return a;
	}
	const list = vectorValues(a);
	if (!list || list.length !== rows) {
		return ERRORS.VALUE;
	}
	const out: boolean[] = [];
	for (const v of list) {
		const b = toBoolean(v);
		if (isCellError(b)) {
			return b;
		}
		out.push(b);
	}
	return out;
}

/** Data rows after the header row that pass the filter. */
export function dataRows(
	count: number,
	first: number,
	filter: readonly boolean[] | null,
): number[] {
	const rows: number[] = [];
	for (let r = first; r < count; r++) {
		if (!filter || filter[r]) {
			rows.push(r);
		}
	}
	return rows;
}
