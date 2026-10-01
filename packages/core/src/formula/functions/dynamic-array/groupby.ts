import { ERRORS, isCellError } from '../../../values/errors';
import type { ArrayValue, FunctionDefinition, Scalar } from '../../types';
import { arrayArg, optionalInteger } from '../lookup/args';
import { resolveAggregator } from './aggregate';
import type { GroupNode, HeaderMode } from './groups';
import {
	buildGroups,
	dataRows,
	groupOrder,
	layoutGroups,
	readFilter,
	readHeaderMode,
	readSortSpec,
	readTotalDepth,
	slotRows,
	sortGroups,
} from './groups';
import { compareForSort } from './sort-order';

/** Leaf paths of a group tree, each as its chain of nodes from the first level down. */
function leafPaths(
	nodes: GroupNode[],
	prefix: GroupNode[] = [],
	out: GroupNode[][] = [],
): GroupNode[][] {
	for (const node of nodes) {
		const path = [...prefix, node];
		if (node.children.length === 0) {
			out.push(path);
		} else {
			leafPaths(node.children, path, out);
		}
	}
	return out;
}

/**
 * field_relationship 1 (table): every distinct key combination is a row and
 * fields sort independently of each other (the sort spec first, then each
 * field ascending). Returned as single-child chains so the layout keeps the
 * order; subtotals do not apply.
 */
function tableTree(
	tree: GroupNode[],
	spec: readonly number[],
	fields: number,
	measure: (rows: readonly number[], column: number) => Scalar,
): GroupNode[] {
	const desc: boolean[] = [];
	const pick: Array<(path: GroupNode[]) => Scalar> = [];
	for (const k of spec) {
		const i = Math.abs(k) - 1;
		pick.push(i < fields ? (p) => p[i].key : (p) => measure(p[p.length - 1].rows, i - fields));
		desc.push(k < 0);
	}
	for (let i = 0; i < fields; i++) {
		pick.push((p) => p[i].key);
		desc.push(false);
	}
	const decorated = leafPaths(tree).map((path) => ({ path, keys: pick.map((f) => f(path)) }));
	decorated.sort((a, b) => {
		for (let k = 0; k < desc.length; k++) {
			const c = compareForSort(a.keys[k], b.keys[k], desc[k]);
			if (c !== 0) {
				return c;
			}
		}
		return 0;
	});
	return decorated.map(({ path }) => {
		const rows = path[path.length - 1].rows;
		let node: GroupNode | null = null;
		for (let i = path.length - 1; i >= 0; i--) {
			node = { key: path[i].key, rows, children: node ? [node] : [] };
		}
		return node as GroupNode;
	});
}

/** Header labels of a fields / values block: from the data, or generated. */
export function headerLabels(
	block: ArrayValue,
	mode: HeaderMode,
	generated: (index: number) => string,
): Scalar[] {
	const out: Scalar[] = [];
	for (let c = 0; c < block.cols; c++) {
		out.push(mode.generate ? generated(c) : mode.hasHeaders ? block.values[0][c] : null);
	}
	return out;
}

/**
 * GROUPBY(row_fields, values, function, [field_headers], [total_depth],
 * [sort_order], [filter_array], [field_relationship])
 *
 * Groups the rows of `values` by the key columns of `row_fields` (several
 * columns form a hierarchy) and aggregates each value column with
 * `function`: an eta-reduced name such as SUM, AVERAGE, COUNT or PERCENTOF,
 * or a LAMBDA. Keys sort ascending unless sort_order says otherwise; totals
 * follow total_depth (default: a "Total" line, plus first-level subtotals
 * for hierarchies). Not supported: arrays of functions.
 */
export const GROUPBY: FunctionDefinition = {
	name: 'GROUPBY',
	minArgs: 3,
	maxArgs: 8,
	params: ['array', 'array', 'lazy', 'value', 'value', 'any', 'array', 'value'],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => {
		const fields = arrayArg(args, 0);
		if (isCellError(fields)) {
			return fields;
		}
		const values = arrayArg(args, 1);
		if (isCellError(values)) {
			return values;
		}
		if (fields.rows !== values.rows) {
			return ERRORS.VALUE;
		}
		const aggregate = resolveAggregator(args[2], ctx);
		if (isCellError(aggregate)) {
			return aggregate;
		}
		const levels = fields.cols;
		const header = readHeaderMode(args, 3, values, levels > 1, ctx.date1904);
		if (isCellError(header)) {
			return header;
		}
		let depth = readTotalDepth(args, 4, levels, ctx.date1904);
		if (isCellError(depth)) {
			return depth;
		}
		const spec = readSortSpec(args, 5, levels, values.cols, ctx);
		if (isCellError(spec)) {
			return spec;
		}
		const filter = readFilter(args, 6, fields.rows);
		if (isCellError(filter)) {
			return filter;
		}
		const relationship = optionalInteger(args, 7, 0, ctx.date1904);
		if (isCellError(relationship)) {
			return relationship;
		}
		if (relationship !== 0 && relationship !== 1) {
			return ERRORS.VALUE;
		}
		const rows = dataRows(fields.rows, header.hasHeaders ? 1 : 0, filter);
		if (rows.length === 0) {
			return ERRORS.CALC;
		}
		const all = Array.from({ length: values.cols }, (_, j) => rows.map((r) => values.values[r][j]));
		const measure = (rs: readonly number[], j: number): Scalar =>
			aggregate(
				rs.map((r) => values.values[r][j]),
				all[j],
			);
		let tree = buildGroups(rows, (r, level) => fields.values[r][level], levels);
		if (relationship === 1) {
			tree = tableTree(tree, spec, levels, measure);
			depth = Math.sign(depth) * Math.min(Math.abs(depth), 1);
		} else {
			sortGroups(
				tree,
				groupOrder(spec, levels, (node, j) => measure(node.rows, j)),
			);
		}
		const { slots, leafRows } = layoutGroups(tree, levels, depth);
		const out: Scalar[][] = [];
		if (header.show) {
			out.push([
				...headerLabels(fields, header, (c) => `Row Field ${c + 1}`),
				...headerLabels(values, header, (j) => `Value ${j + 1}`),
			]);
		}
		for (const slot of slots) {
			const rs = slotRows(slot, leafRows);
			const line = slot.labels.slice();
			for (let j = 0; j < values.cols; j++) {
				line.push(measure(rs, j));
			}
			out.push(line);
		}
		return { kind: 'array', rows: out.length, cols: levels + values.cols, values: out };
	},
};
