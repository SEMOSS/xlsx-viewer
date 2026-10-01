import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition, Scalar } from '../../types';
import { arrayArg, optionalInteger } from '../lookup/args';
import { resolveAggregator } from './aggregate';
import { headerLabels } from './groupby';
import type { Layout, Slot } from './groups';
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

/** Leaf group of every data row in a layout (-1 for rows that are filtered out). */
function leafIndex(layout: Layout, count: number): Int32Array {
	const out = new Int32Array(count).fill(-1);
	layout.leafRows.forEach((rows, leaf) => {
		for (const r of rows) {
			out[r] = leaf;
		}
	});
	return out;
}

/**
 * PIVOTBY(row_fields, col_fields, values, function, [field_headers],
 * [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order],
 * [filter_array], [relative_to])
 *
 * A pivot of `values` aggregated by row and column groups. The result has
 * one header row per column-field level (the top-left block is blank, or
 * holds the row field headers on its last line when headers are shown), a
 * line per row group and a column per column group (times the number of
 * value columns, with an extra header row of value names when there are
 * several). Combinations without data are blank. Totals and sorting work
 * as in GROUPBY on both axes. relative_to (the base PERCENTOF and
 * two-parameter LAMBDAs receive): 0 column total, 1 row total, 2 grand
 * total; the parent-total modes are not supported (`#VALUE!`).
 */
export const PIVOTBY: FunctionDefinition = {
	name: 'PIVOTBY',
	minArgs: 4,
	maxArgs: 11,
	params: [
		'array',
		'array',
		'array',
		'lazy',
		'value',
		'value',
		'any',
		'value',
		'any',
		'array',
		'value',
	],
	filePrefix: '_xlfn.',
	returnsArray: true,
	impl: (args, ctx) => {
		const rowFields = arrayArg(args, 0);
		if (isCellError(rowFields)) {
			return rowFields;
		}
		const colFields = arrayArg(args, 1);
		if (isCellError(colFields)) {
			return colFields;
		}
		const values = arrayArg(args, 2);
		if (isCellError(values)) {
			return values;
		}
		const count = rowFields.rows;
		if (colFields.rows !== count || values.rows !== count) {
			return ERRORS.VALUE;
		}
		const aggregate = resolveAggregator(args[3], ctx);
		if (isCellError(aggregate)) {
			return aggregate;
		}
		const rowLevels = rowFields.cols;
		const colLevels = colFields.cols;
		const width = values.cols;
		const date1904 = ctx.date1904;
		const header = readHeaderMode(args, 4, values, rowLevels > 1 || colLevels > 1, date1904);
		if (isCellError(header)) {
			return header;
		}
		const rowDepth = readTotalDepth(args, 5, rowLevels, date1904);
		if (isCellError(rowDepth)) {
			return rowDepth;
		}
		const rowSpec = readSortSpec(args, 6, rowLevels, width, ctx);
		if (isCellError(rowSpec)) {
			return rowSpec;
		}
		const colDepth = readTotalDepth(args, 7, colLevels, date1904);
		if (isCellError(colDepth)) {
			return colDepth;
		}
		const colSpec = readSortSpec(args, 8, colLevels, width, ctx);
		if (isCellError(colSpec)) {
			return colSpec;
		}
		const filter = readFilter(args, 9, count);
		if (isCellError(filter)) {
			return filter;
		}
		const relative = optionalInteger(args, 10, 0, date1904);
		if (isCellError(relative)) {
			return relative;
		}
		if (relative < 0 || relative > 2) {
			return ERRORS.VALUE;
		}
		const rows = dataRows(count, header.hasHeaders ? 1 : 0, filter);
		if (rows.length === 0) {
			return ERRORS.CALC;
		}
		const measure = (rs: readonly number[], j: number, base: readonly number[]): Scalar =>
			aggregate(
				rs.map((r) => values.values[r][j]),
				base.map((r) => values.values[r][j]),
			);
		const rowTree = buildGroups(rows, (r, level) => rowFields.values[r][level], rowLevels);
		sortGroups(
			rowTree,
			groupOrder(rowSpec, rowLevels, (n, j) => measure(n.rows, j, rows)),
		);
		const colTree = buildGroups(rows, (r, level) => colFields.values[r][level], colLevels);
		sortGroups(
			colTree,
			groupOrder(colSpec, colLevels, (n, j) => measure(n.rows, j, rows)),
		);
		const rowLayout = layoutGroups(rowTree, rowLevels, rowDepth);
		const colLayout = layoutGroups(colTree, colLevels, colDepth);

		// Bucket the data rows by (row leaf, column leaf); totals are unions of buckets.
		const rowLeaf = leafIndex(rowLayout, count);
		const colLeaf = leafIndex(colLayout, count);
		const columnsOfLeaves = colLayout.leafRows.length;
		const buckets = new Map<number, number[]>();
		for (const r of rows) {
			const key = rowLeaf[r] * columnsOfLeaves + colLeaf[r];
			const bucket = buckets.get(key);
			if (bucket) {
				bucket.push(r);
			} else {
				buckets.set(key, [r]);
			}
		}
		const cellRows = (rs: Slot, cs: Slot): number[] => {
			const out: number[] = [];
			for (const l of rs.leaves) {
				for (const m of cs.leaves) {
					const bucket = buckets.get(l * columnsOfLeaves + m);
					if (bucket) {
						for (const r of bucket) {
							out.push(r);
						}
					}
				}
			}
			return out;
		};

		const out: Scalar[][] = [];
		const blanks = (): Scalar[] => new Array<Scalar>(rowLevels).fill(null);
		for (let level = 0; level < colLevels; level++) {
			const line = blanks();
			for (const cs of colLayout.slots) {
				for (let j = 0; j < width; j++) {
					line.push(cs.labels[level]);
				}
			}
			out.push(line);
		}
		if (width > 1) {
			const labels = headerLabels(values, header, (j) => `Value ${j + 1}`);
			const line = blanks();
			for (let k = 0; k < colLayout.slots.length; k++) {
				line.push(...labels);
			}
			out.push(line);
		}
		if (header.show) {
			const labels = headerLabels(rowFields, header, (c) => `Row Field ${c + 1}`);
			const last = out[out.length - 1];
			for (let c = 0; c < rowLevels; c++) {
				last[c] = labels[c];
			}
		}
		const colBases = colLayout.slots.map((cs) => slotRows(cs, colLayout.leafRows));
		for (const rs of rowLayout.slots) {
			const rowBase = relative === 1 ? slotRows(rs, rowLayout.leafRows) : rows;
			const line = rs.labels.slice();
			colLayout.slots.forEach((cs, k) => {
				const cell = cellRows(rs, cs);
				const base = relative === 0 ? colBases[k] : rowBase;
				for (let j = 0; j < width; j++) {
					line.push(cell.length ? measure(cell, j, base) : null);
				}
			});
			out.push(line);
		}
		return {
			kind: 'array',
			rows: out.length,
			cols: rowLevels + colLayout.slots.length * width,
			values: out,
		};
	},
};
