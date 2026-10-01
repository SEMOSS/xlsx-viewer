/** Shared helpers for the chart tests. */
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { cellKey } from '../../address/cell-address';
import { createWorkbook } from '../../sheet/workbook';
import type { CellValue } from '../../types/cell';
import type { ChartAxis, ChartModel, ChartPlot } from '../../types/chart';
import type { Workbook } from '../../types/workbook';
import type { ChartRenderNode, ChartRenderTree } from '../layout/types';

const FIXTURES = fileURLToPath(new URL('../__fixtures__/', import.meta.url));

/** Names of every chart fixture. */
export function fixtureNames(): string[] {
	return readdirSync(FIXTURES)
		.filter((f) => f.endsWith('.xml'))
		.sort();
}

export function loadFixture(name: string): string {
	return readFileSync(FIXTURES + name, 'utf8');
}

/**
 * Workbook with sheets filled from row arrays (row 0 = Excel row 1). Values
 * may carry a number format: `{ v: 45000, fmt: 'm/d/yyyy' }`.
 */
export function makeWorkbook(
	sheets: Record<string, Array<Array<CellValue | { v: CellValue; fmt: string }>>>,
): Workbook {
	const names = Object.keys(sheets);
	const wb = createWorkbook({ sheetNames: names });
	names.forEach((name, i) => {
		const sheet = wb.sheets[i];
		sheets[name].forEach((row, r) => {
			row.forEach((value, c) => {
				if (value !== null && typeof value === 'object' && 'fmt' in value) {
					const numFmtId = 164 + Object.keys(wb.styles.numFmts).length;
					wb.styles.numFmts[numFmtId] = value.fmt;
					wb.styles.cellXfs.push({ numFmtId, fontId: 0, fillId: 0, borderId: 0, xfId: 0 });
					sheet.cells.set(cellKey(r, c), { value: value.v, style: wb.styles.cellXfs.length - 1 });
				} else if (value !== null) {
					sheet.cells.set(cellKey(r, c), { value });
				}
			});
		});
	});
	return wb;
}

/** Every node of a render tree, depth first. */
export function flatten(node: ChartRenderNode | ChartRenderTree): ChartRenderNode[] {
	const root = 'root' in node ? node.root : node;
	const out: ChartRenderNode[] = [];
	const visit = (n: ChartRenderNode) => {
		out.push(n);
		if (n.type === 'group') {
			n.children.forEach(visit);
		}
	};
	visit(root);
	return out;
}

/** Nodes with a role (excluding groups unless asked). */
export function byRole(
	tree: ChartRenderTree,
	role: string,
	includeGroups = false,
): ChartRenderNode[] {
	return flatten(tree).filter((n) => n.role === role && (includeGroups || n.type !== 'group'));
}

/** Text of every text node with a role. */
export function texts(tree: ChartRenderTree, role: string): string[] {
	return byRole(tree, role)
		.filter((n): n is Extract<ChartRenderNode, { type: 'text' }> => n.type === 'text')
		.map((n) => n.lines.join('\n'));
}

/** Throw if any coordinate in the tree is not a finite number. */
export function assertFinite(tree: ChartRenderTree): void {
	for (const n of flatten(tree)) {
		for (const [k, v] of Object.entries(n)) {
			if (typeof v === 'number' && !Number.isFinite(v)) {
				throw new Error(`Non-finite ${k} on ${n.type} ${n.role ?? ''}`);
			}
		}
		if (n.type === 'path' && /NaN|Infinity/.test(n.d)) {
			throw new Error(`Non-finite path on ${n.role ?? ''}: ${n.d}`);
		}
	}
}

/** A chart with literal (cached) data: one series per entry of `values`. */
export function literalChart(
	plot: Omit<ChartPlot, 'series' | 'axisIds'> & { axisIds?: number[] },
	values: Array<Array<number | null>>,
	options: { categories?: string[]; names?: string[]; model?: Partial<ChartModel> } = {},
): ChartModel {
	const xy = plot.type === 'scatter' || plot.type === 'bubble';
	const pieLike = plot.type === 'pie' || plot.type === 'doughnut' || plot.type === 'ofPie';
	const axes: ChartAxis[] = pieLike
		? []
		: [
				{
					id: 1,
					kind: xy ? 'val' : 'cat',
					position: plot.barDirection === 'bar' ? 'l' : 'b',
					crossAxisId: 2,
				},
				{ id: 2, kind: 'val', position: plot.barDirection === 'bar' ? 'b' : 'l', crossAxisId: 1 },
			];
	return {
		plots: [
			{
				...plot,
				axisIds: plot.axisIds ?? axes.map((a) => a.id),
				series: values.map((v, i) => ({
					idx: i,
					order: i,
					name: { text: options.names?.[i] ?? `S${i + 1}` },
					categories: options.categories ? { kind: 'str', cache: options.categories } : undefined,
					values: { kind: 'literal', cache: v, formatCode: 'General' },
				})),
			},
		],
		axes,
		...options.model,
	};
}
