import { cellKey } from '../address/cell-address';
import type { SheetQualifier } from '../formula/ast';
import { cellNodeRange, rangeNodeRange } from '../formula/evaluator/references';
import { structuredRange, tableDataRange } from '../formula/evaluator/structured';
import type { FunctionRegistry } from '../formula/registry';
import type { RangeAddress } from '../types/address';
import { isCellError } from '../values/errors';
import type { FormulaAnalysis } from './analysis';
import type { NameTable, TableTable, WorkbookName } from './lookups';

/** Concrete precedents of one formula cell. */
export interface ResolvedDeps {
	/** Flattened (sheet, cellKey) pairs, without duplicates. */
	readonly cells: number[];
	readonly ranges: Array<{ sheet: number; range: RangeAddress }>;
	readonly volatile: boolean;
	/** Functions that are neither built in nor defined names (upper case). */
	readonly unsupported: string[];
	/** References another workbook. */
	readonly external: boolean;
	/** Uses a defined name whose formula cannot be parsed. */
	readonly brokenName: boolean;
}

export interface ResolveContext {
	readonly registry: FunctionRegistry;
	readonly names: NameTable;
	readonly tables: TableTable;
	sheetIndex(name: string): number;
}

/** Function names the evaluator handles itself. */
const EVALUATOR_FUNCTIONS: ReadonlySet<string> = new Set(['ANCHORARRAY', 'SINGLE']);

class Collector {
	readonly cells: number[] = [];
	readonly ranges: Array<{ sheet: number; range: RangeAddress }> = [];
	private readonly seenCells = new Set<number>();
	private readonly seenRanges = new Set<string>();
	private readonly seenNames = new Set<WorkbookName>();
	volatile = false;
	external = false;
	brokenName = false;
	readonly unsupported: string[] = [];

	constructor(
		private readonly ctx: ResolveContext,
		private readonly sheet: number,
		private readonly row: number,
		private readonly col: number,
	) {}

	private span(q: SheetQualifier | undefined): [number, number] | null {
		if (q === undefined) {
			return [this.sheet, this.sheet];
		}
		if (q.workbook !== undefined) {
			return null;
		}
		const a = this.ctx.sheetIndex(q.sheet);
		const b = q.sheetEnd === undefined ? a : this.ctx.sheetIndex(q.sheetEnd);
		return a < 0 || b < 0 ? null : [Math.min(a, b), Math.max(a, b)];
	}

	private addCell(sheet: number, row: number, col: number): void {
		const key = cellKey(row, col);
		const packed = sheet * 2 ** 34 + key;
		if (!this.seenCells.has(packed)) {
			this.seenCells.add(packed);
			this.cells.push(sheet, key);
		}
	}

	private addRange(sheet: number, range: RangeAddress): void {
		if (range.top === range.bottom && range.left === range.right) {
			this.addCell(sheet, range.top, range.left);
			return;
		}
		const id = `${sheet}:${range.top}:${range.left}:${range.bottom}:${range.right}`;
		if (!this.seenRanges.has(id)) {
			this.seenRanges.add(id);
			this.ranges.push({ sheet, range });
		}
	}

	private addName(def: WorkbookName): void {
		if (this.seenNames.has(def)) {
			return;
		}
		this.seenNames.add(def);
		if (def.parsed.analysis !== null) {
			this.collect(def.parsed.analysis, true);
		} else {
			this.brokenName = true;
		}
	}

	collect(a: FormulaAnalysis, inName: boolean): void {
		const { row, col } = this;
		this.volatile ||= a.volatile;
		this.external ||= a.external;
		for (const c of a.cells) {
			const span = this.span(c.sheet);
			if (span !== null) {
				const r = cellNodeRange(c, inName, row, col);
				for (let s = span[0]; s <= span[1]; s++) {
					this.addCell(s, r.top, r.left);
				}
			}
		}
		for (const node of a.ranges) {
			const span = this.span(node.sheet);
			if (span !== null) {
				const r = rangeNodeRange(node, inName, row, col);
				for (let s = span[0]; s <= span[1]; s++) {
					this.addRange(s, r);
				}
			}
		}
		for (const n of a.names) {
			const span = this.span(n.sheet);
			if (span === null) {
				continue;
			}
			const def = this.ctx.names.find(n.name, span[0]);
			if (def !== undefined) {
				this.addName(def);
			} else if (n.sheet === undefined) {
				const t = this.ctx.tables.find(n.name);
				const range = t === undefined ? null : tableDataRange(t.table);
				if (t !== undefined && range !== null && !isCellError(range)) {
					this.addRange(t.sheet, range);
				}
			}
		}
		for (const s of a.structured) {
			const t =
				s.table !== undefined
					? this.ctx.tables.find(s.table)
					: this.ctx.tables.at(this.sheet, row, col);
			if (t !== undefined) {
				const range = structuredRange(s, t.table, row);
				if (!isCellError(range)) {
					this.addRange(t.sheet, range);
				}
			}
		}
		for (const name of a.functions) {
			const def = this.ctx.registry.get(name);
			if (def !== undefined) {
				this.volatile ||= def.volatile === true;
				continue;
			}
			if (EVALUATOR_FUNCTIONS.has(name)) {
				continue;
			}
			const named = this.ctx.names.find(name, this.sheet);
			if (named !== undefined) {
				this.addName(named);
			} else if (!this.unsupported.includes(name)) {
				this.unsupported.push(name);
			}
		}
	}
}

/** Resolve the analysis of the formula in (sheet, row, col) to concrete precedents. */
export function resolveDependencies(
	analysis: FormulaAnalysis,
	sheet: number,
	row: number,
	col: number,
	ctx: ResolveContext,
): ResolvedDeps {
	const c = new Collector(ctx, sheet, row, col);
	c.collect(analysis, false);
	return {
		cells: c.cells,
		ranges: c.ranges,
		volatile: c.volatile,
		unsupported: c.unsupported,
		external: c.external,
		brokenName: c.brokenName,
	};
}
