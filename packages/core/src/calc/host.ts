import { cellKey } from '../address/cell-address';
import type { EvaluationHost, HostTable, NameDefinition } from '../formula/evaluator/host';
import type { CellInfoProvider } from '../formula/functions/information/cell';
import type { FunctionRegistry } from '../formula/registry';
import type { ArrayValue, RefValue, Scalar } from '../formula/types';
import type { RangeAddress } from '../types/address';
import type { CellError } from '../types/cell';
import type { Workbook } from '../types/workbook';
import { localDateToSerial } from '../values/dates';
import { ERRORS, isCellError } from '../values/errors';
import { CellDetails } from './cell-details';
import { SheetCellIndex } from './cell-index';
import type { DependencyGraph } from './graph';
import type { NameTable, TableTable } from './lookups';
import type { FormulaNode } from './node';
import { RangeReader } from './range-reader';
import type { CalcEngineOptions } from './types';

/**
 * Thrown when an evaluation reads cells whose formulas still have to be
 * recalculated in the current pass; the scheduler computes them first and
 * retries. It is a plain object, not an Error, so throwing it is cheap.
 */
export class PendingCells {
	constructor(readonly nodes: readonly FormulaNode[]) {}
}

/**
 * The evaluator's view of a workbook: cell access (with the recalculation
 * checks), indexed range reads, names, tables, spill ranges and the CELL()
 * details.
 */
export class WorkbookHost implements EvaluationHost {
	/** Reads of cells that are due for recalculation suspend the evaluation (during passes). */
	checking = false;
	/** NOW() of the current pass (one clock reading per recalculation). */
	private passNow: number | null = null;
	private readonly indexes: SheetCellIndex[] = [];
	private sheetLookup = new Map<string, number>();
	private readonly reader: RangeReader;
	private readonly details: CellDetails;
	readonly cellInfo: CellInfoProvider;

	constructor(
		private readonly workbook: Workbook,
		private readonly graph: DependencyGraph,
		private readonly names: NameTable,
		private readonly tables: TableTable,
		private readonly options: CalcEngineOptions,
		readonly registry: FunctionRegistry,
	) {
		this.reader = new RangeReader(workbook, {
			index: (sheet) => this.index(sheet),
			ensureClean: (sheet, range) => this.ensureClean(sheet, range),
		});
		this.details = new CellDetails(workbook);
		this.cellInfo = this.details;
	}

	get date1904(): boolean {
		return this.workbook.date1904;
	}

	/** Forget cached sheet data (after a rebuild). */
	reset(): void {
		this.indexes.length = 0;
		this.sheetLookup = new Map(this.workbook.sheets.map((s, i) => [s.name.toLowerCase(), i]));
		this.details.reset();
		this.reader.clear();
	}

	beginPass(): void {
		this.checking = true;
		this.passNow = null;
		this.reader.clear();
		this.reader.snapshots = true;
	}

	endPass(): void {
		this.checking = false;
		this.passNow = null;
		this.reader.snapshots = false;
		this.reader.clear();
	}

	/** Cell index of a sheet, synchronised with its cell map. */
	index(sheet: number): SheetCellIndex {
		let idx = this.indexes[sheet];
		if (idx === undefined) {
			idx = new SheetCellIndex();
			this.indexes[sheet] = idx;
		}
		const ws = this.workbook.sheets[sheet];
		if (ws !== undefined) {
			idx.sync(ws.cells);
		}
		return idx;
	}

	/** Record that a cell was stored, replaced or deleted. */
	noteCell(sheet: number, row: number, col: number): void {
		const ws = this.workbook.sheets[sheet];
		const idx = this.indexes[sheet];
		if (ws !== undefined && idx !== undefined) {
			idx.note(ws.cells, row, col);
		}
		this.reader.noteWrite(sheet, row, col, ws?.cells.get(cellKey(row, col))?.value ?? null);
	}

	/** Record that the value of an existing cell was replaced. */
	noteValue(sheet: number, row: number, col: number, value: Scalar): void {
		this.reader.noteWrite(sheet, row, col, value);
	}

	getCellValue(sheet: number, row: number, col: number): Scalar {
		const ws = this.workbook.sheets[sheet];
		if (ws === undefined) {
			return ERRORS.REF;
		}
		const key = cellKey(row, col);
		const cell = ws.cells.get(key);
		if (cell === undefined) {
			return null;
		}
		if (this.checking && this.graph.dirtyTotal > 0) {
			const owner =
				cell.formula !== undefined
					? this.graph.node(sheet, key)
					: cell.arrayAnchor !== undefined
						? this.graph.node(sheet, cell.arrayAnchor)
						: undefined;
			if (owner !== undefined && owner.dirty) {
				throw new PendingCells([owner]);
			}
		}
		return cell.value;
	}

	private ensureClean(sheet: number, range: RangeAddress): void {
		if (!this.checking || this.graph.dirtyTotal === 0) {
			return;
		}
		const pending: FormulaNode[] = [];
		this.graph.forEachDirtyIn(sheet, range, (n) => pending.push(n));
		this.graph.forEachDirtySpillIn(sheet, range, (n) => pending.push(n));
		if (pending.length > 0) {
			throw new PendingCells(pending);
		}
	}

	resolveRef(ref: RefValue): ArrayValue {
		return this.reader.resolveRef(ref);
	}

	forEachCell(
		ref: RefValue,
		visit: (value: Scalar, row: number, col: number, sheet: number) => void,
	): void {
		this.reader.forEachCell(ref, visit);
	}

	memo<T>(ref: RefValue, key: string, build: () => T): T {
		return this.reader.memo(ref, key, build);
	}

	getFormula(sheet: number, row: number, col: number): string | undefined {
		return this.workbook.sheets[sheet]?.cells.get(cellKey(row, col))?.formula;
	}

	isRowHidden(sheet: number, row: number): boolean {
		const ws = this.workbook.sheets[sheet];
		if (ws === undefined) {
			return false;
		}
		const info = ws.rows.get(row);
		return info !== undefined ? info.hidden === true : ws.zeroHeight === true;
	}

	isSubtotalCell(sheet: number, row: number, col: number): boolean {
		return this.graph.node(sheet, cellKey(row, col))?.subtotal === true;
	}

	usedRange(sheet: number): RangeAddress | null {
		return this.workbook.sheets[sheet] === undefined ? null : this.index(sheet).usedRange();
	}

	sheetCount(): number {
		return this.workbook.sheets.length;
	}

	sheetName(index: number): string {
		return this.workbook.sheets[index]?.name ?? '';
	}

	sheetIndex(name: string): number {
		const lower = name.toLowerCase();
		const i = this.sheetLookup.get(lower);
		if (i !== undefined && this.workbook.sheets[i]?.name.toLowerCase() === lower) {
			return i;
		}
		return this.workbook.sheets.findIndex((s) => s.name.toLowerCase() === lower);
	}

	now(): number {
		if (this.passNow === null || !this.checking) {
			const date = this.options.now?.() ?? new Date();
			this.passNow = localDateToSerial(date, this.workbook.date1904);
		}
		return this.passNow;
	}

	random(): number {
		return this.options.random?.() ?? Math.random();
	}

	findName(name: string, sheet: number): NameDefinition | undefined {
		return this.names.find(name, sheet);
	}

	findTable(name: string): HostTable | undefined {
		return this.tables.find(name);
	}

	tableAt(sheet: number, row: number, col: number): HostTable | undefined {
		return this.tables.at(sheet, row, col);
	}

	spillRange(sheet: number, row: number, col: number): RangeAddress | CellError {
		const key = cellKey(row, col);
		const cell = this.workbook.sheets[sheet]?.cells.get(key);
		if (cell === undefined || cell.formula === undefined || cell.array?.dynamic !== true) {
			return ERRORS.REF;
		}
		const node = this.graph.node(sheet, key);
		if (node !== undefined && node.dirty && this.checking) {
			throw new PendingCells([node]);
		}
		if (isCellError(cell.value) && cell.value.code === '#SPILL!') {
			return ERRORS.REF;
		}
		if (node !== undefined && !node.arrayResult && node.spill === null) {
			return ERRORS.REF;
		}
		return cell.array.ref;
	}
}
