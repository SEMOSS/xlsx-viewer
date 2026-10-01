import { describe, expect, it } from 'vitest';

import { setCellInput } from '../edit/cell-input';
import { clearRanges } from '../edit/clear-ops';
import { insertRows } from '../edit/structure-ops';
import { EditTransaction, type UndoEntry } from '../edit/transaction';
import { createWorkbook } from '../sheet/workbook';
import type { Workbook } from '../types/workbook';
import { ERRORS } from '../values/errors';
import { CalcEngine } from './engine';
import { cellAt, locate, valueAt, valuesAt } from './test-utils/workbook';

/** Drive the engine the way the viewer store does: transactions, then recalc by the entry. */
class Session {
	readonly wb: Workbook = createWorkbook();
	readonly engine = new CalcEngine(this.wb);
	private readonly undoStack: UndoEntry[] = [];

	constructor() {
		this.engine.rebuild();
	}

	private recalc(entry: UndoEntry): void {
		if (entry.structural || entry.fullRecalc) {
			this.engine.rebuild();
			this.engine.recalculateAll();
		} else if (entry.changedCells.length > 0) {
			this.engine.onCellsChanged(entry.changedCells);
		}
	}

	type(ref: string, text: string): void {
		const tx = new EditTransaction(this.wb, 'type');
		const { sheet, row, col } = locate(this.wb, ref);
		setCellInput(tx, this.wb, sheet, row, col, text);
		this.commit(tx);
	}

	commit(tx: EditTransaction): void {
		const entry = tx.commit();
		this.recalc(entry);
		this.undoStack.push(entry);
	}

	undo(): void {
		const entry = this.undoStack.pop()!;
		entry.undo();
		this.recalc(entry);
	}
}

describe('editing through transactions', () => {
	it('spills typed formulas and survives undo', () => {
		const s = new Session();
		s.type('A1', '3');
		s.type('B1', '=SEQUENCE(A1)');
		expect(valuesAt(s.wb, 'B1:B3')).toEqual([[1], [2], [3]]);
		s.type('B2', 'x');
		expect(valueAt(s.wb, 'B1')).toEqual(ERRORS.SPILL);
		s.undo();
		expect(valuesAt(s.wb, 'B1:B3')).toEqual([[1], [2], [3]]);
		s.undo();
		expect(valuesAt(s.wb, 'B1:B3')).toEqual([[null], [null], [null]]);
	});

	it('recalculates after clearing and structural edits', () => {
		const s = new Session();
		s.type('A1', '1');
		s.type('A2', '2');
		s.type('C1', '=SUM(A1:A2)');
		s.type('D1', '=C1*10');
		expect(valueAt(s.wb, 'D1')).toBe(30);
		const tx = new EditTransaction(s.wb, 'clear');
		clearRanges(tx, s.wb, 0, [{ top: 1, left: 0, bottom: 1, right: 0 }], 'all');
		s.commit(tx);
		expect(valueAt(s.wb, 'D1')).toBe(10);
		const ins = new EditTransaction(s.wb, 'insert');
		insertRows(ins, s.wb, 0, 0, 2);
		s.commit(ins);
		expect(cellAt(s.wb, 'C3')?.formula).toBe('SUM(A3:A4)');
		s.type('A4', '5');
		expect(valueAt(s.wb, 'D3')).toBe(60);
		s.undo();
		expect(valueAt(s.wb, 'D3')).toBe(10);
	});

	it('reports user-typed unknown functions as #NAME?', () => {
		const s = new Session();
		s.type('A1', '=NOTAFUNCTION(1)');
		expect(valueAt(s.wb, 'A1')).toEqual(ERRORS.NAME);
	});
});
