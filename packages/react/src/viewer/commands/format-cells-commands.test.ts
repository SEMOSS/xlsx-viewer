import { describe, expect, it } from 'vitest';
import {
	cellKey,
	createWorkbook,
	EditTransaction,
	applyStylePatch,
	numberFormatCode,
	parseRange,
} from 'xlsx-viewer-core';

import { selectRange, typeCells } from '../components/dialogs/__tests__/render-dialog';
import { ViewerStore } from '../state/viewer-store';
import {
	applyFormatCells,
	formatCellsChanges,
	MIXED,
	readFormatCells,
} from './format-cells-commands';

function storeWith(entries: Record<string, string>) {
	const wb = createWorkbook();
	typeCells(wb, entries);
	return new ViewerStore(wb);
}

function bold(store: ViewerStore, ref: string) {
	const tx = new EditTransaction(store.workbook, 'bold');
	applyStylePatch(tx, store.workbook, 0, [parseRange(ref)!], { font: { bold: true } });
	tx.commit();
}

describe('readFormatCells', () => {
	it('reports uniform values and MIXED where the selection differs', () => {
		const store = storeWith({ A1: '1', A2: '2' });
		bold(store, 'A1');
		selectRange(store, 'A1:A2');
		const { values, multiRow, multiCol } = readFormatCells(store);
		expect(values.bold).toBe(MIXED);
		expect(values.italic).toBe(false);
		expect(values.fontName).toBe('Calibri');
		expect(values.horizontal).toBe('general');
		expect(values.locked).toBe(true);
		expect(values.merged).toBe(false);
		expect(multiRow).toBe(true);
		expect(multiCol).toBe(false);
	});

	it('samples the active cell value and format', () => {
		const store = storeWith({ B2: '12.5%' });
		selectRange(store, 'B2');
		const ctx = readFormatCells(store);
		expect(ctx.sampleValue).toBe(0.125);
		expect(ctx.activeFormat).toBe('0.00%');
	});
});

describe('formatCellsChanges', () => {
	it('includes only the properties that changed', () => {
		const store = storeWith({ A1: 'x' });
		const { values } = readFormatCells(store);
		const changes = formatCellsChanges(values, {
			...values,
			bold: true,
			wrapText: true,
			horizontal: 'general',
		});
		expect(changes.patch).toEqual({ font: { bold: true }, alignment: { wrapText: true } });
		expect(changes.borders).toBeUndefined();
		expect(changes.merge).toBeUndefined();
	});

	it('applies explicit choices over mixed values and maps defaults to removals', () => {
		const store = storeWith({ A1: '1', A2: '2' });
		bold(store, 'A1');
		selectRange(store, 'A1:A2');
		const { values } = readFormatCells(store);
		const changes = formatCellsChanges(values, {
			...values,
			bold: false,
			locked: false,
			underline: 'double',
		});
		expect(changes.patch).toEqual({
			font: { bold: undefined, underline: 'double' },
			protection: { locked: false },
		});
	});

	it('passes border changes and merge toggles through', () => {
		const store = storeWith({});
		const { values } = readFormatCells(store);
		const changes = formatCellsChanges(
			values,
			{ ...values, merged: true },
			{ top: { style: 'thin' } },
		);
		expect(changes.borders).toEqual({ top: { style: 'thin' } });
		expect(changes.merge).toBe(true);
	});
});

describe('applyFormatCells', () => {
	it('formats every selected range in one undoable step', () => {
		const store = storeWith({ A1: '1', C3: '3' });
		store.setSelection({
			active: { row: 0, col: 0 },
			anchor: { row: 0, col: 0 },
			ranges: [parseRange('A1:B2')!, parseRange('C3')!],
		});
		const applied = applyFormatCells(store, {
			patch: { numFmt: '0.00', font: { italic: true } },
			borders: { bottom: { style: 'thick' } },
			merge: true,
		});
		expect(applied).toBe(true);
		const wb = store.workbook;
		const sheet = wb.sheets[0];
		const xf = (ref: string) => {
			const r = parseRange(ref)!;
			return wb.styles.cellXfs[sheet.cells.get(cellKey(r.top, r.left))?.style ?? 0];
		};
		expect(numberFormatCode(xf('C3').numFmtId, wb.styles.numFmts)).toBe('0.00');
		expect(wb.styles.fonts[xf('A1').fontId].italic).toBe(true);
		expect(wb.styles.borders[xf('B2').borderId].bottom).toEqual({ style: 'thick' });
		expect(sheet.merges).toEqual([parseRange('A1:B2')]);
		expect(store.getSnapshot().history.undo).toBe('Format Cells');
		store.undo();
		expect(sheet.merges).toEqual([]);
		expect(sheet.cells.get(cellKey(0, 0))?.style).toBeUndefined();
	});

	it('does nothing without changes', () => {
		const store = storeWith({ A1: '1' });
		expect(applyFormatCells(store, { patch: {} })).toBe(false);
		expect(store.getSnapshot().history.undo).toBeNull();
	});
});
