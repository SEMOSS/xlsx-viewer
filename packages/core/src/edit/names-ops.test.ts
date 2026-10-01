import { describe, expect, it } from 'vitest';

import { parseRange } from '../address/range';
import { createWorkbook } from '../sheet/workbook';
import { getCell } from '../sheet/worksheet';
import type { DefinedName, Workbook } from '../types/workbook';
import {
	addDefinedName,
	compareDefinedNames,
	DEFINED_NAME_DUPLICATE_MESSAGE,
	DEFINED_NAME_INVALID_MESSAGE,
	deleteDefinedName,
	EditTransaction,
	isValidDefinedNameSyntax,
	setCellInput,
	validateDefinedName,
} from './index';

function seed(wb: Workbook, entries: Record<string, string>, sheet = 0): void {
	const tx = new EditTransaction(wb, 'seed');
	for (const [ref, text] of Object.entries(entries)) {
		const r = parseRange(ref)!;
		setCellInput(tx, wb, sheet, r.top, r.left, text);
	}
	tx.commit();
}

function formulaAt(wb: Workbook, sheet: number, ref: string): string | undefined {
	const r = parseRange(ref)!;
	return getCell(wb.sheets[sheet], r.top, r.left)?.formula;
}

function names(wb: Workbook): string[] {
	return wb.definedNames.map((n) => `${n.name}@${n.localSheet ?? 'wb'}=${n.formula}`);
}

describe('isValidDefinedNameSyntax', () => {
	it('accepts names that follow the Excel rules', () => {
		for (const name of [
			'Rate',
			'_x',
			'\\path',
			'Sales.2024',
			'a?b',
			'Umsätze',
			'Größe',
			'AAAA1',
			'XFE1',
			'A1048577',
			'R1C1A',
			'Tax_Rate1',
			'x'.repeat(255),
		]) {
			expect(isValidDefinedNameSyntax(name), name).toBe(true);
		}
	});

	it('rejects bad characters, spaces, references and booleans', () => {
		for (const name of [
			'',
			'Rate 2',
			' Rate',
			'1abc',
			'.abc',
			'?abc',
			'a-b',
			'a$b',
			'A1',
			'a1',
			'$A$1',
			'XFD1048576',
			'AB12',
			'R',
			'r',
			'C',
			'c',
			'R1',
			'C1',
			'RC',
			'R1C1',
			'r2c3',
			'R1C',
			'TRUE',
			'false',
			'x'.repeat(256),
		]) {
			expect(isValidDefinedNameSyntax(name), name).toBe(false);
		}
	});
});

describe('validateDefinedName', () => {
	const wb = createWorkbook({ sheetNames: ['Sheet1', 'Sheet2'] });
	wb.definedNames.push(
		{ name: 'Rate', formula: 'Sheet1!$A$1' },
		{ name: 'Local', formula: 'Sheet2!$A$1', localSheet: 1 },
		{ name: '_xlnm.Print_Area', formula: 'Sheet1!$A$1:$C$9', localSheet: 0, builtin: true },
	);

	it("returns Excel's message for invalid names", () => {
		expect(validateDefinedName(wb, 'my name', undefined)).toBe(DEFINED_NAME_INVALID_MESSAGE);
		expect(DEFINED_NAME_INVALID_MESSAGE).toContain('The name that you entered is not valid.');
		expect(DEFINED_NAME_INVALID_MESSAGE).toContain(
			' - The name contains a space or other invalid characters',
		);
	});

	it('detects duplicates in the same scope, case-insensitively and with built-ins', () => {
		expect(validateDefinedName(wb, 'RATE', undefined)).toBe(DEFINED_NAME_DUPLICATE_MESSAGE);
		expect(validateDefinedName(wb, 'rate', 0)).toBeNull();
		expect(validateDefinedName(wb, 'local', 1)).toBe(DEFINED_NAME_DUPLICATE_MESSAGE);
		expect(validateDefinedName(wb, 'Local', undefined)).toBeNull();
		expect(validateDefinedName(wb, 'print_area', 0)).toBe(DEFINED_NAME_DUPLICATE_MESSAGE);
		expect(validateDefinedName(wb, 'Print_Area', 1)).toBeNull();
	});

	it('ignores the definition being edited', () => {
		expect(validateDefinedName(wb, 'rate', undefined, wb.definedNames[0])).toBeNull();
		expect(validateDefinedName(wb, 'Rate', undefined, { name: 'rate', formula: '' })).toBeNull();
	});

	it('rejects names used by tables', () => {
		const book = createWorkbook();
		book.sheets[0].tables.push({
			id: 1,
			name: 'Table1',
			displayName: 'Sales',
			ref: parseRange('A1:B3')!,
			headerRowCount: 1,
			totalsRowCount: 0,
			columns: [],
		});
		expect(validateDefinedName(book, 'sales', undefined)).toBe(DEFINED_NAME_INVALID_MESSAGE);
		expect(validateDefinedName(book, 'Sales2', undefined)).toBeNull();
	});
});

describe('addDefinedName', () => {
	it('adds a sorted, undoable definition', () => {
		const wb = createWorkbook({ sheetNames: ['Sheet1', 'Sheet2'] });
		wb.dirty = false;
		const tx1 = new EditTransaction(wb, 'Define Name');
		addDefinedName(tx1, wb, { name: 'Zeta', formula: '=Sheet1!$A$1' });
		tx1.commit();
		const tx = new EditTransaction(wb, 'Define Name');
		const stored = addDefinedName(tx, wb, {
			name: 'alpha',
			formula: 'Sheet1!$B$2',
			comment: 'First',
		});
		const entry = tx.commit();
		expect(stored).toEqual({ name: 'alpha', formula: 'Sheet1!$B$2', comment: 'First' });
		expect(names(wb)).toEqual(['alpha@wb=Sheet1!$B$2', 'Zeta@wb=Sheet1!$A$1']);
		expect(wb.dirty).toBe(true);
		expect(entry.structural).toBe(true);

		entry.undo();
		expect(names(wb)).toEqual(['Zeta@wb=Sheet1!$A$1']);
		entry.redo();
		expect(names(wb)).toEqual(['alpha@wb=Sheet1!$B$2', 'Zeta@wb=Sheet1!$A$1']);
	});

	it('stores sheet-scoped Print_Area as the built-in name', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Define Name');
		const stored = addDefinedName(tx, wb, {
			name: 'print_area',
			formula: 'Sheet1!$A$1:$B$2',
			localSheet: 0,
		});
		expect(stored).toMatchObject({ name: '_xlnm.Print_Area', builtin: true, localSheet: 0 });
	});

	it('throws the Excel message and leaves the workbook alone when rejected', () => {
		const wb = createWorkbook();
		wb.definedNames.push({ name: 'Rate', formula: '0.05' });
		const tx = new EditTransaction(wb, 'Define Name');
		expect(() => addDefinedName(tx, wb, { name: 'rate', formula: '1' })).toThrow(
			DEFINED_NAME_DUPLICATE_MESSAGE,
		);
		expect(() => addDefinedName(tx, wb, { name: 'A1', formula: '1' })).toThrow(/not valid/);
		expect(wb.definedNames).toHaveLength(1);
	});

	it('orders names case-insensitively, workbook scope first', () => {
		const list: DefinedName[] = [
			{ name: 'b', formula: '1', localSheet: 0 },
			{ name: 'B', formula: '1' },
			{ name: '_xlnm.Print_Area', formula: '1', localSheet: 0 },
			{ name: 'a', formula: '1' },
		];
		expect(list.sort(compareDefinedNames).map((n) => `${n.name}@${n.localSheet ?? 'wb'}`)).toEqual([
			'a@wb',
			'B@wb',
			'b@0',
			'_xlnm.Print_Area@0',
		]);
	});
});

function scopedWorkbook(): Workbook {
	const wb = createWorkbook({ sheetNames: ['Sheet1', 'Sheet2', 'Sheet3'] });
	wb.definedNames.push(
		{ name: 'Rate', formula: 'Sheet1!$A$1' },
		{ name: 'Rate', formula: 'Sheet2!$A$1', localSheet: 1 },
		{ name: 'Total', formula: 'Rate*2' },
		{ name: 'LocalTotal', formula: 'Rate*3', localSheet: 1 },
	);
	seed(wb, { B1: '=Rate*2', B2: '="Rate"&Rate', B3: '=Sheet2!Rate' }, 0);
	seed(wb, { B1: '=Rate+1', B2: '=Sheet1!Rate' }, 1);
	seed(wb, { B1: '=SUM(Rate,1)' }, 2);
	return wb;
}

describe('deleteDefinedName', () => {
	it('removes the definition, keeps formulas and undoes', () => {
		const wb = scopedWorkbook();
		const tx = new EditTransaction(wb, 'Delete Name');
		expect(deleteDefinedName(tx, wb, { name: 'RATE', formula: '' })).toBe(true);
		const entry = tx.commit();
		expect(names(wb)).toEqual(['Rate@1=Sheet2!$A$1', 'Total@wb=Rate*2', 'LocalTotal@1=Rate*3']);
		expect(formulaAt(wb, 0, 'B1')).toBe('Rate*2');
		entry.undo();
		expect(wb.definedNames).toHaveLength(4);
		expect(wb.definedNames[0]).toEqual({ name: 'Rate', formula: 'Sheet1!$A$1' });
		entry.redo();
		expect(wb.definedNames).toHaveLength(3);
	});

	it('returns false for an unknown name or scope', () => {
		const wb = scopedWorkbook();
		const tx = new EditTransaction(wb, 'Delete Name');
		expect(deleteDefinedName(tx, wb, { name: 'Total', formula: '', localSheet: 0 })).toBe(false);
		expect(deleteDefinedName(tx, wb, { name: 'Nope', formula: '' })).toBe(false);
		expect(wb.definedNames).toHaveLength(4);
	});
});
