import { describe, expect, it } from 'vitest';

import { createWorkbook } from '../sheet/workbook';
import { setCellInput } from './cell-input';
import { insertRows } from './structure-ops';
import { setColumnWidth, setFreezePanes } from './layout-ops';
import { addSheet } from './sheet-ops';
import { applyStylePatch } from './style-ops';
import { sortRange } from './sort-ops';
import {
	createProtectionGuard,
	isActionAllowed,
	isCellLocked,
	ProtectedError,
	rangeHasLockedCells,
} from './protection';
import {
	createProtectionHash,
	legacyPasswordHash,
	verifyProtectionPassword,
} from './protection-password';
import { protectSheet, setWorkbookProtection, unprotectSheet } from './protection-ops';
import { EditTransaction } from './transaction';

const A1 = { top: 0, left: 0, bottom: 0, right: 0 };
const B2 = { top: 1, left: 1, bottom: 1, right: 1 };

function setup() {
	const wb = createWorkbook();
	const tx = new EditTransaction(wb, 'Setup');
	setCellInput(tx, wb, 0, 0, 0, 'locked');
	setCellInput(tx, wb, 0, 1, 1, 'open');
	applyStylePatch(tx, wb, 0, [B2], { protection: { locked: false } });
	protectSheet(tx, wb, 0, { sheet: true, formatColumns: false });
	tx.commit();
	return wb;
}

const guarded = (wb: ReturnType<typeof createWorkbook>, fn: (tx: EditTransaction) => void) => {
	const tx = new EditTransaction(wb, 'Guarded');
	tx.guard = createProtectionGuard(wb);
	fn(tx);
	return tx.commit();
};

describe('sheet protection', () => {
	it('knows locked cells and allowed actions', () => {
		const wb = setup();
		const sheet = wb.sheets[0];
		expect(isCellLocked(wb, sheet, 0, 0)).toBe(true);
		expect(isCellLocked(wb, sheet, 1, 1)).toBe(false);
		expect(rangeHasLockedCells(wb, sheet, B2)).toBe(false);
		expect(rangeHasLockedCells(wb, sheet, { top: 0, left: 0, bottom: 1048575, right: 0 })).toBe(
			true,
		);
		expect(isActionAllowed(sheet, 'formatColumns')).toBe(true);
		expect(isActionAllowed(sheet, 'formatCells')).toBe(false);
		expect(isActionAllowed(sheet, 'selectLockedCells')).toBe(true);
	});

	it('lets values change only in unlocked cells', () => {
		const wb = setup();
		guarded(wb, (tx) => setCellInput(tx, wb, 0, 1, 1, 'typed'));
		expect(wb.sheets[0].cells.get(16384 + 1)?.value).toBe('typed');
		expect(() => guarded(wb, (tx) => setCellInput(tx, wb, 0, 0, 0, 'nope'))).toThrow(
			ProtectedError,
		);
	});

	it('blocks formatting, structure and sorting unless allowed; view changes pass', () => {
		const wb = setup();
		expect(() =>
			guarded(wb, (tx) => applyStylePatch(tx, wb, 0, [B2], { font: { bold: true } })),
		).toThrow(ProtectedError);
		expect(() => guarded(wb, (tx) => insertRows(tx, wb, 0, 3, 1))).toThrow(ProtectedError);
		// Sorting needs the Sort permission even when every cell is unlocked.
		const B2B3 = { top: 1, left: 1, bottom: 2, right: 1 };
		const prep = new EditTransaction(wb, 'Prep');
		applyStylePatch(prep, wb, 0, [B2B3], { protection: { locked: false } });
		setCellInput(prep, wb, 0, 2, 1, 'another');
		prep.commit();
		expect(() =>
			guarded(wb, (tx) => sortRange(tx, wb, 0, B2B3, [{ index: 1, descending: false }])),
		).toThrow(ProtectedError);
		guarded(wb, (tx) => setColumnWidth(tx, wb, 0, 0, 0, 20));
		guarded(wb, (tx) => setFreezePanes(tx, wb, 0, { row: 1, col: 0 }));
		expect(wb.sheets[0].view.pane?.state).toBe('frozen');
	});

	it('unprotects with undo', () => {
		const wb = setup();
		const entry = guarded(wb, (tx) => unprotectSheet(tx, wb, 0));
		expect(wb.sheets[0].protection).toBeUndefined();
		guarded(wb, (tx) => setCellInput(tx, wb, 0, 0, 0, 'free'));
		entry.undo();
		expect(wb.sheets[0].protection?.sheet).toBe(true);
	});

	it('protects the workbook structure', () => {
		const wb = createWorkbook();
		guarded(wb, (tx) => setWorkbookProtection(tx, wb, { lockStructure: true }));
		expect(() => guarded(wb, (tx) => addSheet(tx, wb))).toThrow(/Workbook is protected/);
		guarded(wb, (tx) => setCellInput(tx, wb, 0, 0, 0, 'cells are fine'));
	});
});

describe('protection passwords', () => {
	it('computes Excel legacy hashes', () => {
		expect(legacyPasswordHash('password')).toBe('83AF');
		expect(legacyPasswordHash('test')).toBe('CBEB');
		expect(legacyPasswordHash('secret')).toBe('DAA7');
	});

	it('verifies SHA-512 salted hashes (reference computed independently)', async () => {
		const protection = {
			algorithmName: 'SHA-512',
			saltValue: 'AAECAwQFBgcICQoLDA0ODw==',
			hashValue:
				'beRsOEW5hgnfGpC0r+gdvIpkWH0Z1MMQdzvbTFscKwUs/P8cgikVLIjRJliVxPNFIY8IKWPw9fTenrNFQcPUYQ==',
			spinCount: 1000,
		};
		expect(await verifyProtectionPassword(protection, 'secret')).toBe(true);
		expect(await verifyProtectionPassword(protection, 'Secret')).toBe(false);
		expect(await verifyProtectionPassword({ password: '83AF' }, 'password')).toBe(true);
		expect(await verifyProtectionPassword({}, 'anything')).toBe(true);
	});

	it('creates hashes that verify', async () => {
		const hash = await createProtectionHash('pa55', 2000);
		expect(hash.algorithmName).toBe('SHA-512');
		expect(await verifyProtectionPassword(hash, 'pa55')).toBe(true);
		expect(await verifyProtectionPassword(hash, 'pass')).toBe(false);
	});
});
