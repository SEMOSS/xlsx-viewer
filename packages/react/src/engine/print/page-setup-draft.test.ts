import { describe, expect, it } from 'vitest';
import { createWorkbook, EditTransaction, getPrintArea, getPrintTitles } from 'xlsx-viewer-core';

import { applyDraft, draftFromSheet, validateDraft } from './page-setup-draft';

describe('page setup draft', () => {
	it('reads the sheet into dialog fields', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		sheet.pageSetup = {
			orientation: 'landscape',
			fitToPage: true,
			fitToWidth: 1,
			fitToHeight: 0,
			paperSize: 9,
		};
		sheet.pageMargins = {
			left: 0.25,
			right: 0.25,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3,
		};
		sheet.printOptions = { gridLines: true, horizontalCentered: true };
		sheet.headerFooter = { oddHeader: '&CPage &P', differentFirst: true, firstHeader: 'Cover' };
		wb.definedNames.push(
			{
				name: '_xlnm.Print_Area',
				formula: 'Sheet1!$A$1:$D$20,Sheet1!$F$1',
				localSheet: 0,
				builtin: true,
			},
			{
				name: '_xlnm.Print_Titles',
				formula: 'Sheet1!$A:$B,Sheet1!$1:$2',
				localSheet: 0,
				builtin: true,
			},
		);
		const d = draftFromSheet(wb, 0);
		expect(d).toMatchObject({
			orientation: 'landscape',
			scaleMode: 'fit',
			fitWidth: '1',
			fitHeight: '',
			paperSize: 9,
			firstPageNumber: 'Auto',
			centerH: true,
			gridLines: true,
			differentFirst: true,
			printArea: '$A$1:$D$20,$F$1',
			titleRows: '$1:$2',
			titleCols: '$A:$B',
			pageOrder: 'downThenOver',
		});
		expect(d.margins).toEqual({
			top: '0.75',
			bottom: '0.75',
			left: '0.25',
			right: '0.25',
			header: '0.3',
			footer: '0.3',
		});
		expect(d.texts.oddHeader).toBe('&CPage &P');
		expect(d.texts.firstHeader).toBe('Cover');
	});

	it("validates entries with Excel's messages", () => {
		const wb = createWorkbook();
		const base = draftFromSheet(wb, 0);
		expect(validateDraft(base, 'Sheet1')).toBeNull();
		expect(validateDraft({ ...base, scale: '5' }, 'Sheet1')).toMatch(/10 and 400/);
		expect(validateDraft({ ...base, scaleMode: 'fit', fitWidth: 'x' }, 'Sheet1')).toMatch(/pages/);
		expect(validateDraft({ ...base, firstPageNumber: 'abc' }, 'Sheet1')).toMatch(/first page/);
		expect(validateDraft({ ...base, margins: { ...base.margins, left: '-1' } }, 'Sheet1')).toMatch(
			/Margins/,
		);
		expect(
			validateDraft({ ...base, margins: { ...base.margins, left: '5', right: '5' } }, 'Sheet1'),
		).toBe('Margins do not fit page size.');
		expect(validateDraft({ ...base, printArea: 'A1:B2,(' }, 'Sheet1')).toBe(
			"Reference isn't valid.",
		);
		expect(validateDraft({ ...base, titleRows: 'A:B' }, 'Sheet1')).toBe("Reference isn't valid.");
	});

	it('applies every tab in one transaction', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		const d = draftFromSheet(wb, 0);
		const tx = new EditTransaction(wb, 'Page Setup');
		applyDraft(tx, wb, 0, {
			...d,
			orientation: 'landscape',
			scale: '80',
			paperSize: 5,
			firstPageNumber: '3',
			margins: { ...d.margins, top: '1.5' },
			centerV: true,
			texts: { ...d.texts, oddFooter: '&CPage &P of &N', evenFooter: 'ignored' },
			printArea: 'A1:C9',
			titleRows: '$1:$1',
			headings: true,
			blackAndWhite: true,
			pageOrder: 'overThenDown',
		});
		const entry = tx.commit();
		expect(sheet.pageSetup).toEqual({
			orientation: 'landscape',
			paperSize: 5,
			scale: 80,
			firstPageNumber: 3,
			useFirstPageNumber: true,
			pageOrder: 'overThenDown',
			blackAndWhite: true,
		});
		expect(sheet.pageMargins?.top).toBe(1.5);
		expect(sheet.printOptions).toEqual({ headings: true, verticalCentered: true });
		expect(sheet.headerFooter).toEqual({ oddFooter: '&CPage &P of &N' });
		expect(getPrintArea(wb, 0)).toEqual([{ top: 0, left: 0, bottom: 8, right: 2 }]);
		expect(getPrintTitles(wb, 0)).toEqual({ rows: { start: 0, end: 0 } });
		entry.undo();
		expect(sheet.pageSetup).toBeUndefined();
		expect(wb.definedNames).toEqual([]);
	});

	it('switches to Fit to and back', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		const d = draftFromSheet(wb, 0);
		let tx = new EditTransaction(wb, 'fit');
		applyDraft(tx, wb, 0, { ...d, scaleMode: 'fit', fitWidth: '1', fitHeight: '' });
		tx.commit();
		expect(sheet.pageSetup).toMatchObject({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 });
		tx = new EditTransaction(wb, 'adjust');
		applyDraft(tx, wb, 0, { ...draftFromSheet(wb, 0), scaleMode: 'adjust', scale: '90' });
		tx.commit();
		expect(sheet.pageSetup).toMatchObject({ fitToPage: false, scale: 90 });
	});
});
