import { describe, expect, it } from 'vitest';

import { cellKey } from '../address/cell-address';
import { createWorkbook } from '../sheet/workbook';
import { DOCUMENT_TEXT_PROPERTIES, setDocumentProperties, workbookStatistics } from './index';

function cleanWorkbook() {
	const wb = createWorkbook({ sheetNames: ['Data', 'Calc'] });
	wb.properties = {
		title: 'Budget',
		creator: 'Ann',
		created: '2024-01-01T10:00:00Z',
		application: 'Microsoft Excel',
	};
	wb.dirty = false;
	return wb;
}

describe('setDocumentProperties', () => {
	it('sets trimmed text properties and marks the workbook dirty', () => {
		const wb = cleanWorkbook();
		const before = wb.properties;
		expect(
			setDocumentProperties(wb, { subject: '  Q3 numbers ', company: 'Contoso', keywords: 'a; b' }),
		).toBe(true);
		expect(wb.properties).toEqual({
			title: 'Budget',
			creator: 'Ann',
			created: '2024-01-01T10:00:00Z',
			application: 'Microsoft Excel',
			subject: 'Q3 numbers',
			company: 'Contoso',
			keywords: 'a; b',
		});
		expect(wb.dirty).toBe(true);
		// The previous object is replaced, not mutated.
		expect(before.subject).toBeUndefined();
	});

	it('removes properties set to empty text', () => {
		const wb = cleanWorkbook();
		expect(setDocumentProperties(wb, { title: '   ', creator: '' })).toBe(true);
		expect('title' in wb.properties).toBe(false);
		expect('creator' in wb.properties).toBe(false);
		expect(wb.properties.created).toBe('2024-01-01T10:00:00Z');
	});

	it('reports no change for identical or only differently padded values', () => {
		const wb = cleanWorkbook();
		expect(
			setDocumentProperties(wb, {
				title: ' Budget ',
				creator: 'Ann',
				subject: '',
				manager: undefined,
			}),
		).toBe(false);
		expect(wb.dirty).toBe(false);
		expect(wb.properties.title).toBe('Budget');
	});

	it('ignores dates and fields that are not text properties', () => {
		const wb = cleanWorkbook();
		expect(
			setDocumentProperties(wb, {
				created: '2030-01-01T00:00:00Z',
				modified: '2030-01-01T00:00:00Z',
				application: 'Other',
				custom: [{ name: 'x', value: 1 }],
			}),
		).toBe(false);
		expect(wb.properties.created).toBe('2024-01-01T10:00:00Z');
		expect(wb.properties.application).toBe('Microsoft Excel');
		expect(wb.properties.custom).toBeUndefined();
	});

	it('edits every listed text property, comments keep inner line breaks', () => {
		const wb = cleanWorkbook();
		const patch: Record<string, string> = {};
		for (const key of DOCUMENT_TEXT_PROPERTIES) {
			patch[key] = `${key} value`;
		}
		patch.description = '\nLine 1\nLine 2\n';
		expect(setDocumentProperties(wb, patch)).toBe(true);
		expect(wb.properties.manager).toBe('manager value');
		expect(wb.properties.lastModifiedBy).toBe('lastModifiedBy value');
		expect(wb.properties.description).toBe('Line 1\nLine 2');
	});
});

describe('workbookStatistics', () => {
	it('counts sheets, non-empty cells and formulas', () => {
		const wb = cleanWorkbook();
		const [data, calc] = wb.sheets;
		data.cells.set(cellKey(0, 0), { value: 'Name' });
		data.cells.set(cellKey(1, 0), { value: 5 });
		data.cells.set(cellKey(2, 0), { value: null, style: 3 });
		data.cells.set(cellKey(3, 0), { value: '' });
		calc.cells.set(cellKey(0, 0), { value: 10, formula: 'Data!A2*2' });
		calc.cells.set(cellKey(1, 0), { value: null, formula: 'Data!A3' });
		calc.cells.set(cellKey(2, 0), { value: false });
		expect(workbookStatistics(wb)).toEqual({ sheets: 2, cells: 5, formulas: 2 });
	});

	it('is all zero but the sheet count for an empty workbook', () => {
		expect(workbookStatistics(createWorkbook())).toEqual({ sheets: 1, cells: 0, formulas: 0 });
	});
});
