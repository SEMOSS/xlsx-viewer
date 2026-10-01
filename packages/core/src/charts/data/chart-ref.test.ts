import { describe, expect, it } from 'vitest';
import { makeWorkbook } from '../__tests__/helpers';
import { formatChartRef, resolveChartRef } from './chart-ref';

const wb = makeWorkbook({ Sheet1: [[1]], 'My Sheet': [[2]], "Bob's": [[3]] });
wb.definedNames.push(
	{ name: 'Sales', formula: 'Sheet1!$B$2:$B$6' },
	{ name: 'Local', formula: "'My Sheet'!$A$1:$A$3", localSheet: 1 },
	{ name: 'Alias', formula: 'Sales' },
	{ name: 'Dynamic', formula: 'OFFSET(Sheet1!$A$1,0,0,5,1)' },
	{ name: 'Loop', formula: 'Loop' },
);

describe('resolveChartRef', () => {
	it('resolves plain and absolute references', () => {
		expect(resolveChartRef('Sheet1!$B$2:$B$6', wb, 0)).toEqual([
			{ sheetIndex: 0, range: { top: 1, left: 1, bottom: 5, right: 1 } },
		]);
		expect(resolveChartRef('=Sheet1!A1', wb, 0)).toEqual([
			{ sheetIndex: 0, range: { top: 0, left: 0, bottom: 0, right: 0 } },
		]);
	});

	it('resolves quoted sheet names, including escaped apostrophes', () => {
		expect(resolveChartRef("'My Sheet'!$A$2:$A$6", wb, 0)?.[0].sheetIndex).toBe(1);
		expect(resolveChartRef("'Bob''s'!$C$1", wb, 0)).toEqual([
			{ sheetIndex: 2, range: { top: 0, left: 2, bottom: 0, right: 2 } },
		]);
	});

	it('resolves multi-area unions in order', () => {
		expect(resolveChartRef("(Sheet1!$A$1,Sheet1!$A$3:$A$4,'My Sheet'!$B$2)", wb, 0)).toEqual([
			{ sheetIndex: 0, range: { top: 0, left: 0, bottom: 0, right: 0 } },
			{ sheetIndex: 0, range: { top: 2, left: 0, bottom: 3, right: 0 } },
			{ sheetIndex: 1, range: { top: 1, left: 1, bottom: 1, right: 1 } },
		]);
	});

	it('resolves workbook and sheet scoped defined names', () => {
		expect(resolveChartRef('Sales', wb, 0)?.[0].range).toEqual({
			top: 1,
			left: 1,
			bottom: 5,
			right: 1,
		});
		expect(resolveChartRef('Alias', wb, 0)?.[0].range).toEqual({
			top: 1,
			left: 1,
			bottom: 5,
			right: 1,
		});
		expect(resolveChartRef('[0]!Sales', wb, 0)?.[0].sheetIndex).toBe(0);
		expect(resolveChartRef("'My Sheet'!Local", wb, 0)?.[0]).toEqual({
			sheetIndex: 1,
			range: { top: 0, left: 0, bottom: 2, right: 0 },
		});
		expect(resolveChartRef('Local', wb, 1)?.[0].sheetIndex).toBe(1);
		expect(resolveChartRef('Local', wb, 0)).toBeNull();
	});

	it('returns null for references it cannot read', () => {
		expect(resolveChartRef('[1]Sheet1!$A$1', wb, 0)).toBeNull();
		expect(resolveChartRef('Missing!$A$1', wb, 0)).toBeNull();
		expect(resolveChartRef('Sheet1:Sheet2!$A$1', wb, 0)).toBeNull();
		expect(resolveChartRef('Dynamic', wb, 0)).toBeNull();
		expect(resolveChartRef('Loop', wb, 0)).toBeNull();
		expect(resolveChartRef('Unknown', wb, 0)).toBeNull();
		expect(resolveChartRef('(Sheet1!$A$1', wb, 0)).toBeNull();
		expect(resolveChartRef('', wb, 0)).toBeNull();
	});
});

describe('formatChartRef', () => {
	it('quotes sheet names when needed and uses absolute references', () => {
		expect(formatChartRef('Sheet1', { top: 1, left: 1, bottom: 5, right: 1 })).toBe(
			'Sheet1!$B$2:$B$6',
		);
		expect(formatChartRef('My Sheet', { top: 0, left: 0, bottom: 0, right: 0 })).toBe(
			"'My Sheet'!$A$1",
		);
		expect(formatChartRef("Bob's", { top: 0, left: 0, bottom: 1, right: 1 })).toBe(
			"'Bob''s'!$A$1:$B$2",
		);
	});
});
