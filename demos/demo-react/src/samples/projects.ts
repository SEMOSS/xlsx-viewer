import type { RangeAddress, StylePatch, Workbook } from 'xlsx-viewer-core';
import {
	addConditionalFormat,
	applyStylePatch,
	createCellIsRule,
	createIconSetRule,
	createTable,
	createWorkbook,
	EditTransaction,
	groupLines,
	parseRange,
	setCellInput,
	setColumnWidth,
	setFreezePanes,
} from 'xlsx-viewer-core';

/** A project tracker: a styled table, status lists, due-date rules, progress icons and grouped phases. */
export function buildProjectWorkbook(): Workbook {
	const wb = createWorkbook({ sheetNames: ['Tracker'], title: 'Project Tracker' });
	const tx = new EditTransaction(wb, 'sample');
	const r = (a1: string): RangeAddress => parseRange(a1)!;
	const style = (a1: string, patch: StylePatch) => applyStylePatch(tx, wb, 0, [r(a1)], patch);
	// Dates relative to today, so the "Days left" column always looks current.
	const day = (offset: number) => {
		const d = new Date();
		d.setDate(d.getDate() + offset);
		return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
	};
	const rows: string[][] = [
		['Task', 'Owner', 'Status', 'Start', 'Due', 'Progress', 'Days left'],
		['Phase 1: Discovery', '', '', '', '', '', ''],
		['Stakeholder interviews', 'Ana', 'Done', day(-60), day(-48), '1', ''],
		['Requirements', 'Ben', 'Done', day(-52), day(-30), '1', ''],
		['Phase 2: Build', '', '', '', '', '', ''],
		['Data model', 'Chen', 'In progress', day(-28), day(4), '0.7', ''],
		['User interface', 'Dana', 'In progress', day(-14), day(21), '0.45', ''],
		['Integrations', 'Ben', 'Not started', day(5), day(40), '0', ''],
		['Phase 3: Launch', '', '', '', '', '', ''],
		['Testing', 'Eli', 'Not started', day(35), day(62), '0', ''],
		['Training', 'Ana', 'Blocked', day(-3), day(12), '0.1', ''],
		['Go live', 'Chen', 'Not started', day(70), day(74), '0', ''],
	];
	rows.forEach((row, i) => row.forEach((v, j) => v !== '' && setCellInput(tx, wb, 0, i, j, v)));
	for (let i = 2; i <= rows.length; i++) {
		if (rows[i - 1][1] !== '') {
			setCellInput(tx, wb, 0, i - 1, 6, `=IF(C${i}="Done","",E${i}-TODAY())`);
		}
	}
	createTable(tx, wb, 0, r('A1:G12'), { styleName: 'TableStyleMedium9' });
	style('D2:E12', { numFmt: 'd mmm yyyy' });
	style('F2:F12', { numFmt: '0%' });
	style('G2:G12', { numFmt: '0;[Red]-0' });
	for (const a1 of ['A2:G2', 'A5:G5', 'A9:G9']) {
		style(a1, {
			font: { bold: true },
			fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFDDEBF7' } },
		});
	}
	// Tasks are grouped under their phase rows (summary rows above the detail).
	wb.sheets[0].summaryBelow = false;
	groupLines(tx, wb, 0, 'rows', 2, 3);
	groupLines(tx, wb, 0, 'rows', 5, 7);
	groupLines(tx, wb, 0, 'rows', 9, 11);

	const sheet = wb.sheets[0];
	sheet.dataValidations.push({
		ranges: [r('C3:C12')],
		type: 'list',
		formula1: '"Not started,In progress,Blocked,Done"',
		inCellDropdown: true,
		showInputMessage: true,
		promptTitle: 'Status',
		prompt: 'Choose a status from the list.',
		showErrorMessage: true,
		errorStyle: 'stop',
		error: 'Use one of the listed statuses.',
	});
	addConditionalFormat(
		sheet,
		[r('C3:C12')],
		createCellIsRule('equal', ['"Done"'], 'greenFillDarkGreenText'),
		wb.styles,
	);
	addConditionalFormat(
		sheet,
		[r('C3:C12')],
		createCellIsRule('equal', ['"Blocked"'], 'lightRedFillDarkRedText'),
		wb.styles,
	);
	addConditionalFormat(
		sheet,
		[r('C3:C12')],
		createCellIsRule('equal', ['"In progress"'], 'yellowFillDarkYellowText'),
		wb.styles,
	);
	addConditionalFormat(sheet, [r('F3:F12')], createIconSetRule('3Symbols'), wb.styles);
	sheet.notes.push({
		ref: { row: 10, col: 2 },
		author: 'Ana',
		text: 'Waiting for the training rooms to be booked.',
	});
	sheet.commentsDirty = true;

	setColumnWidth(tx, wb, 0, 0, 0, 26);
	setColumnWidth(tx, wb, 0, 1, 1, 10);
	setColumnWidth(tx, wb, 0, 2, 2, 14);
	setColumnWidth(tx, wb, 0, 3, 4, 13);
	setColumnWidth(tx, wb, 0, 5, 6, 11);
	setFreezePanes(tx, wb, 0, { row: 1, col: 1 });
	tx.commit();
	return wb;
}
