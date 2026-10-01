import type { RangeAddress, StylePatch, Workbook } from 'xlsx-viewer-core';
import {
	addConditionalFormat,
	addSheet,
	applyBorderPreset,
	createChartFromRange,
	createColorScaleRule,
	createDataBarRule,
	createIconSetRule,
	createTable,
	applyStylePatch,
	createWorkbook,
	EditTransaction,
	mergeRange,
	parseRange,
	setCellInput,
	setColumnWidth,
	setFreezePanes,
	setRowHeight,
} from 'xlsx-viewer-core';

const ACCOUNTING = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"??_);_(@_)';

/** Build a showcase workbook in code (formatting, formulas, merges, frozen panes). */
export function buildSampleWorkbook(): Workbook {
	const wb = createWorkbook({
		sheetNames: ['Sales', 'Formats', 'Formulas'],
		title: 'Quarterly Sales',
	});
	const tx = new EditTransaction(wb, 'sample');
	const r = (a1: string): RangeAddress => parseRange(a1)!;
	const put = (sheet: number, a1: string, text: string) => {
		const ref = r(a1);
		setCellInput(tx, wb, sheet, ref.top, ref.left, text);
	};
	const style = (sheet: number, a1: string, patch: StylePatch) =>
		applyStylePatch(tx, wb, sheet, [r(a1)], patch);
	const green = { rgb: 'FF107C41' };
	const white = { rgb: 'FFFFFFFF' };

	// ── Sales ───────────────────────────────────────────────────────────
	put(0, 'A1', 'Quarterly Sales Report');
	mergeRange(tx, wb, 0, r('A1:G1'), { center: true });
	style(0, 'A1', {
		font: { size: 16, bold: true, color: white },
		fill: { type: 'pattern', patternType: 'solid', fgColor: green },
	});
	setRowHeight(tx, wb, 0, 0, 0, 27);
	put(0, 'A2', 'Fiscal year 2026, amounts in USD');
	style(0, 'A2', { font: { italic: true, color: { rgb: 'FF5B625E' } } });

	const headers = ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total', 'Share'];
	headers.forEach((h, i) => setCellInput(tx, wb, 0, 2, i, h));
	style(0, 'A3:G3', {
		font: { bold: true, color: { rgb: 'FF0B5A2E' } },
		fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFD9EAD3' } },
		alignment: { horizontal: 'center' },
	});
	applyBorderPreset(tx, wb, 0, r('A3:G3'), 'bottom', { style: 'medium', color: green });

	const data: Array<[string, number[]]> = [
		['North', [42150, 45890, 51200, 58740]],
		['South', [38900, 36120, 40980, 44610]],
		['East', [51230, 54870, 49910, 61020]],
		['West', [29840, 33410, 37760, 35290]],
		['Central', [33310, 31980, 34420, 39870]],
		['Online', [61200, 70450, 82310, 97640]],
	];
	data.forEach(([region, values], i) => {
		const row = 3 + i;
		setCellInput(tx, wb, 0, row, 0, region);
		values.forEach((v, j) => setCellInput(tx, wb, 0, row, 1 + j, String(v)));
		setCellInput(tx, wb, 0, row, 5, `=SUM(B${row + 1}:E${row + 1})`);
		setCellInput(tx, wb, 0, row, 6, `=F${row + 1}/$F$10`);
	});
	put(0, 'A10', 'Total');
	for (const col of ['B', 'C', 'D', 'E', 'F']) {
		put(0, `${col}10`, `=SUM(${col}4:${col}9)`);
	}
	put(0, 'G10', '=SUM(G4:G9)');
	style(0, 'B4:F10', { numFmt: ACCOUNTING });
	style(0, 'G4:G10', { numFmt: '0.0%' });
	style(0, 'A10:G10', { font: { bold: true } });
	applyBorderPreset(tx, wb, 0, r('A10:G10'), 'topDoubleBottom');
	style(0, 'A4:A9', { alignment: { indent: 1 } });

	put(0, 'A12', 'Average per region');
	put(0, 'B12', '=AVERAGE(F4:F9)');
	put(0, 'A13', 'Best quarter');
	put(0, 'B13', '=INDEX(B3:E3,MATCH(MAX(B10:E10),B10:E10,0))');
	put(0, 'A14', 'Growth Q1 to Q4');
	put(0, 'B14', '=E10/B10-1');
	put(0, 'A15', 'Regions above $200k');
	put(0, 'B15', '=COUNTIF(F4:F9,">200000")');
	style(0, 'B12', { numFmt: ACCOUNTING });
	style(0, 'B14', { numFmt: '0.0%' });
	style(0, 'A12:A15', { font: { bold: true } });
	put(0, 'D12', 'Tip: double-click a cell to edit, or type = to start a formula.');
	style(0, 'D12', { font: { italic: true, color: { rgb: 'FF7F7F7F' } } });

	setColumnWidth(tx, wb, 0, 0, 0, 22);
	setColumnWidth(tx, wb, 0, 1, 5, 13.5);
	setColumnWidth(tx, wb, 0, 6, 6, 9);

	// Sparklines: a quarterly trend per region.
	setCellInput(tx, wb, 0, 2, 7, 'Trend');
	style(0, 'H3', {
		font: { bold: true, color: { rgb: 'FF0B5A2E' } },
		fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFD9EAD3' } },
		alignment: { horizontal: 'center' },
	});
	applyBorderPreset(tx, wb, 0, r('H3'), 'bottom', { style: 'medium', color: green });
	setColumnWidth(tx, wb, 0, 7, 7, 14);
	wb.sheets[0].sparklineGroups.push({
		type: 'line',
		markers: false,
		high: true,
		low: true,
		highColor: { rgb: 'FF107C41' },
		lowColor: { rgb: 'FFD13438' },
		seriesColor: { rgb: 'FF2F6B8F' },
		lineWeight: 1.25,
		sparklines: data.map((_, i) => ({
			formula: `Sales!B${i + 4}:E${i + 4}`,
			location: { row: i + 3, col: 7 },
		})),
	});
	setFreezePanes(tx, wb, 0, { row: 3, col: 1 });

	// ── Formats ─────────────────────────────────────────────────────────
	const formats: Array<[string, string, string]> = [
		['General', '1234.5678', 'General'],
		['Number (2 dp)', '1234.5678', '#,##0.00'],
		['Currency', '-1234.5', '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)'],
		['Accounting', '1234.5', ACCOUNTING],
		['Percentage', '0.4567', '0.00%'],
		['Fraction', '1.375', '# ?/?'],
		['Scientific', '123456789', '0.00E+00'],
		['Short date', '3/14/2026', 'm/d/yyyy'],
		['Long date', '3/14/2026', 'dddd, mmmm d, yyyy'],
		['Time', '13:45:30', 'h:mm:ss AM/PM'],
		['Elapsed hours', '1.75', '[h]:mm'],
		['Text', "'00123", '@'],
		['Boolean', 'TRUE', 'General'],
		['Error', '#N/A', 'General'],
	];
	put(1, 'A1', 'Format');
	put(1, 'B1', 'Displayed');
	put(1, 'C1', 'Code');
	style(1, 'A1:C1', {
		font: { bold: true },
		fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFE7E6E6' } },
	});
	formats.forEach(([label, value, code], i) => {
		setCellInput(tx, wb, 1, i + 1, 0, label);
		setCellInput(tx, wb, 1, i + 1, 1, value);
		applyStylePatch(tx, wb, 1, [{ top: i + 1, left: 1, bottom: i + 1, right: 1 }], {
			numFmt: code,
		});
		setCellInput(tx, wb, 1, i + 1, 2, `'${code}`);
	});
	put(1, 'E1', 'Styles');
	style(1, 'E1', { font: { bold: true } });
	put(1, 'E2', 'Bold');
	style(1, 'E2', { font: { bold: true } });
	put(1, 'E3', 'Italic red');
	style(1, 'E3', { font: { italic: true, color: { rgb: 'FFC00000' } } });
	put(1, 'E4', 'Underlined');
	style(1, 'E4', { font: { underline: 'single' } });
	put(1, 'E5', 'Strikethrough');
	style(1, 'E5', { font: { strike: true } });
	put(1, 'E6', 'Fill + center');
	style(1, 'E6', {
		fill: { type: 'pattern', patternType: 'solid', fgColor: { theme: 4, tint: 0.6 } },
		alignment: { horizontal: 'center' },
	});
	put(1, 'E7', 'This text wraps inside its cell because Wrap Text is on.');
	style(1, 'E7', { alignment: { wrapText: true, vertical: 'top' } });
	setRowHeight(tx, wb, 1, 6, 6, 45);
	put(1, 'E8', 'Rotated');
	style(1, 'E8', { alignment: { textRotation: 45, horizontal: 'center' } });
	setRowHeight(tx, wb, 1, 7, 7, 40);
	put(1, 'E9', 'Pattern fill');
	style(1, 'E9', {
		fill: {
			type: 'pattern',
			patternType: 'lightUp',
			fgColor: { rgb: 'FF107C41' },
			bgColor: { rgb: 'FFFFFFFF' },
		},
	});
	put(1, 'E10', 'All borders');
	applyBorderPreset(tx, wb, 1, r('E10'), 'all', { style: 'medium', color: { rgb: 'FF4472C4' } });
	put(1, 'E11', 'This long label overflows into the empty cells on its right');
	const link = { font: { underline: 'single' as const, color: { theme: 10 } } };
	put(1, 'E12', 'Excel help (web link)');
	style(1, 'E12', link);
	wb.sheets[1].hyperlinks.push({ ref: r('E12'), target: 'https://support.microsoft.com/excel' });
	put(1, 'E13', 'Go to the Formulas sheet');
	style(1, 'E13', link);
	wb.sheets[1].hyperlinks.push({
		ref: r('E13'),
		location: 'Formulas!A1',
		tooltip: 'Formulas!A1 (internal link)',
	});
	setColumnWidth(tx, wb, 1, 0, 0, 16);
	setColumnWidth(tx, wb, 1, 1, 1, 26);
	setColumnWidth(tx, wb, 1, 2, 2, 34);
	setColumnWidth(tx, wb, 1, 4, 4, 18);

	// ── Formulas ────────────────────────────────────────────────────────
	const formulas: Array<[string, string]> = [
		['=SUM(1,2,3)', 'Basic arithmetic'],
		['=ROUND(PI(),4)', 'Math'],
		['=TEXT(DATE(2026,3,14),"dddd")', 'Dates and TEXT'],
		['=UPPER("excel")&" "&LEN("viewer")', 'Text functions'],
		['=IF(Sales!F10>1000000,"Over $1M","Under $1M")', 'Cross-sheet reference'],
		['=VLOOKUP("East",Sales!A4:F9,6,FALSE)', 'Lookup'],
		['=XLOOKUP("Online",Sales!A4:A9,Sales!E4:E9)', 'XLOOKUP'],
		['=SUMPRODUCT(Sales!B4:B9,Sales!C4:C9)/1E6', 'SUMPRODUCT'],
		['=AVERAGEIFS(Sales!F4:F9,Sales!F4:F9,">200000")', 'Conditional average'],
		['=LET(x,5,y,7,x*y)', 'LET'],
	];
	put(2, 'A1', 'Formula');
	put(2, 'B1', 'Result');
	put(2, 'C1', 'What it shows');
	style(2, 'A1:C1', {
		font: { bold: true },
		fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFE7E6E6' } },
	});
	formulas.forEach(([f, label], i) => {
		setCellInput(tx, wb, 2, i + 1, 0, `'${f}`);
		setCellInput(tx, wb, 2, i + 1, 1, f);
		setCellInput(tx, wb, 2, i + 1, 2, label);
	});
	put(2, 'E1', 'Dynamic array (spills):');
	style(2, 'E1', { font: { bold: true } });
	put(2, 'E2', '=SEQUENCE(5,1,10,10)');
	put(2, 'F2', '=SORT(Sales!A4:A9)');
	setColumnWidth(tx, wb, 2, 0, 0, 44);
	setColumnWidth(tx, wb, 2, 1, 1, 16);
	setColumnWidth(tx, wb, 2, 2, 2, 24);
	setColumnWidth(tx, wb, 2, 4, 5, 14);

	// ── Showcase: conditional formatting, a chart, a note and a table ─────
	const sales = wb.sheets[0];
	addConditionalFormat(
		sales,
		[r('F4:F9')],
		createDataBarRule({ rgb: 'FF638EC6' }, true),
		wb.styles,
	);
	addConditionalFormat(sales, [r('B4:E9')], createColorScaleRule(2), wb.styles);
	addConditionalFormat(sales, [r('G4:G9')], createIconSetRule('3Arrows'), wb.styles);
	sales.notes.push({
		ref: { row: 8, col: 0 },
		author: 'Demo User',
		text: 'Online sales include the marketplace channel from Q2 onwards.',
	});
	sales.commentsDirty = true;
	const chart = createChartFromRange(wb, 0, r('A3:E9'), 'column');
	chart.title = { text: 'Sales by Region and Quarter' };
	sales.drawings.push({
		kind: 'chart',
		id: 'chart-sales',
		name: 'Sales chart',
		chart,
		anchor: {
			type: 'twoCell',
			from: { row: 16, col: 0, rowOff: 0, colOff: 0 },
			to: { row: 34, col: 6, rowOff: 0, colOff: 0 },
		},
		dirty: true,
	});
	sales.drawingsDirty = true;

	const tableSheet = addSheet(tx, wb, 3, 'Table');
	const rows: string[][] = [
		['Product', 'Category', 'Units', 'Price', 'Revenue'],
		['Laptop Pro 14', 'Computers', '120', '1499', '=C2*D2'],
		['Wireless Mouse', 'Accessories', '640', '29.9', '=C3*D3'],
		['4K Monitor', 'Displays', '85', '399', '=C4*D4'],
		['USB-C Dock', 'Accessories', '210', '189', '=C5*D5'],
		['Tablet Air', 'Computers', '160', '649', '=C6*D6'],
		['Keyboard', 'Accessories', '420', '79', '=C7*D7'],
		['Ultrawide 34', 'Displays', '40', '699', '=C8*D8'],
	];
	rows.forEach((row, i) => row.forEach((v, j) => setCellInput(tx, wb, tableSheet, i, j, v)));
	applyStylePatch(tx, wb, tableSheet, [r('D2:E8')], { numFmt: '"$"#,##0.00' });
	createTable(tx, wb, tableSheet, r('A1:E8'), { styleName: 'TableStyleMedium2' });
	setColumnWidth(tx, wb, tableSheet, 0, 1, 18);
	setColumnWidth(tx, wb, tableSheet, 2, 4, 13);
	// Data validation: categories from a list, units as whole numbers, each with an input message.
	wb.sheets[tableSheet].dataValidations.push(
		{
			ranges: [r('B2:B8')],
			type: 'list',
			formula1: '"Computers,Accessories,Displays"',
			inCellDropdown: true,
			showInputMessage: true,
			promptTitle: 'Category',
			prompt: 'Pick a category from the list.',
			showErrorMessage: true,
			errorStyle: 'stop',
			errorTitle: 'Unknown category',
			error: 'Choose Computers, Accessories or Displays.',
		},
		{
			ranges: [r('C2:C8')],
			type: 'whole',
			operator: 'between',
			formula1: '0',
			formula2: '10000',
			showInputMessage: true,
			promptTitle: 'Units',
			prompt: 'A whole number from 0 to 10,000.',
			showErrorMessage: true,
			errorStyle: 'warning',
			error: 'That is outside the usual range.',
		},
	);

	tx.commit();
	wb.view.activeTab = 0;
	return wb;
}
