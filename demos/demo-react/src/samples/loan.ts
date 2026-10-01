import type { RangeAddress, StylePatch, Workbook } from 'xlsx-viewer-core';
import {
	addConditionalFormat,
	applyBorderPreset,
	applyStylePatch,
	createChartFromRange,
	createDataBarRule,
	createWorkbook,
	EditTransaction,
	mergeRange,
	parseRange,
	setCellInput,
	setColumnWidth,
	setFreezePanes,
} from 'xlsx-viewer-core';

const MONEY = '"$"#,##0.00';

/** A loan calculator: inputs with validation, PMT / IPMT / PPMT and an amortization schedule with a chart. */
export function buildLoanWorkbook(): Workbook {
	const wb = createWorkbook({ sheetNames: ['Loan'], title: 'Loan Calculator' });
	const tx = new EditTransaction(wb, 'sample');
	const r = (a1: string): RangeAddress => parseRange(a1)!;
	const put = (a1: string, text: string) => {
		const ref = r(a1);
		setCellInput(tx, wb, 0, ref.top, ref.left, text);
	};
	const style = (a1: string, patch: StylePatch) => applyStylePatch(tx, wb, 0, [r(a1)], patch);
	const blue = { rgb: 'FF1F4E79' };

	put('A1', 'Loan Calculator');
	mergeRange(tx, wb, 0, r('A1:F1'));
	style('A1', {
		font: { size: 18, bold: true, color: { rgb: 'FFFFFFFF' } },
		fill: { type: 'pattern', patternType: 'solid', fgColor: blue },
		alignment: { vertical: 'center' },
	});

	const inputs: Array<[string, string, string]> = [
		['Loan amount', '250000', MONEY],
		['Annual interest rate', '5.25%', '0.00%'],
		['Term (years)', '30', '0'],
		['Payments per year', '12', '0'],
	];
	inputs.forEach(([label, value, fmt], i) => {
		put(`A${i + 3}`, label);
		put(`B${i + 3}`, value);
		style(`B${i + 3}`, {
			numFmt: fmt,
			fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFFFF2CC' } },
			protection: { locked: false },
		});
	});
	applyBorderPreset(tx, wb, 0, r('A3:B6'), 'all', { style: 'thin', color: { rgb: 'FFBFBFBF' } });

	put('D3', 'Payment');
	put('E3', '=-PMT(B4/B6,B5*B6,B3)');
	put('D4', 'Total paid');
	put('E4', '=E3*B5*B6');
	put('D5', 'Total interest');
	put('E5', '=E4-B3');
	put('D6', 'Payoff date');
	put('E6', '=EDATE(TODAY(),B5*12)');
	style('E3:E5', { numFmt: MONEY, font: { bold: true } });
	style('E6', { numFmt: 'mmm yyyy', font: { bold: true } });
	style('D3:D6', { font: { color: { rgb: 'FF595959' } } });

	// Amortization schedule: the first five years, one row per year.
	const header = ['Year', 'Payments', 'Interest', 'Principal', 'Balance'];
	header.forEach((h, i) => setCellInput(tx, wb, 0, 8, i, h));
	style('A9:E9', {
		font: { bold: true, color: { rgb: 'FFFFFFFF' } },
		fill: { type: 'pattern', patternType: 'solid', fgColor: blue },
		alignment: { horizontal: 'center' },
	});
	for (let y = 1; y <= 30; y++) {
		const row = 9 + y;
		put(`A${row}`, String(y));
		put(`B${row}`, '=$E$3*$B$6');
		put(`C${row}`, `=-CUMIPMT($B$4/$B$6,$B$5*$B$6,$B$3,${(y - 1) * 12 + 1},${y * 12},0)`);
		put(`D${row}`, `=-CUMPRINC($B$4/$B$6,$B$5*$B$6,$B$3,${(y - 1) * 12 + 1},${y * 12},0)`);
		put(`E${row}`, y === 1 ? `=$B$3-D${row}` : `=E${row - 1}-D${row}`);
	}
	style('B10:E39', { numFmt: MONEY });
	addConditionalFormat(
		wb.sheets[0],
		[r('C10:C39')],
		createDataBarRule({ rgb: 'FFF4B183' }, true),
		wb.styles,
	);
	addConditionalFormat(
		wb.sheets[0],
		[r('D10:D39')],
		createDataBarRule({ rgb: 'FF9DC3E6' }, true),
		wb.styles,
	);

	wb.sheets[0].dataValidations.push(
		{
			ranges: [r('B3')],
			type: 'decimal',
			operator: 'between',
			formula1: '1000',
			formula2: '10000000',
			showInputMessage: true,
			promptTitle: 'Loan amount',
			prompt: 'Between $1,000 and $10,000,000.',
			showErrorMessage: true,
			errorStyle: 'stop',
			error: 'Enter an amount between $1,000 and $10,000,000.',
		},
		{
			ranges: [r('B5')],
			type: 'list',
			formula1: '"10,15,20,25,30"',
			inCellDropdown: true,
			showErrorMessage: true,
			errorStyle: 'stop',
			error: 'Pick a term from the list.',
		},
	);
	wb.sheets[0].notes.push({
		ref: { row: 2, col: 1 },
		author: 'Demo User',
		text: 'Yellow cells are inputs. Everything else recalculates.',
	});
	wb.sheets[0].commentsDirty = true;

	// One series (Balance); its categories default to 1..30, the loan years.
	const chart = createChartFromRange(wb, 0, r('E10:E39'), 'line');
	chart.plots[0].series[0] = { ...chart.plots[0].series[0], name: { text: 'Balance' } };
	chart.title = { text: 'Remaining balance' };
	chart.legend = undefined;
	wb.sheets[0].drawings.push({
		kind: 'chart',
		id: 'chart-balance',
		name: 'Balance chart',
		chart,
		anchor: {
			type: 'twoCell',
			from: { row: 8, col: 6, rowOff: 0, colOff: 0 },
			to: { row: 26, col: 13, rowOff: 0, colOff: 0 },
		},
		dirty: true,
	});
	wb.sheets[0].drawingsDirty = true;

	setColumnWidth(tx, wb, 0, 0, 0, 22);
	setColumnWidth(tx, wb, 0, 1, 1, 14);
	setColumnWidth(tx, wb, 0, 2, 2, 14);
	setColumnWidth(tx, wb, 0, 3, 3, 16);
	setColumnWidth(tx, wb, 0, 4, 4, 15);
	setFreezePanes(tx, wb, 0, { row: 9, col: 0 });
	tx.commit();
	return wb;
}
