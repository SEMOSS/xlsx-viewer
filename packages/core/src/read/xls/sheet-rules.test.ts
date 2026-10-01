import { describe, expect, it } from 'vitest';

import { RT } from './biff-opcodes';
import { readXlsWorkbook } from './index';
import { buildXls, definedName, externSheet, supbookSelf } from './test-utils/biff-builder';
import { record } from './test-utils/bytes';
import { ptg } from './test-utils/ptg-builder';
import { cf, condFmt, dv, dval } from './test-utils/sheet-builder';

function readSheet(records: number[][], globals: number[][] = []) {
	return readXlsWorkbook(buildXls(globals, [{ name: 'Sheet1', records }]));
}

describe('data validation', () => {
	it('reads explicit lists, ranges, operators and messages', () => {
		const wb = readSheet(
			[
				dval(4),
				dv({
					type: 3,
					list: true,
					allowBlank: true,
					showInput: true,
					showError: true,
					promptTitle: 'Pick',
					prompt: 'Choose one',
					errorTitle: 'Oops',
					error: 'Not in list',
					f1: ptg.str('Yes\u0000No\u0000"Maybe"'),
					ranges: [[1, 9, 0, 0]],
				}),
				dv({
					type: 1,
					op: 0,
					errStyle: 1,
					f1: ptg.int(1),
					f2: ptg.int(10),
					ranges: [
						[0, 0, 2, 3],
						[5, 5, 5, 5],
					],
				}),
				dv({ type: 3, suppressDropdown: true, f1: ptg.name(1), ranges: [[0, 0, 4, 4]] }),
				dv({ type: 7, f1: [...ptg.refN(0, -1), ...ptg.int(0), ...ptg.gt], ranges: [[3, 6, 2, 2]] }),
				dv({
					type: 4,
					op: 6,
					f1: [...ptg.int(2000), ...ptg.int(1), ...ptg.int(1), ...ptg.func(65)],
					ranges: [[0, 0, 6, 6]],
				}),
			],
			[
				definedName({ name: 'Choices', rgce: ptg.ref3d(0, 0, 0) }),
				supbookSelf(1),
				externSheet([[0, 0, 0]]),
			],
		);
		const [list, whole, named, custom, date] = wb.sheets[0].dataValidations;
		expect(list).toEqual({
			ranges: [{ top: 1, left: 0, bottom: 9, right: 0 }],
			type: 'list',
			formula1: '"Yes,No,""Maybe"""',
			allowBlank: true,
			inCellDropdown: true,
			showInputMessage: true,
			showErrorMessage: true,
			promptTitle: 'Pick',
			prompt: 'Choose one',
			errorTitle: 'Oops',
			error: 'Not in list',
		});
		expect(whole).toEqual({
			ranges: [
				{ top: 0, left: 2, bottom: 0, right: 3 },
				{ top: 5, left: 5, bottom: 5, right: 5 },
			],
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '10',
			errorStyle: 'warning',
		});
		expect(named.formula1).toBe('Choices');
		expect(named.inCellDropdown).toBe(false);
		expect(custom.formula1).toBe('B4>0');
		expect(date.operator).toBe('greaterThanOrEqual');
		expect(date.formula1).toBe('DATE(2000,1,1)');
	});
});

describe('conditional formatting', () => {
	it('reads cell-value and expression rules with their formats', () => {
		const wb = readSheet([
			condFmt(2, [
				[1, 10, 1, 2],
				[12, 12, 1, 1],
			]),
			cf({ kind: 1, op: 1, f1: ptg.int(5), f2: ptg.int(10), fontColor: 10, bold: true }),
			cf({
				kind: 2,
				f1: [...ptg.refN(0, -1, { colAbs: false }), ...ptg.str('x'), ...ptg.eq],
				fill: 13,
			}),
			condFmt(1, [[0, 0, 5, 5]]),
			cf({
				kind: 1,
				op: 5,
				f1: ptg.ref(0, 0, { rowAbs: true, colAbs: true }),
				border: 10,
				italic: true,
			}),
		]);
		const formats = wb.sheets[0].conditionalFormats;
		expect(formats).toHaveLength(2);
		expect(formats[0].ranges).toEqual([
			{ top: 1, left: 1, bottom: 10, right: 2 },
			{ top: 12, left: 1, bottom: 12, right: 1 },
		]);
		expect(formats[0].rules).toEqual([
			{
				type: 'cellIs',
				priority: 1,
				stopIfTrue: true,
				dxfId: 0,
				formulas: ['5', '10'],
				operator: 'between',
			},
			{ type: 'expression', priority: 2, stopIfTrue: true, dxfId: 1, formulas: ['A2="x"'] },
		]);
		expect(formats[1].rules[0]).toMatchObject({
			type: 'cellIs',
			operator: 'greaterThan',
			formulas: ['$A$1'],
			dxfId: 2,
		});
		const [red, filled, bordered] = wb.styles.dxfs;
		expect(red).toEqual({ font: { bold: true, color: { indexed: 10 } } });
		expect(filled).toEqual({
			fill: { type: 'pattern', patternType: 'solid', bgColor: { indexed: 13 } },
		});
		expect(bordered.font).toEqual({ italic: true, bold: false });
		expect(bordered.border).toEqual({
			left: { style: 'thin', color: { indexed: 10 } },
			right: { style: 'thin', color: { indexed: 10 } },
			top: { style: 'thin', color: { indexed: 10 } },
			bottom: { style: 'thin', color: { indexed: 10 } },
		});
	});

	it('reports Excel 2007 rule records and ignores stray CF records', () => {
		const wb = readSheet([
			cf({ kind: 2, f1: ptg.bool(true) }),
			record(RT.CONDFMT12, [0, 0]),
			condFmt(1, [[0, 0, 0, 0]]),
			cf({ kind: 9, f1: ptg.bool(true) }),
		]);
		expect(wb.sheets[0].conditionalFormats).toEqual([]);
		const codes = wb.warnings.map((w) => w.code);
		expect(codes).toContain('conditional-format-skipped');
		expect(codes).toContain('conditional-format-unsupported');
	});
});
