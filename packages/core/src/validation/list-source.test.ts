import { describe, expect, it } from 'vitest';
import { createFakeEvaluator } from '../conditional/test-utils/fake-evaluator';
import { rng, setCells } from '../conditional/test-utils/fixtures';
import { createWorkbook } from '../sheet/workbook';
import type { Cell, CellValue } from '../types/cell';
import type { DataValidation } from '../types/data-validation';
import { dateToSerial } from '../values/dates';
import { ERRORS } from '../values/errors';
import { getListEntries, getListItems, parseLiteralList } from './list-source';
import { findDefinedName, resolveReference } from './reference';
import type { DataValidationContext } from './types';
import { validateValue } from './validate';

function setup(
	main: Record<string, CellValue | Cell> = {},
	other: Record<string, CellValue | Cell> = {},
) {
	const workbook = createWorkbook({ sheetNames: ['Sheet1', 'My Data'] });
	setCells(workbook.sheets[0], main);
	setCells(workbook.sheets[1], other);
	const evaluate = createFakeEvaluator(workbook);
	const ctx = (
		row = 0,
		col = 0,
		extra: Partial<DataValidationContext> = {},
	): DataValidationContext => ({
		evaluate,
		sheetIndex: 0,
		row,
		col,
		date1904: false,
		workbook,
		...extra,
	});
	return { workbook, ctx, evaluate };
}

const list = (formula1: string, extra: Partial<DataValidation> = {}): DataValidation => ({
	ranges: [rng('A1:A10')],
	type: 'list',
	formula1,
	...extra,
});

describe('literal lists', () => {
	it('parses quoted comma-separated items', () => {
		expect(parseLiteralList('"Yes,No,Maybe"')).toEqual(['Yes', 'No', 'Maybe']);
		expect(parseLiteralList('"say ""hi"",bye"')).toEqual(['say "hi"', 'bye']);
		expect(parseLiteralList('"a, b"')).toEqual(['a', ' b']);
		expect(parseLiteralList('$A$1:$A$3')).toBeUndefined();
		expect(parseLiteralList('"a"&"b"')).toBeUndefined();
	});

	it('returns dropdown items and validates membership case-insensitively', () => {
		const { ctx } = setup();
		const rule = list('"Yes,No,,Maybe"');
		expect(getListItems(rule, ctx())).toEqual(['Yes', 'No', 'Maybe']);
		expect(validateValue('yes', rule, ctx()).valid).toBe(true);
		expect(validateValue('MAYBE', rule, ctx()).valid).toBe(true);
		expect(validateValue('Perhaps', rule, ctx()).valid).toBe(false);
		expect(validateValue('', rule, ctx()).valid).toBe(false);
	});

	it('matches numbers numerically', () => {
		const { ctx } = setup();
		const rule = list('"1,2.50,3"');
		expect(validateValue(2.5, rule, ctx()).valid).toBe(true);
		expect(validateValue('2.5', rule, ctx()).valid).toBe(true);
		expect(validateValue(4, rule, ctx()).valid).toBe(false);
		expect(validateValue(true, list('"TRUE,FALSE"'), ctx()).valid).toBe(true);
	});
});

describe('reference lists', () => {
	it('reads ranges on the same sheet, skipping blanks', () => {
		const { ctx } = setup({ C1: 'Red', C2: null, C3: 'Green', C4: '', C5: 'Blue' });
		expect(getListItems(list('$C$1:$C$5'), ctx())).toEqual(['Red', 'Green', 'Blue']);
		expect(validateValue('green', list('$C$1:$C$5'), ctx()).valid).toBe(true);
		expect(validateValue('Purple', list('$C$1:$C$5', { allowBlank: true }), ctx()).valid).toBe(
			false,
		);
	});

	it('reads other sheets with quoted names', () => {
		const { ctx } = setup({}, { A1: 'North', A2: 'South' });
		expect(getListItems(list("'My Data'!$A$1:$A$2"), ctx())).toEqual(['North', 'South']);
	});

	it('shifts relative list references with the validated cell', () => {
		const { ctx } = setup({ C1: 'a', C2: 'b', C3: 'c' });
		const rule = list('C1:C2');
		expect(getListItems(rule, ctx(0, 0))).toEqual(['a', 'b']);
		expect(getListItems(rule, ctx(1, 0))).toEqual(['b', 'c']);
	});

	it('clamps whole columns to the used range', () => {
		const { ctx } = setup({ D1: 'x', D3: 'y', F9: 1 });
		expect(getListItems(list('$D:$D'), ctx())).toEqual(['x', 'y']);
	});

	it('resolves defined names (sheet scope first) and applies the blank quirk', () => {
		const { workbook, ctx } = setup({ C1: 'g1', C2: 'g2', E1: 'l1', E3: 'l2' });
		workbook.definedNames.push({ name: 'Colors', formula: 'Sheet1!$C$1:$C$2' });
		workbook.definedNames.push({ name: 'colors', formula: 'Sheet1!$E$1:$E$3', localSheet: 0 });
		workbook.definedNames.push({ name: 'Nested', formula: 'Colors' });
		expect(findDefinedName(workbook, 'COLORS', 1)?.formula).toBe('Sheet1!$C$1:$C$2');
		expect(getListItems(list('Colors'), ctx())).toEqual(['l1', 'l2']);
		expect(resolveReference('Nested', workbook, 1)).toMatchObject({ sheetIndex: 0, named: true });
		expect(validateValue('anything', list('Colors', { allowBlank: true }), ctx()).valid).toBe(true);
		expect(validateValue('anything', list('Colors'), ctx()).valid).toBe(false);
		expect(validateValue('L2', list('Colors'), ctx()).valid).toBe(true);
	});

	it('shows formatted values and matches underlying values', () => {
		const { workbook, ctx } = setup();
		workbook.styles.cellXfs.push({
			numFmtId: 9,
			fontId: 0,
			fillId: 0,
			borderId: 0,
			xfId: 0,
			applyNumberFormat: true,
		});
		const serial = dateToSerial(2024, 3, 1);
		setCells(workbook.sheets[0], {
			C1: { value: 0.5, style: 1 },
			C2: 42,
			C3: true,
			C4: ERRORS.NA,
			C5: serial,
		});
		const entries = getListEntries(list('$C$1:$C$5'), ctx());
		expect(entries.map((e) => e.value)).toEqual([0.5, 42, true, ERRORS.NA, serial]);
		expect(entries.map((e) => e.text)).toEqual(['50%', '42', 'TRUE', '#N/A', String(serial)]);
		expect(validateValue(0.5, list('$C$1:$C$5'), ctx()).valid).toBe(true);
		expect(validateValue('50%', list('$C$1:$C$5'), ctx()).valid).toBe(true);
		expect(validateValue('42', list('$C$1:$C$5'), ctx()).valid).toBe(true);
		expect(validateValue('3/1/2024', list('$C$1:$C$5'), ctx()).valid).toBe(true);
		expect(validateValue(43, list('$C$1:$C$5'), ctx()).valid).toBe(false);
		const custom = getListItems(
			list('$C$1:$C$2'),
			ctx(0, 0, { formatValue: (v) => `<${String(v)}>` }),
		);
		expect(custom).toEqual(['<0.5>', '<42>']);
	});
});

describe('formula lists', () => {
	it('collects array results in order', () => {
		const { ctx } = setup();
		expect(getListItems(list('ARR("x",,"y",3)'), ctx())).toEqual(['x', 'y', '3']);
		expect(validateValue(3, list('ARR("x",3)'), ctx()).valid).toBe(true);
	});

	it('reads references returned by formulas and defined names the parser cannot resolve', () => {
		const { workbook, ctx } = setup({ C1: 'p', C2: 'q' });
		workbook.definedNames.push({ name: 'Dyn', formula: 'IF(TRUE,$C$1:$C$2)' });
		expect(getListItems(list('IF(TRUE,$C$1:$C$2)'), ctx())).toEqual(['p', 'q']);
		expect(getListItems(list('Dyn'), ctx())).toEqual(['p', 'q']);
	});

	it('handles scalar and error results, and no source', () => {
		const { ctx, evaluate } = setup();
		expect(getListItems(list('"a"&"b"'), ctx())).toEqual(['ab']);
		expect(getListItems(list('NA()'), ctx())).toEqual([]);
		expect(getListItems(list('Missing!$A$1:$A$2'), ctx())).toEqual([]);
		expect(getListItems(list('ARR("a",NA())'), ctx())).toEqual(['a', '#N/A']);
		expect(getListItems({ ranges: [rng('A1')], type: 'list' }, ctx())).toEqual([]);
		expect(getListItems({ ranges: [rng('A1')], type: 'whole', formula1: '"a,b"' }, ctx())).toEqual(
			[],
		);
		expect(validateValue('z', { ranges: [rng('A1')], type: 'list' }, ctx()).valid).toBe(true);
		const bare: DataValidationContext = {
			evaluate,
			sheetIndex: 0,
			row: 0,
			col: 0,
			date1904: false,
		};
		expect(getListItems(list('$C$1:$C$2'), bare)).toEqual([]);
		expect(getListItems(list('ARR("k")'), bare)).toEqual(['k']);
	});
});
