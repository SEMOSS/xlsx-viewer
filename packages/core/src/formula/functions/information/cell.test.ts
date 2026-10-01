import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import type { CellInfoProvider } from './cell';
import { cellFormatCode, hasNegativeColor, hasParentheses } from './cell-format';
import { INFORMATION_FUNCTIONS } from './index';

const CELL = fn(INFORMATION_FUNCTIONS, 'CELL');
const INFO = fn(INFORMATION_FUNCTIONS, 'INFO');

function context(): MockContext {
	const ctx = new MockContext(
		{ A1: 5, B2: 'label', C3: ERRORS.NA },
		{ 'Other Sheet': { A1: 1 }, Data: {} },
	);
	return ctx.at('D4');
}

/** A context that also provides the optional CELL capabilities. */
class StyledContext extends MockContext implements CellInfoProvider {
	getNumberFormat(_sheet: number, row: number): string | undefined {
		return row === 0 ? '$#,##0.00_);[Red]($#,##0.00)' : row === 1 ? '(0.00)' : undefined;
	}
	getHorizontalAlignment(_sheet: number, row: number): string | undefined {
		return row === 1 ? 'center' : row === 2 ? 'right' : undefined;
	}
	isCellLocked(_sheet: number, row: number): boolean {
		return row !== 0;
	}
	getColumnWidth(_sheet: number, col: number): { width: number; isDefault: boolean } {
		return col === 0 ? { width: 12.71, isDefault: false } : { width: 8.43, isDefault: true };
	}
	getWorkbookName(): string {
		return 'Book1.xlsx';
	}
}

describe('CELL', () => {
	const ctx = context();
	const f = (...args: FunctionArg[]) => call(CELL, args, ctx);

	it('describes the position of the reference', () => {
		expect(f('address', ctx.ref('B2'))).toBe('$B$2');
		expect(f('address', ctx.ref('B2:C9'))).toBe('$B$2');
		expect(f('address', ctx.ref('A1', 1))).toBe("'Other Sheet'!$A$1");
		expect(f('address', ctx.ref('AA10', 2))).toBe('Data!$AA$10');
		expect(f('row', ctx.ref('B2:C9'))).toBe(2);
		expect(f('col', ctx.ref('AA10'))).toBe(27);
		expect(f('ROW')).toBe(4);
		expect(f('address')).toBe('$D$4');
	});

	it('describes the contents of the reference', () => {
		expect(f('contents', ctx.ref('A1'))).toBe(5);
		expect(f('contents', ctx.ref('B2'))).toBe('label');
		expect(f('contents', ctx.ref('Z9'))).toBe(0);
		expect(f('contents', ctx.ref('C3'))).toBe(ERRORS.NA);
		expect(f('type', ctx.ref('A1'))).toBe('v');
		expect(f('type', ctx.ref('B2'))).toBe('l');
		expect(f('type', ctx.ref('Z9'))).toBe('b');
		expect(f('type', ctx.ref('C3'))).toBe('v');
	});

	it('reports Excel defaults without formatting information', () => {
		expect(f('format', ctx.ref('A1'))).toBe('G');
		expect(f('color', ctx.ref('A1'))).toBe(0);
		expect(f('parentheses', ctx.ref('A1'))).toBe(0);
		expect(f('prefix', ctx.ref('B2'))).toBe("'");
		expect(f('prefix', ctx.ref('A1'))).toBe('');
		expect(f('protect', ctx.ref('A1'))).toBe(1);
		expect(f('width', ctx.ref('A1'))).toEqual(arr([8, true]));
		expect(f('filename', ctx.ref('A1'))).toBe('');
	});

	it('uses formatting information when the context provides it', () => {
		const styled = new StyledContext({ A1: -5, A2: 'centered', A3: 'right' }, { 'My Sheet': {} });
		const g = (...args: FunctionArg[]) => call(CELL, args, styled);
		expect(g('format', styled.ref('A1'))).toBe('C2-');
		expect(g('color', styled.ref('A1'))).toBe(1);
		expect(g('format', styled.ref('A2'))).toBe('F2()');
		expect(g('parentheses', styled.ref('A2'))).toBe(1);
		expect(g('prefix', styled.ref('A2'))).toBe('^');
		expect(g('prefix', styled.ref('A3'))).toBe('"');
		expect(g('protect', styled.ref('A1'))).toBe(0);
		expect(g('protect', styled.ref('A2'))).toBe(1);
		expect(g('width', styled.ref('A1'))).toEqual(arr([12, false]));
		expect(g('filename', styled.ref('A1'))).toBe('[Book1.xlsx]Sheet1');
		expect(g('address', styled.ref('B2', 1))).toBe("'[Book1.xlsx]My Sheet'!$B$2");
	});

	it('validates its arguments', () => {
		expect(f('bogus', ctx.ref('A1'))).toBe(ERRORS.VALUE);
		expect(f('row', 5)).toBe(ERRORS.VALUE);
		expect(f('row', ERRORS.REF)).toBe(ERRORS.REF);
		expect(f(ERRORS.NA, ctx.ref('A1'))).toBe(ERRORS.NA);
		expect(CELL.volatile).toBe(true);
	});
});

describe('cellFormatCode', () => {
	it.each([
		['General', 'G'],
		['0', 'F0'],
		['#,##0', ',0'],
		['0.00', 'F2'],
		['#,##0.00', ',2'],
		['$#,##0_);($#,##0)', 'C0'],
		['$#,##0_);[Red]($#,##0)', 'C0-'],
		['$#,##0.00_);($#,##0.00)', 'C2'],
		['$#,##0.00_);[Red]($#,##0.00)', 'C2-'],
		['"$"#,##0.00_);\\("$"#,##0.00\\)', 'C2'],
		['0%', 'P0'],
		['0.00%', 'P2'],
		['0.00E+00', 'S2'],
		['# ?/?', 'G'],
		['# ??/??', 'G'],
		['m/d/yy', 'D4'],
		['m/d/yy h:mm', 'D4'],
		['mm/dd/yy', 'D4'],
		['[$-409]m/d/yyyy', 'D4'],
		['d-mmm-yy', 'D1'],
		['dd-mmm-yy', 'D1'],
		['d-mmm', 'D2'],
		['mmm-yy', 'D3'],
		['mm/dd', 'D5'],
		['h:mm AM/PM', 'D7'],
		['h:mm:ss AM/PM', 'D6'],
		['h:mm', 'D9'],
		['h:mm:ss', 'D8'],
		['@', 'G'],
		['#,##0_);(#,##0)', ',0'],
		['(0)', 'F0()'],
	])('%s -> %s', (code, expected) => {
		expect(cellFormatCode(code)).toBe(expected);
	});

	it('detects negative colors and parentheses', () => {
		expect(hasNegativeColor('0;[Red]0')).toBe(true);
		expect(hasNegativeColor('[Blue]0')).toBe(true);
		expect(hasNegativeColor('0;0')).toBe(false);
		expect(hasParentheses('0_);(0)')).toBe(false);
		expect(hasParentheses('(0);(0)')).toBe(true);
		expect(cellFormatCode(undefined)).toBe('G');
	});
});

describe('INFO', () => {
	const ctx = context();
	const f = (...args: FunctionArg[]) => call(INFO, args, ctx);

	it('describes the environment', () => {
		expect(f('numfile')).toBe(3);
		expect(f('recalc')).toBe('Automatic');
		expect(f('release')).toBe('16.0');
		expect(f('SYSTEM')).toBe('pcdos');
		expect(f('osversion')).toMatch(/^Windows/);
		expect(f('origin')).toBe('$A:$A$1');
	});

	it('reports unavailable and invalid types', () => {
		expect(f('directory')).toBe(ERRORS.NA);
		expect(f('memavail')).toBe(ERRORS.NA);
		expect(f('nonsense')).toBe(ERRORS.VALUE);
		expect(f(ERRORS.REF)).toBe(ERRORS.REF);
		expect(INFO.volatile).toBe(true);
	});
});
