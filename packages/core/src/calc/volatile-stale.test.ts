import { describe, expect, it } from 'vitest';

import { ERRORS } from '../values/errors';
import { CalcEngine } from './engine';
import { calculated, cellAt, edit, makeWorkbook, refs, valueAt } from './test-utils/workbook';

describe('volatile functions', () => {
	it('recalculate NOW / TODAY / RAND on every recalculation', () => {
		let clock = new Date(2024, 0, 1, 12, 0, 0);
		let n = 0;
		const { wb, engine } = calculated(
			{ Sheet1: { A1: '=NOW()', A2: '=TODAY()', A3: '=RAND()', B3: '=A3*2', C1: 1, C2: '=C1+1' } },
			{ now: () => clock, random: () => (n += 0.125) },
		);
		expect(valueAt(wb, 'A1')).toBe(45292.5);
		expect(valueAt(wb, 'A2')).toBe(45292);
		expect(valueAt(wb, 'B3')).toBe(0.5);
		clock = new Date(2024, 0, 2, 6, 0, 0);
		const r = edit(engine, 'C1', 5);
		expect(valueAt(wb, 'A1')).toBe(45293.25);
		expect(valueAt(wb, 'A2')).toBe(45293);
		expect(valueAt(wb, 'A3')).toBe(0.375);
		expect(refs(wb, r.changed)).toEqual([
			'Sheet1!A1',
			'Sheet1!A2',
			'Sheet1!C2',
			'Sheet1!A3',
			'Sheet1!B3',
		]);
		const v = engine.recalculateVolatile();
		expect(refs(wb, v.changed)).toEqual(['Sheet1!A3', 'Sheet1!B3']);
		expect(valueAt(wb, 'B3')).toBe(1);
	});

	it('treat OFFSET, INDIRECT and names using them as volatile', () => {
		const { wb, engine } = calculated(
			{ Sheet1: { A1: 1, A2: 2, B1: '=SUM(OFFSET(A1,0,0,2))', B2: '=INDIRECT("A2")', B3: '=Dyn' } },
			{},
			(w) => w.definedNames.push({ name: 'Dyn', formula: 'OFFSET(Sheet1!$A$1,1,0)' }),
		);
		expect(valueAt(wb, 'B3')).toBe(2);
		const r = engine.recalculateVolatile();
		expect(r.changed).toEqual([]);
		edit(engine, 'A2', 7);
		expect(valueAt(wb, 'B1')).toBe(8);
		expect(valueAt(wb, 'B2')).toBe(7);
		expect(valueAt(wb, 'B3')).toBe(7);
	});
});

describe('unsupported functions', () => {
	it('keep cached values from the file, flag the cell stale and warn once', () => {
		const wb = makeWorkbook({
			Sheet1: {
				A1: 2,
				B1: { f: 'FUTUREFN(A1)', v: 99 },
				B2: { f: 'FUTUREFN(A1)+1', v: 100 },
				B3: { f: 'OTHERFN()', v: 'cached' },
				C1: { f: 'B1+1', v: 0 },
				D1: { f: 'FUTUREFN(1)' },
			},
		});
		const engine = new CalcEngine(wb);
		engine.rebuild();
		expect(cellAt(wb, 'B1')).toMatchObject({ value: 99, stale: true });
		expect(cellAt(wb, 'B3')).toMatchObject({ value: 'cached', stale: true });
		expect(valueAt(wb, 'D1')).toEqual(ERRORS.NAME);
		expect(wb.warnings.filter((w) => w.code === 'unsupported-function')).toHaveLength(2);
		expect(wb.warnings[0]).toMatchObject({ sheet: 'Sheet1', cell: 'B1' });
		expect(wb.warnings[0].message).toContain('FUTUREFN');

		engine.recalculateAll();
		expect(valueAt(wb, 'B1')).toBe(99);
		expect(valueAt(wb, 'C1')).toBe(100);
		edit(engine, 'A1', 5);
		expect(valueAt(wb, 'B1')).toBe(99);

		engine.rebuild();
		new CalcEngine(wb).rebuild();
		expect(wb.warnings.filter((w) => w.code === 'unsupported-function')).toHaveLength(2);
	});

	it('stay stale when their cells are rewritten by edits (sorting, re-touching)', () => {
		const wb = makeWorkbook({
			Sheet1: { A1: { f: 'FUTUREFN(1)', v: 5 }, B1: { f: 'A1*2', v: 0 } },
		});
		const engine = new CalcEngine(wb);
		engine.rebuild();
		engine.onCellsChanged([{ sheet: 0, row: 0, col: 0 }]);
		expect(cellAt(wb, 'A1')).toMatchObject({ value: 5, stale: true });
		expect(valueAt(wb, 'B1')).toBe(10);
		// A sort moves the stale cell to A2.
		const moved = cellAt(wb, 'A1')!;
		wb.sheets[0].cells.delete(0);
		wb.sheets[0].cells.set(16384, { ...moved, formula: 'FUTUREFN(1)' });
		engine.onCellsChanged([
			{ sheet: 0, row: 0, col: 0 },
			{ sheet: 0, row: 1, col: 0 },
		]);
		expect(cellAt(wb, 'A2')).toMatchObject({ value: 5, stale: true });
		expect(valueAt(wb, 'B1')).toBe(0);
	});

	it('give #NAME? for unknown functions typed by the user', () => {
		const { wb, engine } = calculated({ Sheet1: { A1: 1 } });
		edit(engine, 'B1', '=NOPE(A1)');
		expect(valueAt(wb, 'B1')).toEqual(ERRORS.NAME);
		expect(cellAt(wb, 'B1')?.stale).toBeUndefined();
	});

	it('keep cached values of external references and unparsable formulas', () => {
		const wb = makeWorkbook({
			Sheet1: {
				A1: { f: '[1]Prices!A1*2', v: 42 },
				A2: { f: 'SUM((', v: 7 },
				A3: { f: 'A1+A2', v: 0 },
			},
		});
		const engine = new CalcEngine(wb);
		engine.rebuild();
		engine.recalculateAll();
		expect(cellAt(wb, 'A1')).toMatchObject({ value: 42, stale: true });
		expect(cellAt(wb, 'A2')).toMatchObject({ value: 7, stale: true });
		expect(valueAt(wb, 'A3')).toBe(49);
		expect(wb.warnings.map((w) => w.code).sort()).toEqual([
			'external-reference',
			'formula-parse-error',
		]);
	});

	it('recompute stale cells once the function is available', () => {
		const wb = makeWorkbook({ Sheet1: { A1: 4, B1: { f: 'SQRT(A1)', v: 0 } } });
		wb.sheets[0].cells.set(1, { value: 0, formula: 'SQRT(A1)', stale: true });
		new CalcEngine(wb).rebuild();
		expect(cellAt(wb, 'B1')).toEqual({ value: 2, formula: 'SQRT(A1)' });
	});
});

describe('CELL() through the engine', () => {
	it('reports addresses, formats, protection and widths from the workbook', () => {
		const { wb } = calculated(
			{
				Sheet1: {
					A1: 1.5,
					B1: '=CELL("address",A1)',
					B2: '=CELL("format",A1)',
					B3: '=CELL("protect",A1)',
					B4: '=CELL("width",A1)',
					B6: '=CELL("contents",A1)',
					B7: '=CELL("type",C9)',
					B8: '=CELL("row")',
				},
				Data: { A1: 'x' },
			},
			{},
			(w) => {
				const styles = w.styles;
				styles.numFmts[164] = '0.00%';
				styles.cellXfs.push({
					numFmtId: 164,
					fontId: 0,
					fillId: 0,
					borderId: 0,
					protection: { locked: false },
				});
				const a1 = w.sheets[0].cells.get(0)!;
				w.sheets[0].cells.set(0, { ...a1, style: styles.cellXfs.length - 1 });
				w.sheets[0].cols.push({ min: 0, max: 0, width: 20.7109375, customWidth: true });
			},
		);
		expect(valueAt(wb, 'B1')).toBe('$A$1');
		expect(valueAt(wb, 'B2')).toBe('P2');
		expect(valueAt(wb, 'B3')).toBe(0);
		expect(valueAt(wb, 'B4')).toBe(20);
		expect(valueAt(wb, 'C4')).toBe(false);
		expect(valueAt(wb, 'B6')).toBe(1.5);
		expect(valueAt(wb, 'B7')).toBe('b');
		expect(valueAt(wb, 'B8')).toBe(8);
	});

	it('uses Excel defaults for unformatted cells and other sheets', () => {
		const { wb } = calculated({
			Sheet1: {
				A1: '=CELL("format",Z1)',
				A2: '=CELL("protect",Z1)',
				A3: '=CELL("address",Data!B2)',
				A4: '=CELL("width",Z1)',
			},
			Data: {},
		});
		expect(valueAt(wb, 'A1')).toBe('G');
		expect(valueAt(wb, 'A2')).toBe(1);
		expect(valueAt(wb, 'A3')).toBe('Data!$B$2');
		expect(valueAt(wb, 'A4')).toBe(8);
		expect(valueAt(wb, 'B4')).toBe(true);
	});
});
