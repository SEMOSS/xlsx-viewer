import { describe, expect, it } from 'vitest';

import { computeOutline } from '../sheet/outline';
import { createWorkbook } from '../sheet/workbook';
import { groupLines, setGroupCollapsed, showOutlineLevel, ungroupLines } from './outline-ops';
import { EditTransaction } from './transaction';

const run = (wb: ReturnType<typeof createWorkbook>, fn: (tx: EditTransaction) => void) => {
	const tx = new EditTransaction(wb, 'Outline');
	fn(tx);
	return tx.commit();
};

describe('outline operations', () => {
	it('groups and ungroups rows, tracking the deepest level, with undo', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		run(wb, (tx) => groupLines(tx, wb, 0, 'rows', 1, 4));
		run(wb, (tx) => groupLines(tx, wb, 0, 'rows', 2, 3));
		expect(sheet.rows.get(2)?.outlineLevel).toBe(2);
		expect(sheet.outlineLevelRow).toBe(2);
		const entry = run(wb, (tx) => ungroupLines(tx, wb, 0, 'rows', 2, 3));
		expect(sheet.outlineLevelRow).toBe(1);
		expect(sheet.rows.get(2)?.outlineLevel).toBe(1);
		entry.undo();
		expect(sheet.rows.get(2)?.outlineLevel).toBe(2);
		expect(sheet.outlineLevelRow).toBe(2);
	});

	it('collapses and expands a group, keeping collapsed inner groups hidden', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		run(wb, (tx) => {
			groupLines(tx, wb, 0, 'rows', 1, 5);
			groupLines(tx, wb, 0, 'rows', 2, 3);
		});
		const inner = computeOutline(sheet, 'rows').groups.find((g) => g.level === 2)!;
		run(wb, (tx) => setGroupCollapsed(tx, wb, 0, 'rows', inner, true));
		expect(sheet.rows.get(2)?.hidden).toBe(true);
		expect(sheet.rows.get(4)?.collapsed).toBe(true);
		const outer = computeOutline(sheet, 'rows').groups.find((g) => g.level === 1)!;
		run(wb, (tx) => setGroupCollapsed(tx, wb, 0, 'rows', outer, true));
		expect([1, 2, 3, 4, 5].every((r) => sheet.rows.get(r)?.hidden)).toBe(true);
		run(wb, (tx) => setGroupCollapsed(tx, wb, 0, 'rows', outer, false));
		expect(sheet.rows.get(1)?.hidden).toBeFalsy();
		expect(sheet.rows.get(2)?.hidden).toBe(true);
		expect(sheet.rows.get(4)?.hidden).toBeFalsy();
	});

	it('shows outline levels with the level buttons, for columns too', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		run(wb, (tx) => {
			groupLines(tx, wb, 0, 'cols', 1, 4);
			groupLines(tx, wb, 0, 'cols', 2, 3);
		});
		run(wb, (tx) => showOutlineLevel(tx, wb, 0, 'cols', 2));
		const hiddenCols = (i: number) =>
			sheet.cols.find((c) => i >= c.min && i <= c.max)?.hidden ?? false;
		expect([1, 2, 3, 4].map(hiddenCols)).toEqual([false, true, true, false]);
		run(wb, (tx) => showOutlineLevel(tx, wb, 0, 'cols', 1));
		expect([1, 2, 3, 4].map(hiddenCols)).toEqual([true, true, true, true]);
		run(wb, (tx) => showOutlineLevel(tx, wb, 0, 'cols', 3));
		expect([1, 2, 3, 4].map(hiddenCols)).toEqual([false, false, false, false]);
	});
});
