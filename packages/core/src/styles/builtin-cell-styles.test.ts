import { describe, expect, it } from 'vitest';

import { createWorkbook } from '../sheet/workbook';
import { setCellInput } from '../edit/cell-input';
import { applyCellStyle } from '../edit/cell-style-ops';
import { applyStylePatch } from '../edit/style-ops';
import { EditTransaction } from '../edit/transaction';
import {
	BUILTIN_CELL_STYLES,
	cellStyleIncludes,
	ensureCellStyle,
	findCellStyle,
	getBuiltinCellStyle,
} from './builtin-cell-styles';
import { StyleResolver } from './resolve-style';

const A1 = { top: 0, left: 0, bottom: 0, right: 0 };

describe('built-in cell styles', () => {
	it('defines the Excel gallery with unique ids', () => {
		expect(BUILTIN_CELL_STYLES).toHaveLength(47);
		expect(new Set(BUILTIN_CELL_STYLES.map((s) => s.builtinId)).size).toBe(47);
		expect(getBuiltinCellStyle('20% - accent3')).toMatchObject({ builtinId: 38 });
		expect(getBuiltinCellStyle('Accent6')).toMatchObject({ builtinId: 49 });
	});

	it('adds a missing built-in style once, only appending', () => {
		const wb = createWorkbook();
		const styles = wb.styles;
		const before = {
			fonts: styles.fonts.length,
			xfs: styles.cellStyleXfs.length,
			named: styles.cellStyles.length,
		};
		const id = ensureCellStyle(styles, 'Good', wb.theme)!;
		expect(styles.cellStyleXfs.length).toBe(before.xfs + 1);
		expect(styles.cellStyles.length).toBe(before.named + 1);
		expect(styles.fonts.length).toBe(before.fonts + 1);
		expect(ensureCellStyle(styles, 'good', wb.theme)).toBe(id);
		expect(findCellStyle(styles, 'Good')).toMatchObject({ name: 'Good', builtinId: 26, xfId: id });
		expect(cellStyleIncludes(styles.cellStyleXfs[id])).toMatchObject({
			font: true,
			fill: true,
			numberFormat: false,
			border: false,
		});
		expect(ensureCellStyle(styles, 'No Such Style')).toBeUndefined();
	});

	it('applies only what a style includes', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Style');
		setCellInput(tx, wb, 0, 0, 0, '12.5');
		applyStylePatch(tx, wb, 0, [A1], { numFmt: '0.00%', alignment: { horizontal: 'center' } });
		applyCellStyle(tx, wb, 0, [A1], 'Bad');
		const entry = tx.commit();
		const resolver = new StyleResolver(wb.styles, wb.theme);
		const cellStyle = wb.sheets[0].cells.get(0)?.style;
		const resolved = resolver.resolve(cellStyle);
		expect(resolved.font.color).toEqual({ rgb: 'FF9C0006' });
		expect(resolved.fill).toMatchObject({ patternType: 'solid', fgColor: { rgb: 'FFFFC7CE' } });
		expect(resolved.numFmt).toBe('0.00%');
		expect(resolved.alignment.horizontal).toBe('center');
		expect(wb.styles.cellXfs[cellStyle!].xfId).toBe(findCellStyle(wb.styles, 'Bad')!.xfId);
		entry.undo();
		expect(wb.sheets[0].cells.get(0)).toBeUndefined();
	});

	it('Normal resets every aspect; number-format styles change only the format', () => {
		const wb = createWorkbook();
		const tx = new EditTransaction(wb, 'Style');
		setCellInput(tx, wb, 0, 0, 0, '3');
		applyCellStyle(tx, wb, 0, [A1], 'Heading 1');
		applyCellStyle(tx, wb, 0, [A1], 'Currency');
		const resolver = () => new StyleResolver(wb.styles, wb.theme);
		let resolved = resolver().resolve(wb.sheets[0].cells.get(0)?.style);
		expect(resolved.font.bold).toBe(true);
		expect(resolved.numFmtId).toBe(44);
		applyCellStyle(tx, wb, 0, [A1], 'Normal');
		resolved = resolver().resolve(wb.sheets[0].cells.get(0)?.style);
		expect(resolved.font.bold).toBeFalsy();
		expect(resolved.numFmtId).toBe(0);
		expect(resolved.border.bottom).toBeUndefined();
	});

	it('uses the theme heading font for Title', () => {
		const wb = createWorkbook();
		const id = ensureCellStyle(wb.styles, 'Title', wb.theme)!;
		const font = wb.styles.fonts[wb.styles.cellStyleXfs[id].fontId];
		expect(font).toMatchObject({ size: 18, scheme: 'major', name: wb.theme.fonts.majorLatin });
	});
});
