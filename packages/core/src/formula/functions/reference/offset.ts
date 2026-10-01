import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../../../address/constants';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { integerOf, isOmittedArg, singleArea } from './shared';

/**
 * OFFSET(reference, rows, cols, [height], [width]): a reference shifted by
 * rows / cols, optionally resized. Height and width default to the size of
 * the reference; negative sizes extend up / left from the shifted corner.
 * Results outside the grid are `#REF!`. Volatile, like Excel.
 */
const OFFSET: FunctionDefinition = {
	name: 'OFFSET',
	minArgs: 3,
	maxArgs: 5,
	params: ['ref', 'value', 'value', 'value', 'value'],
	volatile: true,
	impl: (args, ctx) => {
		const base = singleArea(args[0]);
		if (isCellError(base)) {
			return base;
		}
		if (base.sheetEnd !== undefined && base.sheetEnd !== base.sheet) {
			return ERRORS.VALUE;
		}
		const r = base.range;
		const rows = integerOf(args[1], ctx);
		if (isCellError(rows)) {
			return rows;
		}
		const cols = integerOf(args[2], ctx);
		if (isCellError(cols)) {
			return cols;
		}
		const height = isOmittedArg(args, 3) ? r.bottom - r.top + 1 : integerOf(args[3], ctx);
		if (isCellError(height)) {
			return height;
		}
		const width = isOmittedArg(args, 4) ? r.right - r.left + 1 : integerOf(args[4], ctx);
		if (isCellError(width)) {
			return width;
		}
		if (height === 0 || width === 0) {
			return ERRORS.REF;
		}
		const row0 = r.top + rows;
		const col0 = r.left + cols;
		const top = height > 0 ? row0 : row0 + height + 1;
		const bottom = height > 0 ? row0 + height - 1 : row0;
		const left = width > 0 ? col0 : col0 + width + 1;
		const right = width > 0 ? col0 + width - 1 : col0;
		if (top < 0 || left < 0 || bottom > MAX_ROW_INDEX || right > MAX_COL_INDEX) {
			return ERRORS.REF;
		}
		return { kind: 'ref', sheet: base.sheet, range: { top, left, bottom, right } };
	},
};

export const OFFSET_FUNCTIONS: FunctionDefinition[] = [OFFSET];
