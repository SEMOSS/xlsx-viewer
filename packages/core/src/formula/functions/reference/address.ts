import { columnToLetters } from '../../../address/cell-address';
import { MAX_COLS, MAX_ROWS } from '../../../address/constants';
import { quoteSheetName } from '../../../address/sheet-name';
import { ERRORS, isCellError } from '../../../values/errors';
import { toBoolean, toText } from '../../coerce';
import type { FunctionDefinition, Scalar } from '../../types';
import { integerOf, isOmittedArg } from './shared';

/**
 * ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text]): reference
 * text. abs_num 1 = `$A$1`, 2 = `A$1`, 3 = `$A1`, 4 = `A1`; in R1C1 style
 * the relative parts are written in brackets (`R[1]C[1]`).
 */

function a1Text(row: number, col: number, absRow: boolean, absCol: boolean): string {
	return `${absCol ? '$' : ''}${columnToLetters(col - 1)}${absRow ? '$' : ''}${row}`;
}

function r1c1Text(row: number, col: number, absRow: boolean, absCol: boolean): string {
	return `R${absRow ? row : `[${row}]`}C${absCol ? col : `[${col}]`}`;
}

const ADDRESS: FunctionDefinition = {
	name: 'ADDRESS',
	minArgs: 2,
	maxArgs: 5,
	params: ['value'],
	impl: (args, ctx) => {
		const row = integerOf(args[0], ctx);
		if (isCellError(row)) {
			return row;
		}
		const col = integerOf(args[1], ctx);
		if (isCellError(col)) {
			return col;
		}
		const abs = isOmittedArg(args, 2) ? 1 : integerOf(args[2], ctx);
		if (isCellError(abs)) {
			return abs;
		}
		const a1 = isOmittedArg(args, 3) ? true : toBoolean(args[3] as Scalar);
		if (isCellError(a1)) {
			return a1;
		}
		if (row < 1 || row > MAX_ROWS || col < 1 || col > MAX_COLS || abs < 1 || abs > 4) {
			return ERRORS.VALUE;
		}
		const absRow = abs === 1 || abs === 2;
		const absCol = abs === 1 || abs === 3;
		const text = a1 ? a1Text(row, col, absRow, absCol) : r1c1Text(row, col, absRow, absCol);
		if (args.length < 5 || isOmittedArg(args, 4)) {
			return text;
		}
		const sheet = toText(args[4] as Scalar);
		if (isCellError(sheet)) {
			return sheet;
		}
		return sheet === '' ? `!${text}` : `${quoteSheetName(sheet)}!${text}`;
	},
};

export const ADDRESS_FUNCTIONS: FunctionDefinition[] = [ADDRESS];
