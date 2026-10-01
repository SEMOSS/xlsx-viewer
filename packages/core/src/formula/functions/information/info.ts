import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionDefinition } from '../../types';
import { textArg } from '../math/args';

/**
 * INFO(type_text): information about the operating environment. The engine
 * reports what Excel for Windows reports for an automatically recalculated
 * workbook; file-system and memory details are not available (#N/A, which
 * is also what current Excel returns for the obsolete memory types).
 */
const INFO: FunctionDefinition = {
	name: 'INFO',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	volatile: true,
	impl: (args, ctx) => {
		const type = textArg(args[0]);
		if (isCellError(type)) {
			return type;
		}
		switch (type.trim().toLowerCase()) {
			case 'numfile':
				return ctx.sheetCount();
			case 'osversion':
				return 'Windows (64-bit) NT 10.00';
			case 'recalc':
				return 'Automatic';
			case 'release':
				return '16.0';
			case 'system':
				return 'pcdos';
			case 'origin':
				return '$A:$A$1';
			case 'directory':
			case 'memavail':
			case 'memused':
			case 'totmem':
				return ERRORS.NA;
			default:
				return ERRORS.VALUE;
		}
	},
};

export const INFO_FUNCTIONS: FunctionDefinition[] = [INFO];
