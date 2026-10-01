import type { FunctionDefinition } from '../../types';
import { ADDRESS_FUNCTIONS } from './address';
import { INDEX_FUNCTIONS } from './index-fn';
import { INDIRECT_FUNCTIONS } from './indirect';
import { OFFSET_FUNCTIONS } from './offset';
import { POSITION_FUNCTIONS } from './position';
import { SHEET_FUNCTIONS } from './sheet';

/**
 * Reference functions: ROW, ROWS, COLUMN, COLUMNS, AREAS, OFFSET, INDIRECT,
 * INDEX, ADDRESS, SHEET, SHEETS, FORMULATEXT, ISREF.
 */
export const REFERENCE_FUNCTIONS: FunctionDefinition[] = [
	...POSITION_FUNCTIONS,
	...OFFSET_FUNCTIONS,
	...INDIRECT_FUNCTIONS,
	...INDEX_FUNCTIONS,
	...ADDRESS_FUNCTIONS,
	...SHEET_FUNCTIONS,
];
