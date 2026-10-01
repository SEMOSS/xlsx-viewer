import type { FunctionDefinition } from '../../types';
import { CELL_FUNCTIONS } from './cell';
import { INFO_FUNCTIONS } from './info';
import { IS_FUNCTIONS } from './is';

/** Information functions (ISREF lives with the reference functions). */
export const INFORMATION_FUNCTIONS: FunctionDefinition[] = [
	...IS_FUNCTIONS,
	...CELL_FUNCTIONS,
	...INFO_FUNCTIONS,
];

export type { CellInfoProvider } from './cell';
export { cellFormatCode } from './cell-format';
