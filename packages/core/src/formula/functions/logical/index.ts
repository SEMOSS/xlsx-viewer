import type { FunctionDefinition } from '../../types';
import { BOOLEAN_FUNCTIONS } from './boolean';
import { CONDITION_FUNCTIONS } from './conditions';

/** Logical functions: IF, IFS, IFERROR, IFNA, SWITCH, CHOOSE, AND, OR, XOR, NOT, TRUE, FALSE. */
export const LOGICAL_FUNCTIONS: FunctionDefinition[] = [
	...CONDITION_FUNCTIONS,
	...BOOLEAN_FUNCTIONS,
];

export {
	arrayModeOf,
	branchValue,
	buildArray,
	elementAt,
	evalLifted,
	liftedShape,
	liftValue,
} from './lift';
export { withFilePrefix } from './prefix';
