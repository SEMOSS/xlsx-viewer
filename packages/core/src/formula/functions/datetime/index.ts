import type { FunctionDefinition } from '../../types';
import { ARITHMETIC_FUNCTIONS } from './arithmetic';
import { COMPONENT_FUNCTIONS } from './components';
import { CONSTRUCT_FUNCTIONS } from './construct';
import { WORKDAY_FUNCTIONS } from './workdays';

/** Date and time functions (DATE, YEARFRAC, NETWORKDAYS.INTL, ...). */
export const DATETIME_FUNCTIONS: FunctionDefinition[] = [
	...CONSTRUCT_FUNCTIONS,
	...COMPONENT_FUNCTIONS,
	...ARITHMETIC_FUNCTIONS,
	...WORKDAY_FUNCTIONS,
];

export { yearFrac, type Basis } from './day-count';
