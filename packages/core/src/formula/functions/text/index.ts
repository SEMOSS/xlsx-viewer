import type { FunctionDefinition } from '../../types';
import { BASIC_TEXT_FUNCTIONS } from './basic';
import { CHAR_FUNCTIONS } from './chars';
import { FORMAT_FUNCTIONS } from './format';
import { JOIN_FUNCTIONS } from './join';
import { REGEX_FUNCTIONS } from './regex';
import { SEARCH_FUNCTIONS } from './search';
import { SPLIT_FUNCTIONS } from './split';
import { SUBSTRING_FUNCTIONS } from './substring';
import { TO_TEXT_FUNCTIONS } from './to-text';

/** Text functions. */
export const TEXT_FUNCTIONS: FunctionDefinition[] = [
	...BASIC_TEXT_FUNCTIONS,
	...SUBSTRING_FUNCTIONS,
	...SEARCH_FUNCTIONS,
	...CHAR_FUNCTIONS,
	...JOIN_FUNCTIONS,
	...TO_TEXT_FUNCTIONS,
	...FORMAT_FUNCTIONS,
	...SPLIT_FUNCTIONS,
	...REGEX_FUNCTIONS,
];

export { formatDecimal, parseLocaleNumber } from './format';
