import type { FunctionDefinition } from '../types';
import { DATABASE_FUNCTIONS } from './database';
import { DATETIME_FUNCTIONS } from './datetime';
import { DYNAMIC_ARRAY_FUNCTIONS } from './dynamic-array';
import { ENGINEERING_FUNCTIONS } from './engineering';
import { FINANCIAL_FUNCTIONS } from './financial';
import { INFORMATION_FUNCTIONS } from './information';
import { LAMBDA_FUNCTIONS } from './lambda';
import { LOGICAL_FUNCTIONS } from './logical';
import { LOOKUP_FUNCTIONS } from './lookup';
import { MATH_FUNCTIONS } from './math';
import { MISC_FUNCTIONS } from './misc';
import { REFERENCE_FUNCTIONS } from './reference';
import { STATISTICAL_FUNCTIONS } from './statistical';
import { TEXT_FUNCTIONS } from './text';

/** Every built-in function module, in registration order (later modules win on name clashes). */
export const FUNCTION_MODULES: ReadonlyArray<readonly FunctionDefinition[]> = [
	LOGICAL_FUNCTIONS,
	REFERENCE_FUNCTIONS,
	LAMBDA_FUNCTIONS,
	MATH_FUNCTIONS,
	STATISTICAL_FUNCTIONS,
	TEXT_FUNCTIONS,
	INFORMATION_FUNCTIONS,
	LOOKUP_FUNCTIONS,
	DYNAMIC_ARRAY_FUNCTIONS,
	DATETIME_FUNCTIONS,
	FINANCIAL_FUNCTIONS,
	ENGINEERING_FUNCTIONS,
	DATABASE_FUNCTIONS,
	MISC_FUNCTIONS,
];

export {
	DATABASE_FUNCTIONS,
	DATETIME_FUNCTIONS,
	DYNAMIC_ARRAY_FUNCTIONS,
	ENGINEERING_FUNCTIONS,
	FINANCIAL_FUNCTIONS,
	INFORMATION_FUNCTIONS,
	LAMBDA_FUNCTIONS,
	LOGICAL_FUNCTIONS,
	LOOKUP_FUNCTIONS,
	MATH_FUNCTIONS,
	REFERENCE_FUNCTIONS,
	STATISTICAL_FUNCTIONS,
	TEXT_FUNCTIONS,
};
