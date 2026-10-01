/** Every catalog entry, gathered from the per-category data files in dialog order. */
import { COMPATIBILITY_INFO } from './compatibility';
import { CUBE_INFO } from './cube';
import { DATABASE_INFO } from './database';
import { DATETIME_INFO } from './datetime';
import { ENGINEERING_INFO } from './engineering';
import { COMPLEX_INFO } from './engineering-complex';
import { FINANCIAL_INFO } from './financial';
import { SECURITIES_INFO } from './financial-securities';
import { INFORMATION_INFO } from './information';
import { LOGICAL_INFO } from './logical';
import { LOOKUP_INFO } from './lookup';
import { LOOKUP_ARRAY_INFO } from './lookup-arrays';
import { MATH_INFO } from './math';
import { MATH_TRIG_INFO } from './math-trig';
import { ANALYSIS_INFO } from './statistical-analysis';
import { DISTRIBUTION_INFO } from './statistical-distributions';
import { STATISTICAL_INFO } from './statistical';
import { TEXT_INFO } from './text';
import { TEXT_MODERN_INFO } from './text-modern';
import type { FunctionInfo } from './types';
import { WEB_INFO } from './web';

/** All entries; a name must appear only once (the catalog tests check it). */
export const ALL_FUNCTION_INFO: readonly FunctionInfo[] = [
	...FINANCIAL_INFO,
	...SECURITIES_INFO,
	...DATETIME_INFO,
	...MATH_INFO,
	...MATH_TRIG_INFO,
	...STATISTICAL_INFO,
	...DISTRIBUTION_INFO,
	...ANALYSIS_INFO,
	...LOOKUP_INFO,
	...LOOKUP_ARRAY_INFO,
	...DATABASE_INFO,
	...TEXT_INFO,
	...TEXT_MODERN_INFO,
	...LOGICAL_INFO,
	...INFORMATION_INFO,
	...ENGINEERING_INFO,
	...COMPLEX_INFO,
	...CUBE_INFO,
	...COMPATIBILITY_INFO,
	...WEB_INFO,
];
