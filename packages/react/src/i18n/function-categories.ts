import type { FunctionCategory } from 'xlsx-viewer-core';

import type { TranslationKey } from './translations-en';

/**
 * Translation key of each function category. The core names categories in
 * English (`'Date & Time'`); the ribbon and Insert Function show these labels.
 */
export const FUNCTION_CATEGORY_KEYS: Readonly<Record<FunctionCategory, TranslationKey>> = {
	Financial: 'functionCategory.financial',
	'Date & Time': 'functionCategory.dateTime',
	'Math & Trig': 'functionCategory.mathTrig',
	Statistical: 'functionCategory.statistical',
	'Lookup & Reference': 'functionCategory.lookupReference',
	Database: 'functionCategory.database',
	Text: 'functionCategory.text',
	Logical: 'functionCategory.logical',
	Information: 'functionCategory.information',
	Engineering: 'functionCategory.engineering',
	Cube: 'functionCategory.cube',
	Compatibility: 'functionCategory.compatibility',
	Web: 'functionCategory.web',
};
