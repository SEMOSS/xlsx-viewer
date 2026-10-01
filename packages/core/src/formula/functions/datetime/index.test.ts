import { describe, expect, it } from 'vitest';

import { XLFN_FUNCTIONS } from '../../prefixes';
import { DATETIME_FUNCTIONS } from './index';

const EXPECTED = [
	'DATE',
	'DATEDIF',
	'DATEVALUE',
	'DAY',
	'DAYS',
	'DAYS360',
	'EDATE',
	'EOMONTH',
	'HOUR',
	'ISOWEEKNUM',
	'MINUTE',
	'MONTH',
	'NETWORKDAYS',
	'NETWORKDAYS.INTL',
	'NOW',
	'SECOND',
	'TIME',
	'TIMEVALUE',
	'TODAY',
	'WEEKDAY',
	'WEEKNUM',
	'WORKDAY',
	'WORKDAY.INTL',
	'YEAR',
	'YEARFRAC',
];

describe('DATETIME_FUNCTIONS', () => {
	it('registers every date and time function exactly once', () => {
		const names = DATETIME_FUNCTIONS.map((d) => d.name);
		expect(new Set(names).size).toBe(names.length);
		expect([...names].sort()).toEqual([...EXPECTED].sort());
	});

	it('marks Excel 2010+ functions with the _xlfn. file prefix and NOW / TODAY as volatile', () => {
		for (const def of DATETIME_FUNCTIONS) {
			expect(def.filePrefix, def.name).toBe(XLFN_FUNCTIONS.has(def.name) ? '_xlfn.' : undefined);
			expect(def.volatile === true, def.name).toBe(def.name === 'NOW' || def.name === 'TODAY');
			expect(def.minArgs).toBeLessThanOrEqual(def.maxArgs);
		}
	});
});
