import { describe, expect, it } from 'vitest';

import { XLFN_FUNCTIONS } from '../../prefixes';
import { FINANCIAL_FUNCTIONS } from './index';

const EXPECTED = [
	'ACCRINT',
	'ACCRINTM',
	'AMORDEGRC',
	'AMORLINC',
	'COUPDAYBS',
	'COUPDAYS',
	'COUPDAYSNC',
	'COUPNCD',
	'COUPNUM',
	'COUPPCD',
	'CUMIPMT',
	'CUMPRINC',
	'DB',
	'DDB',
	'DISC',
	'DOLLARDE',
	'DOLLARFR',
	'DURATION',
	'EFFECT',
	'FV',
	'FVSCHEDULE',
	'INTRATE',
	'IPMT',
	'IRR',
	'ISPMT',
	'MDURATION',
	'MIRR',
	'NOMINAL',
	'NPER',
	'NPV',
	'ODDFPRICE',
	'ODDFYIELD',
	'ODDLPRICE',
	'ODDLYIELD',
	'PDURATION',
	'PMT',
	'PPMT',
	'PRICE',
	'PRICEDISC',
	'PRICEMAT',
	'PV',
	'RATE',
	'RECEIVED',
	'RRI',
	'SLN',
	'SYD',
	'TBILLEQ',
	'TBILLPRICE',
	'TBILLYIELD',
	'VDB',
	'XIRR',
	'XNPV',
	'YIELD',
	'YIELDDISC',
	'YIELDMAT',
];

describe('FINANCIAL_FUNCTIONS', () => {
	it('registers every financial function exactly once', () => {
		const names = FINANCIAL_FUNCTIONS.map((d) => d.name);
		expect(new Set(names).size).toBe(names.length);
		expect([...names].sort()).toEqual([...EXPECTED].sort());
	});

	it('marks Excel 2010+ functions with the _xlfn. file prefix', () => {
		for (const def of FINANCIAL_FUNCTIONS) {
			expect(def.filePrefix, def.name).toBe(XLFN_FUNCTIONS.has(def.name) ? '_xlfn.' : undefined);
			expect(def.minArgs).toBeLessThanOrEqual(def.maxArgs);
			expect(def.volatile).toBeUndefined();
		}
	});
});
