import { describe, expect, it } from 'vitest';

import { XLFN_FUNCTIONS } from '../../prefixes';
import { ENGINEERING_FUNCTIONS } from './index';

const EXPECTED = [
	'BESSELI',
	'BESSELJ',
	'BESSELK',
	'BESSELY',
	'BIN2DEC',
	'BIN2HEX',
	'BIN2OCT',
	'BITAND',
	'BITLSHIFT',
	'BITOR',
	'BITRSHIFT',
	'BITXOR',
	'COMPLEX',
	'CONVERT',
	'DEC2BIN',
	'DEC2HEX',
	'DEC2OCT',
	'DELTA',
	'ERF',
	'ERF.PRECISE',
	'ERFC',
	'ERFC.PRECISE',
	'GESTEP',
	'HEX2BIN',
	'HEX2DEC',
	'HEX2OCT',
	'IMABS',
	'IMAGINARY',
	'IMARGUMENT',
	'IMCONJUGATE',
	'IMCOS',
	'IMCOSH',
	'IMCOT',
	'IMCSC',
	'IMCSCH',
	'IMDIV',
	'IMEXP',
	'IMLN',
	'IMLOG10',
	'IMLOG2',
	'IMPOWER',
	'IMPRODUCT',
	'IMREAL',
	'IMSEC',
	'IMSECH',
	'IMSIN',
	'IMSINH',
	'IMSQRT',
	'IMSUB',
	'IMSUM',
	'IMTAN',
	'OCT2BIN',
	'OCT2DEC',
	'OCT2HEX',
];

describe('ENGINEERING_FUNCTIONS', () => {
	it('registers every engineering function exactly once', () => {
		const names = ENGINEERING_FUNCTIONS.map((d) => d.name);
		expect(new Set(names).size).toBe(names.length);
		expect([...names].sort()).toEqual([...EXPECTED].sort());
	});

	it('marks Excel 2010+ functions with the _xlfn. file prefix', () => {
		for (const def of ENGINEERING_FUNCTIONS) {
			expect(def.filePrefix, def.name).toBe(XLFN_FUNCTIONS.has(def.name) ? '_xlfn.' : undefined);
			expect(def.minArgs).toBeLessThanOrEqual(def.maxArgs);
			expect(def.volatile).toBeUndefined();
		}
	});
});
