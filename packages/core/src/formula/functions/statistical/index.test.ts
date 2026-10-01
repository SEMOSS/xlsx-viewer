import { describe, expect, it } from 'vitest';

import { XLFN_FUNCTIONS } from '../../prefixes';
import { FunctionRegistry } from '../../registry';
import { STATISTICAL_FUNCTIONS } from '.';

const EXPECTED = `AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS BETA.DIST BETA.INV BETADIST BETAINV
BINOM.DIST BINOM.DIST.RANGE BINOM.INV BINOMDIST CHIDIST CHIINV CHISQ.DIST CHISQ.DIST.RT
CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHITEST CONFIDENCE CONFIDENCE.NORM CONFIDENCE.T CORREL COUNT
COUNTA COUNTBLANK COUNTIF COUNTIFS COVAR COVARIANCE.P COVARIANCE.S CRITBINOM DEVSQ EXPON.DIST
EXPONDIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FDIST FINV FISHER FISHERINV FORECAST
FORECAST.LINEAR FREQUENCY FTEST GAMMA GAMMA.DIST GAMMA.INV GAMMADIST GAMMAINV GAMMALN
GAMMALN.PRECISE GAUSS GEOMEAN GROWTH HARMEAN HYPGEOM.DIST HYPGEOMDIST INTERCEPT KURT LARGE
LINEST LOGEST LOGINV LOGNORM.DIST LOGNORM.INV LOGNORMDIST MAX MAXA MAXIFS MEDIAN MIN MINA
MINIFS MODE MODE.MULT MODE.SNGL NEGBINOM.DIST NEGBINOMDIST NORM.DIST NORM.INV NORM.S.DIST
NORM.S.INV NORMDIST NORMINV NORMSDIST NORMSINV PEARSON PERCENTILE PERCENTILE.EXC PERCENTILE.INC
PERCENTRANK PERCENTRANK.EXC PERCENTRANK.INC PERMUT PERMUTATIONA PHI POISSON POISSON.DIST PROB
QUARTILE QUARTILE.EXC QUARTILE.INC RANK RANK.AVG RANK.EQ RSQ SKEW SKEW.P SLOPE SMALL STANDARDIZE
STDEV STDEV.P STDEV.S STDEVA STDEVP STDEVPA STEYX T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T
T.TEST TDIST TINV TREND TRIMMEAN TTEST VAR VAR.P VAR.S VARA VARP VARPA WEIBULL WEIBULL.DIST
Z.TEST ZTEST`.split(/\s+/);

describe('STATISTICAL_FUNCTIONS', () => {
	const names = STATISTICAL_FUNCTIONS.map((d) => d.name);

	it('defines every statistical function exactly once', () => {
		expect([...names].sort()).toEqual([...EXPECTED].sort());
		expect(new Set(names).size).toBe(names.length);
		expect(names).toHaveLength(145);
	});

	it('marks Excel 2010+ names with the _xlfn. file prefix', () => {
		for (const def of STATISTICAL_FUNCTIONS) {
			expect(def.filePrefix === '_xlfn.', def.name).toBe(XLFN_FUNCTIONS.has(def.name));
		}
		expect(STATISTICAL_FUNCTIONS.filter((d) => d.filePrefix)).toHaveLength(62);
	});

	it('flags the functions that return arrays', () => {
		const arrays = STATISTICAL_FUNCTIONS.filter((d) => d.returnsArray).map((d) => d.name);
		expect(arrays.sort()).toEqual([
			'FREQUENCY',
			'GROWTH',
			'LINEST',
			'LOGEST',
			'MODE.MULT',
			'TREND',
		]);
	});

	it('registers in a function registry', () => {
		const reg = new FunctionRegistry(STATISTICAL_FUNCTIONS);
		expect(reg.size).toBe(145);
		expect(reg.get('norm.s.dist')?.name).toBe('NORM.S.DIST');
	});

	it('declares consistent argument counts', () => {
		for (const def of STATISTICAL_FUNCTIONS) {
			expect(def.minArgs, def.name).toBeLessThanOrEqual(def.maxArgs);
			expect(def.params.length, def.name).toBeGreaterThan(0);
		}
	});
});
