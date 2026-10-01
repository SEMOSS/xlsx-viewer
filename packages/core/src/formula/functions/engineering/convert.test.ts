import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, fn, MockContext } from '../../test-utils/mock-context';
import type { FunctionArg } from '../../types';
import { resolveUnit } from './convert';
import { UNITS } from './convert-units';
import { ENGINEERING_FUNCTIONS } from './index';

const ctx = new MockContext();
const convert = (...args: FunctionArg[]) => call(fn(ENGINEERING_FUNCTIONS, 'CONVERT'), args, ctx);

describe('CONVERT documented examples', () => {
	it('match Excel', () => {
		expect(convert(1, 'lbm', 'kg')).toBe(0.45359237);
		expect(convert(68, 'F', 'C')).toBe(20);
		expect(convert(2.5, 'ft', 'sec')).toBe(ERRORS.NA);
		expect(convert(convert(100, 'ft', 'm') as number, 'ft', 'm')).toBeCloseTo(9.290304, 12);
	});
});

describe('CONVERT by category', () => {
	it.each([
		// mass
		[1, 'kg', 'lbm', 2.20462262184878],
		[1, 'ozm', 'g', 28.349523125],
		[1, 'stone', 'lbm', 14],
		[1, 'ton', 'lbm', 2000],
		[1, 'uk_ton', 'LTON', 1],
		[1, 'cwt', 'lbm', 100],
		[1, 'uk_cwt', 'lbm', 112],
		[1, 'sg', 'kg', 14.5939029372064],
		[7000, 'grain', 'lbm', 1],
		// distance
		[1, 'in', 'cm', 2.54],
		[1, 'mi', 'km', 1.609344],
		[1, 'mi', 'ft', 5280],
		[1, 'Nmi', 'm', 1852],
		[1, 'yd', 'in', 36],
		[1, 'ell', 'in', 45],
		[72, 'Pica', 'in', 1],
		[6, 'pica', 'in', 1],
		[1, 'ly', 'km', 9460730472580.8],
		[1, 'ang', 'nm', 0.1],
		[1, 'survey_mi', 'mi', 1.000002000004],
		// time
		[1, 'yr', 'day', 365.25],
		[1, 'day', 'hr', 24],
		[6, 'mn', 'hr', 0.1],
		[1, 'hr', 'sec', 3600],
		[1, 'ms', 's', 0.001],
		// pressure
		[1, 'atm', 'Pa', 101325],
		[1, 'atm', 'mmHg', 760.002100178515],
		[1, 'atm', 'Torr', 760],
		[1, 'psi', 'Pa', 6894.75729316836],
		[1, 'kPa', 'Pa', 1000],
		// force
		[1, 'lbf', 'N', 4.4482216152605],
		[1, 'N', 'dyn', 100000],
		[1, 'pond', 'N', 0.00980665],
		// energy
		[1, 'cal', 'J', 4.1868],
		[1, 'c', 'J', 4.184],
		[1, 'kWh', 'J', 3600000],
		[1, 'BTU', 'J', 1055.05585262],
		[1, 'J', 'e', 10000000],
		[1, 'eV', 'J', 1.602176634e-19],
		[1, 'HPh', 'Wh', 745.69987158227],
		[1, 'flb', 'J', 0.0421401100938048],
		// power
		[1, 'HP', 'W', 745.69987158227],
		[1, 'PS', 'W', 735.49875],
		[1, 'kW', 'W', 1000],
		// magnetism
		[1, 'T', 'ga', 10000],
		// volume
		[1, 'gal', 'l', 3.785411784],
		[1, 'gal', 'qt', 4],
		[1, 'qt', 'pt', 2],
		[1, 'pt', 'cup', 2],
		[1, 'cup', 'oz', 8],
		[1, 'oz', 'tbs', 2],
		[1, 'tbs', 'tsp', 3],
		[1, 'uk_gal', 'l', 4.54609],
		[1, 'uk_gal', 'uk_pt', 8],
		[1, 'barrel', 'gal', 42],
		[1, 'm3', 'l', 1000],
		[1, 'ft3', 'in3', 1728],
		[1, 'ml', 'l', 0.001],
		[1, 'GRT', 'ft3', 100],
		[1, 'MTON', 'ft3', 40],
		[1, 'tspm', 'ml', 5],
		// area
		[1, 'ha', 'm2', 10000],
		[1, 'ar', 'm^2', 100],
		[1, 'uk_acre', 'ft2', 43560],
		[1, 'mi2', 'uk_acre', 640],
		[1, 'Morgen', 'm2', 2500],
		[1, 'ft^2', 'in2', 144],
		[1, 'yd2', 'ft2', 9],
		// information
		[1, 'byte', 'bit', 8],
		[1, 'kbyte', 'byte', 1000],
		[1, 'kibyte', 'byte', 1024],
		[1, 'Mibyte', 'kibyte', 1024],
		[1, 'Gibit', 'bit', 1073741824],
		// speed
		[1, 'kn', 'm/h', 1852],
		[1, 'mph', 'm/s', 0.44704],
		[1, 'admkn', 'kn', 1.00063930885529],
		[1, 'km/h', 'm/s', 0.277777777777778],
		[1, 'm/hr', 'm/h', 1],
	])('%d %s = %d %s', (value, from, to, expected) => {
		const got = convert(value, from as string, to as string) as number;
		expect(Math.abs(got - (expected as number))).toBeLessThanOrEqual(
			Math.abs(expected as number) * 1e-13,
		);
	});
});

describe('prefixes', () => {
	it('apply powers of the prefix to areas and volumes', () => {
		expect(convert(1, 'km2', 'm2')).toBe(1e6);
		expect(convert(1, 'km^2', 'm2')).toBe(1e6);
		expect(convert(1, 'cm3', 'm3')).toBe(1e-6);
		expect(convert(1, 'm3', 'dm^3')).toBe(1000);
		expect(convert(1, 'kang2', 'ang2')).toBe(1e6);
	});

	it('cover the whole decimal range', () => {
		expect(convert(1, 'Ym', 'm')).toBe(1e24);
		expect(convert(1, 'ym', 'm')).toBe(1e-24);
		expect(convert(1, 'dam', 'm')).toBe(10);
		expect(convert(1, 'em', 'm')).toBe(10);
		expect(convert(1, 'hm', 'm')).toBe(100);
		expect(convert(1, 'um', 'nm')).toBe(1000);
		expect(convert(1, 'Gs', 'Ms')).toBe(1000);
	});

	it('only applies to metric units', () => {
		expect(convert(1, 'kft', 'm')).toBe(ERRORS.NA);
		expect(convert(1, 'kmi', 'm')).toBe(ERRORS.NA);
		expect(convert(1, 'kibit', 'bit')).toBe(1024);
		expect(convert(1, 'kim', 'm')).toBe(ERRORS.NA);
		expect(convert(1, 'kC', 'C')).toBe(ERRORS.NA);
	});

	it('prefer exact unit names', () => {
		expect(resolveUnit('min')).toMatchObject({ base: 'min', exp10: 0 });
		expect(resolveUnit('mi')).toMatchObject({ base: 'mi', exp10: 0 });
		expect(resolveUnit('Pa')).toMatchObject({ base: 'Pa', exp10: 0 });
		expect(resolveUnit('pm')).toMatchObject({ base: 'm', exp10: -12 });
		expect(resolveUnit('cm3')).toMatchObject({ base: 'm3', exp10: -6 });
		expect(resolveUnit('Gibyte')).toMatchObject({ base: 'byte', exp2: 30 });
		expect(resolveUnit('xyz')).toBeNull();
	});
});

describe('temperature', () => {
	it.each([
		[100, 'C', 'F', 212],
		[0, 'C', 'K', 273.15],
		[32, 'F', 'C', 0],
		[-40, 'F', 'C', -40],
		[0, 'K', 'F', -459.67],
		[100, 'cel', 'kel', 373.15],
		[0, 'C', 'Rank', 491.67],
		[80, 'Reau', 'C', 100],
		[100, 'C', 'Reau', 80],
		[491.67, 'Rank', 'F', 32],
		[1000, 'mK', 'K', 1],
		[1, 'kK', 'C', 726.85],
		[212, 'fah', 'cel', 100],
	])('%d %s = %d %s', (value, from, to, expected) => {
		expect(convert(value, from as string, to as string)).toBeCloseTo(expected as number, 10);
	});
});

describe('CONVERT validation', () => {
	it('returns #N/A for unknown units and mismatched categories', () => {
		expect(convert(1, 'm', 'kg')).toBe(ERRORS.NA);
		expect(convert(1, 'M', 'm')).toBe(ERRORS.NA);
		expect(convert(1, 'LBM', 'kg')).toBe(ERRORS.NA);
		expect(convert(1, '', 'm')).toBe(ERRORS.NA);
		expect(convert(1, 'm', 1)).toBe(ERRORS.NA);
		expect(convert(1, 'C', 'J')).toBe(ERRORS.NA);
	});

	it('requires a number and keeps the value for identical units', () => {
		expect(convert('x', 'm', 'ft')).toBe(ERRORS.VALUE);
		expect(convert(true, 'm', 'ft')).toBe(ERRORS.VALUE);
		expect(convert('12', 'in', 'ft')).toBe(1);
		expect(convert(12.5, 'm', 'm')).toBe(12.5);
		expect(convert(ERRORS.DIV0, 'm', 'ft')).toBe(ERRORS.DIV0);
	});

	it('lifts over arrays', () => {
		expect(convert(arr([0, 100]), 'C', 'F')).toEqual(arr([32, 212]));
	});

	it('has a definition for every documented unit name', () => {
		for (const name of [
			'g',
			'sg',
			'lbm',
			'u',
			'ozm',
			'grain',
			'cwt',
			'shweight',
			'uk_cwt',
			'lcwt',
			'hweight',
			'stone',
			'ton',
			'uk_ton',
			'LTON',
			'brton',
		]) {
			expect(UNITS.get(name)?.category, name).toBe('mass');
		}
		expect(UNITS.size).toBeGreaterThan(130);
	});
});
