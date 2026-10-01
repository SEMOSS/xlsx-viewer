import { describe, expect, it } from 'vitest';
import { XLFN_FUNCTIONS, XLWS_FUNCTIONS } from '../language/function-prefixes';
import { FunctionRegistry } from '../registry';
import { ALL_FUNCTION_INFO } from './entries';
import {
	formatFunctionSyntax,
	FUNCTION_CATEGORIES,
	type FunctionInfo,
	getFunctionInfo,
	listFunctions,
	searchFunctions,
} from './index';
import { CATALOG_VERSIONS, versionTable } from './versions';

/** File-format encodings of operators (`A1#` and `@x`), not worksheet functions. */
const INTERNAL_NAMES = new Set(['ANCHORARRAY', 'SINGLE']);

/** Names in the file-prefix table that Excel does not offer as worksheet functions. */
const HIDDEN_PREFIXED = new Set([...INTERNAL_NAMES, 'ECMA.CEILING', 'QUERYSTRING']);

/** Excel's own syntax puts the optional starting accumulator before the required arguments. */
const LEADING_OPTIONAL = new Set(['REDUCE', 'SCAN']);

function info(name: string): FunctionInfo {
	const found = getFunctionInfo(name);
	if (!found) {
		throw new Error(`no catalog entry for ${name}`);
	}
	return found;
}

describe('function catalog data', () => {
	it('has no duplicate names', () => {
		const seen = new Set<string>();
		const duplicates = ALL_FUNCTION_INFO.map((f) => f.name).filter(
			(name) => seen.has(name) || !seen.add(name),
		);
		expect(duplicates, `duplicate entries: ${duplicates.join(', ')}`).toEqual([]);
		expect(listFunctions()).toHaveLength(ALL_FUNCTION_INFO.length);
	});

	it('has an entry for every function the engine registers', () => {
		const missing = FunctionRegistry.createDefault()
			.names()
			.filter((name) => !INTERNAL_NAMES.has(name) && !getFunctionInfo(name));
		expect(missing, `functions without catalog help: ${missing.join(', ')}`).toEqual([]);
	});

	it('covers the full Excel function library, including the newest functions', () => {
		expect(ALL_FUNCTION_INFO.length).toBeGreaterThanOrEqual(500);
		const newest = [
			'XLOOKUP',
			'LET',
			'LAMBDA',
			'BYROW',
			'ISOMITTED',
			'TEXTSPLIT',
			'TEXTBEFORE',
			'TEXTAFTER',
			'REGEXTEST',
			'REGEXEXTRACT',
			'REGEXREPLACE',
			'GROUPBY',
			'PIVOTBY',
			'PERCENTOF',
			'TRIMRANGE',
			'IMAGE',
			'STOCKHISTORY',
			'CUBEVALUE',
			'WEBSERVICE',
		];
		expect(newest.filter((name) => !getFunctionInfo(name))).toEqual([]);
		const prefixed = [...XLFN_FUNCTIONS, ...XLWS_FUNCTIONS].filter((n) => !HIDDEN_PREFIXED.has(n));
		expect(prefixed.filter((name) => !getFunctionInfo(name))).toEqual([]);
	});

	it('uses upper-case names and Insert Function categories', () => {
		for (const f of ALL_FUNCTION_INFO) {
			expect(f.name, f.name).toMatch(/^[A-Z][A-Z0-9]*(\.[A-Z0-9]+)*$/);
			expect(FUNCTION_CATEGORIES, f.name).toContain(f.category);
		}
		for (const category of FUNCTION_CATEGORIES) {
			expect(listFunctions(category).length, category).toBeGreaterThan(0);
		}
	});

	it('describes every function in one concise sentence', () => {
		for (const f of ALL_FUNCTION_INFO) {
			expect(f.description, f.name).toMatch(/^[A-Z].*\.$/);
			expect(f.description, f.name).not.toMatch(/\.\s+[A-Z]/);
			expect(f.description.length, f.name).toBeLessThanOrEqual(120);
		}
	});

	it('names and describes every argument', () => {
		for (const f of ALL_FUNCTION_INFO) {
			const names = f.args.map((a) => a.name);
			expect(new Set(names).size, `${f.name} repeats an argument name`).toBe(names.length);
			for (const a of f.args) {
				expect(a.name, f.name).toMatch(/^[A-Za-z][A-Za-z0-9_']*$/);
				expect(a.description.trim(), `${f.name}(${a.name})`).not.toBe('');
				expect(a.description, `${f.name}(${a.name})`).toMatch(/\.$/);
				if (a.repeating) {
					expect(a.optional, `${f.name}(${a.name}) repeats but is required`).toBe(true);
				}
			}
		}
	});

	it('never puts a required argument after an optional one, except around repeating groups', () => {
		const offenders: string[] = [];
		for (const f of ALL_FUNCTION_INFO) {
			if (LEADING_OPTIONAL.has(f.name)) {
				continue;
			}
			let seenOptional = false;
			for (const a of f.args) {
				if (a.repeating) {
					continue;
				}
				if (a.optional) {
					seenOptional = true;
				} else if (seenOptional) {
					offenders.push(`${f.name}(${a.name})`);
				}
			}
		}
		expect(offenders).toEqual([]);
		expect(info('REDUCE').args.map((a) => a.optional === true)).toEqual([true, false, false]);
	});

	it('contains no em-dash characters', () => {
		const text = JSON.stringify(ALL_FUNCTION_INFO);
		expect(text.includes('\u2014')).toBe(false);
	});

	it('tags functions added after Excel 2007 with their version', () => {
		const table = versionTable();
		const names = table.map(([name]) => name);
		expect(names.filter((name, i) => names.indexOf(name) !== i)).toEqual([]);
		expect(names.filter((name) => !getFunctionInfo(name))).toEqual([]);
		for (const f of ALL_FUNCTION_INFO) {
			if (f.since !== undefined) {
				expect(CATALOG_VERSIONS, f.name).toContain(f.since);
			}
		}
		const prefixedWithoutVersion = [...XLFN_FUNCTIONS, ...XLWS_FUNCTIONS].filter((name) => {
			const f = getFunctionInfo(name);
			return f !== undefined && f.since === undefined;
		});
		expect(prefixedWithoutVersion).toEqual([]);
		expect(info('XLOOKUP').since).toBe('2021');
		expect(info('IFS').since).toBe('2019');
		expect(info('NORM.DIST').since).toBe('2010');
		expect(info('LAMBDA').since).toBe('365');
		expect(info('VLOOKUP').since).toBeUndefined();
	});

	it('assigns the expected categories', () => {
		expect(info('VLOOKUP').category).toBe('Lookup & Reference');
		expect(info('SUM').category).toBe('Math & Trig');
		expect(info('NORMDIST').category).toBe('Compatibility');
		expect(info('NORM.DIST').category).toBe('Statistical');
		expect(info('CONCATENATE').category).toBe('Compatibility');
		expect(info('ISOMITTED').category).toBe('Information');
		expect(info('MAP').category).toBe('Logical');
		expect(info('CUBEMEMBER').category).toBe('Cube');
		expect(info('FILTERXML').category).toBe('Web');
		expect(info('DGET').category).toBe('Database');
		expect(info('IMSUM').category).toBe('Engineering');
		expect(info('YEARFRAC').category).toBe('Date & Time');
		expect(info('XNPV').category).toBe('Financial');
		expect(info('TEXTSPLIT').category).toBe('Text');
	});
});

describe('formatFunctionSyntax', () => {
	it.each([
		['VLOOKUP', 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])'],
		['SUM', 'SUM(number1, [number2], ...)'],
		['IF', 'IF(logical_test, [value_if_true], [value_if_false])'],
		[
			'XLOOKUP',
			'XLOOKUP(lookup_value, lookup_array, return_array, ' +
				'[if_not_found], [match_mode], [search_mode])',
		],
		[
			'LET',
			'LET(name1, name_value1, calculation_or_name2, [name_value2, calculation_or_name3], ...)',
		],
	])('matches the Excel syntax line of %s', (name, expected) => {
		expect(formatFunctionSyntax(info(name))).toBe(expected);
	});

	it.each([
		['NOW', 'NOW()'],
		['SUMIFS', 'SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)'],
		['TEXTJOIN', 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)'],
		['GETPIVOTDATA', 'GETPIVOTDATA(data_field, pivot_table, [field1, item1], ...)'],
		['LAMBDA', 'LAMBDA([parameter1, ...], calculation)'],
		['MAP', 'MAP(array1, [array2, ...], lambda)'],
		['REDUCE', 'REDUCE([initial_value], array, lambda)'],
		['NORM.S.DIST', 'NORM.S.DIST(z, cumulative)'],
		['FORECAST.LINEAR', "FORECAST.LINEAR(x, known_y's, known_x's)"],
	])('formats %s', (name, expected) => {
		expect(formatFunctionSyntax(info(name))).toBe(expected);
	});

	it('renders a required repeating group without brackets', () => {
		const custom: FunctionInfo = {
			name: 'DEMO',
			category: 'Logical',
			description: 'Demo.',
			args: [
				{ name: 'a', description: 'A.' },
				{ name: 'b', description: 'B.', repeating: true },
			],
		};
		expect(formatFunctionSyntax(custom)).toBe('DEMO(a, b, ...)');
	});
});

describe('catalog lookups', () => {
	it('finds functions by any case and ignores file prefixes', () => {
		expect(getFunctionInfo('vlookup')?.name).toBe('VLOOKUP');
		expect(getFunctionInfo('Norm.S.Dist')?.name).toBe('NORM.S.DIST');
		expect(getFunctionInfo('_xlfn.XLOOKUP')?.name).toBe('XLOOKUP');
		expect(getFunctionInfo('_xlfn._xlws.FILTER')?.name).toBe('FILTER');
		expect(getFunctionInfo('NOSUCHFUNCTION')).toBeUndefined();
		expect(getFunctionInfo('')).toBeUndefined();
	});

	it('lists functions sorted by name, optionally by category', () => {
		const names = listFunctions().map((f) => f.name);
		expect(names).toEqual([...names].sort());
		const cube = listFunctions('Cube').map((f) => f.name);
		expect(cube).toEqual([
			'CUBEKPIMEMBER',
			'CUBEMEMBER',
			'CUBEMEMBERPROPERTY',
			'CUBERANKEDMEMBER',
			'CUBESET',
			'CUBESETCOUNT',
			'CUBEVALUE',
		]);
		const total = FUNCTION_CATEGORIES.reduce((n, c) => n + listFunctions(c).length, 0);
		expect(total).toBe(names.length);
	});

	it('returns a fresh list that callers may change', () => {
		const first = listFunctions();
		first.length = 0;
		expect(listFunctions().length).toBeGreaterThan(0);
	});

	it('autocompletes by prefix in alphabetical order', () => {
		const names = searchFunctions('SU').map((f) => f.name);
		expect(names.slice(0, 6)).toEqual([
			'SUBSTITUTE',
			'SUBTOTAL',
			'SUM',
			'SUMIF',
			'SUMIFS',
			'SUMPRODUCT',
		]);
		expect(names.every((name) => name.startsWith('SU'))).toBe(true);
		expect(searchFunctions('norm.').map((f) => f.name)).toEqual([
			'NORM.DIST',
			'NORM.INV',
			'NORM.S.DIST',
			'NORM.S.INV',
		]);
		expect(searchFunctions('S', 3)).toHaveLength(3);
		expect(searchFunctions('')).toEqual([]);
		expect(searchFunctions('QQQ')).toEqual([]);
	});

	it('hands out frozen entries', () => {
		const sum = info('SUM');
		expect(Object.isFrozen(sum)).toBe(true);
		expect(Object.isFrozen(sum.args)).toBe(true);
		expect(Object.isFrozen(sum.args[0])).toBe(true);
	});
});
