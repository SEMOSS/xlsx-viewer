import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../sheet/workbook';
import { ERRORS } from '../values/errors';
import { forEachCellInRanges, sanitizeRange, boundingBox, anyRangeContains } from './range-cells';
import { duplicateKey, RangeStatistics } from './statistics';
import { rng, setCells, setColumn } from './test-utils/fixtures';

function stats(values: Parameters<typeof setColumn>[2], ranges = [rng('A1:A100')]) {
	const sheet = createWorkbook().sheets[0];
	setColumn(sheet, 'A1', values);
	return new RangeStatistics(sheet, ranges);
}

describe('RangeStatistics', () => {
	it('collects only numbers, sorted', () => {
		const s = stats([5, 'text', 1, true, null, ERRORS.NA, 3, '4']);
		expect([...s.values]).toEqual([1, 3, 5]);
		expect(s.count).toBe(3);
		expect(s.min).toBe(1);
		expect(s.max).toBe(5);
		expect(s.sum).toBe(9);
		expect(s.average).toBe(3);
	});

	it('handles an empty range', () => {
		const s = stats(['a', 'b']);
		expect(s.count).toBe(0);
		expect(s.min).toBeUndefined();
		expect(s.max).toBeUndefined();
		expect(s.average).toBeUndefined();
		expect(s.stdDev).toBeUndefined();
		expect(s.percentile(0.5)).toBeUndefined();
		expect(s.large(1)).toBeUndefined();
	});

	it('computes the sample standard deviation (STDEV.S)', () => {
		const s = stats([2, 4, 4, 4, 5, 5, 7, 9]);
		expect(s.stdDev).toBeCloseTo(2.13809, 5);
		expect(stats([42]).stdDev).toBe(0);
	});

	it('computes PERCENTILE.INC', () => {
		const s = stats([1, 2, 3, 4]);
		expect(s.percentile(0)).toBe(1);
		expect(s.percentile(1)).toBe(4);
		expect(s.percentile(0.5)).toBe(2.5);
		expect(s.percentile(0.25)).toBeCloseTo(1.75);
		expect(s.percentile(0.9)).toBeCloseTo(3.7);
		expect(s.percentile(-1)).toBe(1);
		expect(s.percentile(2)).toBe(4);
		expect(stats([7]).percentile(0.3)).toBe(7);
	});

	it('answers LARGE / SMALL with duplicates', () => {
		const s = stats([10, 10, 10, 5, 1]);
		expect(s.large(1)).toBe(10);
		expect(s.large(3)).toBe(10);
		expect(s.large(4)).toBe(5);
		expect(s.small(1)).toBe(1);
		expect(s.small(2)).toBe(5);
		expect(s.large(0)).toBeUndefined();
		expect(s.small(6)).toBeUndefined();
	});

	it('counts duplicates case-insensitively and by type', () => {
		const s = stats(['Apple', 'apple', 'APPLE', 1, '1', true, null, ERRORS.NA, ERRORS.NA]);
		expect(s.occurrences(duplicateKey('apple')!)).toBe(3);
		expect(s.occurrences(duplicateKey(1)!)).toBe(1);
		expect(s.occurrences(duplicateKey('1')!)).toBe(1);
		expect(s.occurrences(duplicateKey(true)!)).toBe(1);
		expect(s.occurrences(duplicateKey(false)!)).toBe(0);
	});

	it('scans the cell map for whole-column ranges', () => {
		const sheet = createWorkbook().sheets[0];
		setCells(sheet, { A1: 1, A500000: 2, B1: 100 });
		const s = new RangeStatistics(sheet, [rng('A:A')]);
		expect([...s.values]).toEqual([1, 2]);
	});

	it('counts overlapping cells once', () => {
		const sheet = createWorkbook().sheets[0];
		setColumn(sheet, 'A1', [1, 2, 3]);
		const s = new RangeStatistics(sheet, [rng('A1:A2'), rng('A2:A3')]);
		expect([...s.values]).toEqual([1, 2, 3]);
	});
});

describe('duplicateKey', () => {
	it('normalises numbers to 15 significant digits', () => {
		expect(duplicateKey(0.1 + 0.2)).toBe(duplicateKey(0.3));
		expect(duplicateKey(1)).not.toBe(duplicateKey('1'));
		expect(duplicateKey(null)).toBeUndefined();
		expect(duplicateKey(ERRORS.DIV0)).toBeUndefined();
		expect(duplicateKey(Number.NaN)).toBeUndefined();
	});
});

describe('range helpers', () => {
	it('sanitises ranges', () => {
		expect(sanitizeRange({ top: 5, left: 3, bottom: 1, right: 0 })).toEqual({
			top: 1,
			left: 0,
			bottom: 5,
			right: 3,
		});
		expect(sanitizeRange({ top: -2, left: 0, bottom: 2, right: 0 })).toEqual({
			top: 0,
			left: 0,
			bottom: 2,
			right: 0,
		});
		expect(sanitizeRange({ top: Number.NaN, left: 0, bottom: 2, right: 0 })).toBeUndefined();
	});

	it('computes bounding boxes and containment', () => {
		const ranges = [rng('B2:C3'), rng('E5')];
		expect(boundingBox(ranges)).toEqual(rng('B2:E5'));
		expect(boundingBox([])).toBeUndefined();
		expect(anyRangeContains(ranges, 4, 4)).toBe(true);
		expect(anyRangeContains(ranges, 3, 3)).toBe(false);
	});

	it('visits stored cells inside ranges with both strategies', () => {
		const sheet = createWorkbook().sheets[0];
		setCells(sheet, { A1: 1, B2: 2, C3: 3, Z99: 4 });
		const small: number[] = [];
		forEachCellInRanges(sheet, [rng('A1:B2')], (v) => small.push(v as number));
		expect(small.sort()).toEqual([1, 2]);
		const large: number[] = [];
		forEachCellInRanges(sheet, [rng('A1:C1000')], (v) => large.push(v as number));
		expect(large.sort()).toEqual([1, 2, 3]);
	});
});
