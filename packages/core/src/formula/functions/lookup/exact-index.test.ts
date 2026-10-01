import { describe, expect, it } from 'vitest';

import type { Scalar } from '../../types';
import { buildExactIndex, exactSearch } from './exact-index';
import type { Vec } from './grid';

function vec(values: Scalar[], indexed = true): Vec {
	const v: Vec = { length: values.length, dataLength: values.length, get: (i) => values[i] };
	if (indexed) {
		let built: ReturnType<typeof buildExactIndex> | undefined;
		v.exactIndex = () => (built ??= buildExactIndex(v));
	}
	return v;
}

const data: Scalar[] = Array.from({ length: 100 }, (_, i) => (i % 2 ? `Item ${i}` : i));
data[50] = 'Apple';
data[70] = 'APPLE';
data[80] = true;

describe('exactSearch', () => {
	it('finds first and last matches through the index', () => {
		const v = vec(data);
		expect(exactSearch(v, 'apple')).toBe(50);
		expect(exactSearch(v, 'apple', {}, true)).toBe(70);
		expect(exactSearch(v, 42)).toBe(42);
		expect(exactSearch(v, true)).toBe(80);
		expect(exactSearch(v, 'missing')).toBe(-1);
	});

	it('agrees with the linear scan, including wildcards and escapes', () => {
		const indexed = vec(data);
		const plain = vec(data, false);
		for (const x of ['Item 3', 'item 99', 'Item*', 'App?e', '~*', 7, 8, false, 'Apple']) {
			expect(exactSearch(indexed, x, { wildcards: true })).toBe(
				exactSearch(plain, x, { wildcards: true }),
			);
		}
	});
});
