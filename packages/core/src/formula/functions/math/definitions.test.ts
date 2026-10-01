import { describe, expect, it } from 'vitest';

import { filePrefixFor } from '../../prefixes';
import { FunctionRegistry } from '../../registry';
import { MATH_FUNCTIONS } from './index';

describe('math function definitions', () => {
	it.each(MATH_FUNCTIONS.map((def) => [def.name, def] as const))(
		'%s is well formed',
		(name, def) => {
			expect(name).toBe(name.toUpperCase());
			expect(def.minArgs).toBeGreaterThanOrEqual(0);
			expect(def.minArgs).toBeLessThanOrEqual(def.maxArgs);
			expect(def.maxArgs).toBeLessThanOrEqual(255);
			if (def.maxArgs > 0) {
				expect(def.params.length).toBeGreaterThan(0);
			}
			// The file prefix must agree with the shared prefix table used when saving.
			expect(def.filePrefix ?? '').toBe(filePrefixFor(name));
		},
	);

	it('registers every name once', () => {
		const names = MATH_FUNCTIONS.map((d) => d.name);
		expect(new Set(names).size).toBe(names.length);
		expect(new FunctionRegistry(MATH_FUNCTIONS).size).toBe(names.length);
	});
});
